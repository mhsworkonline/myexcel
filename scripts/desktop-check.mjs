// One-time verification of the packaged desktop app (Windows).
// Launches the release exe with WebView2 remote debugging and drives it over CDP.
// Usage: node scripts/desktop-check.mjs [path-to-exe]
import { chromium } from '@playwright/test';
import ExcelJS from 'exceljs';
import { execSync, spawn } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const exe = process.argv[2] || path.resolve('src-tauri/target/release/myexcel.exe');
const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'myexcel-desktop-'));
const file = path.join(dir, 'Budget.xlsx');
const wbx = new ExcelJS.Workbook();
const ws = wbx.addWorksheet('Budget');
ws.addRows([['Item', 'Cost'], ['Rent', 1200], ['Food', 450]]);
ws.getCell('B4').value = { formula: 'SUM(B2:B3)', result: 1650 };
await wbx.xlsx.writeFile(file);

const results = [];
const check = (name, ok, extra) => results.push(`${ok ? 'PASS' : 'FAIL'} ${name}${extra !== undefined ? ' :: ' + JSON.stringify(extra) : ''}`);

const port = 9333;
const child = spawn(exe, [file], { env: { ...process.env, WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS: `--remote-debugging-port=${port}` }, stdio: 'ignore' });
let browser;
for (let i = 0; i < 40 && !browser; i++) {
  await new Promise((r) => setTimeout(r, 500));
  try {
    browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`);
  } catch {
    /* not up yet */
  }
}
if (!browser) throw new Error('could not attach to the WebView2 instance');
const page = browser.contexts()[0].pages()[0];
const consoleErrors = [];
page.on('console', (m) => m.type() === 'error' && consoleErrors.push(m.text()));
page.on('pageerror', (e) => consoleErrors.push(String(e)));
await page.waitForFunction(() => window.__myexcel && window.__myexcel.S().file.name === 'Budget', null, { timeout: 20000 }).catch(() => undefined);

const info = await page.evaluate(async () => {
  const X = window.__myexcel;
  const st = X.S();
  return {
    name: st.file.name,
    handle: st.file.handle,
    cell: st.wb.activeSheet.getCell(1, 0)?.v,
    total: st.engine.getValue(st.wb.activeSheet, 3, 1),
    title: document.title,
    csp: document.querySelector('meta[http-equiv="Content-Security-Policy"]')?.getAttribute('content') ?? null,
    tauri: !!window.__TAURI_INTERNALS__,
  };
});
check('opened file passed at launch', info.name === 'Budget' && info.cell === 'Rent', info);
check('formula computed', info.total === 1650, info.total);
check('handle is the file path', typeof info.handle === 'string' && info.handle.toLowerCase().endsWith('budget.xlsx'), info.handle);

// A snapshot left by a previous crash (this script kills the app while dirty) must be offered on launch
const hadSnapshot = fs.existsSync(path.join(process.env.APPDATA, 'app.myexcel.desktop', 'recovery', 'snapshot.json'));
await page.waitForTimeout(800);
const offered = await page.evaluate(() => window.__myexcel.S().dialog?.type === 'recovery');
if (hadSnapshot) check('recovery offered after previous crash', offered);
if (offered) await page.click('[data-testid="dialog-cancel"]'); // Discard

// Edit and Ctrl+S → written straight back to disk, no dialog
await page.evaluate(() => {
  const X = window.__myexcel;
  const sh = X.S().wb.activeSheet;
  X.store.transact('t', (tx) => X.edit.writeInput(tx, sh, 2, 1, '500'));
});
await page.waitForTimeout(300);
const titleDirty = await page.evaluate(() => window.__myexcel.S().file.dirty);
check('edit marks dirty', titleDirty === true);
await page.keyboard.press('Control+s');
await page.waitForFunction(() => !window.__myexcel.S().file.dirty, null, { timeout: 15000 }).catch(() => undefined);
const back = new ExcelJS.Workbook();
await back.xlsx.readFile(file);
check('Ctrl+S saved in place', back.getWorksheet('Budget').getCell('B3').value === 500, back.getWorksheet('Budget').getCell('B3').value);

// settings + recovery files in the app data/config dirs
await page.evaluate(async () => {
  const X = window.__myexcel;
  const sh = X.S().wb.activeSheet;
  X.store.transact('t', (tx) => X.edit.writeInput(tx, sh, 5, 0, 'unsaved'));
});
const appDir = path.join(process.env.APPDATA, 'app.myexcel.desktop');
const settingsPath = path.join(appDir, 'settings.json');
const settings = fs.existsSync(settingsPath) ? JSON.parse(fs.readFileSync(settingsPath, 'utf8')) : null;
check('settings.json in app config dir', !!settings && Array.isArray(settings.recentFiles), settings && { theme: settings.theme, recent: settings.recentFiles.length });
check('recent files include launched file', !!settings && settings.recentFiles.some((r) => (r.path || '').toLowerCase() === file.toLowerCase()));
const recoveryPath = path.join(appDir, 'recovery', 'snapshot.json');
let recovered = false;
for (let i = 0; i < 40 && !recovered; i++) {
  await new Promise((r) => setTimeout(r, 1000));
  recovered = fs.existsSync(recoveryPath);
}
check('recovery snapshot written within 30s while dirty', recovered);
if (recovered) {
  const rec = JSON.parse(fs.readFileSync(recoveryPath, 'utf8'));
  check('recovery snapshot has the unsaved edit', JSON.stringify(rec.data).includes('unsaved'), rec.fileName);
}

// native clipboard: copy via the ribbon path and read with PowerShell
await page.evaluate(() => {
  const X = window.__myexcel;
  X.store.setState({ sel: { ranges: [{ r1: 0, c1: 0, r2: 1, c2: 1 }], active: { r: 0, c: 0 }, anchor: { r: 0, c: 0 } } });
});
await page.evaluate(() => document.querySelector('[data-testid="tab-Home"]')?.click());
await page.click('button[title="Copy (Ctrl+C)"]').catch(() => undefined);
await page.waitForTimeout(800);
const clipText = execSync('powershell -NoProfile -Command "Get-Clipboard -Raw"').toString();
const clipHtml = execSync('powershell -NoProfile -Command "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Clipboard]::GetText([System.Windows.Forms.TextDataFormat]::Html)"').toString();
check('native clipboard text', clipText.includes('Item\tCost'), clipText.slice(0, 40));
check('native clipboard CF_HTML (for Excel)', clipHtml.includes('StartFragment') && clipHtml.includes('<table'), clipHtml.slice(0, 60));

// native window title follows file name / status
const hwndTitle = execSync(`powershell -NoProfile -Command "(Get-Process -Id ${child.pid}).MainWindowTitle"`).toString().trim();
check('native window title shows file + status', /Budget .*MyExcel/.test(hwndTitle), hwndTitle);

check('no CSP violations or console errors', consoleErrors.length === 0, consoleErrors.slice(0, 5));
const net = await page.evaluate(async () => {
  try {
    await fetch('https://example.com/');
    return 'allowed';
  } catch (e) {
    return 'blocked';
  }
});
check('CSP blocks network requests', net === 'blocked', net);
// Close with unsaved changes: the native Save / Don't Save / Cancel prompt must keep the window open
execSync(`powershell -NoProfile -Command "(Get-Process -Id ${child.pid}).CloseMainWindow() | Out-Null"`);
await new Promise((r) => setTimeout(r, 2500));
let alive = true;
try {
  execSync(`powershell -NoProfile -Command "Get-Process -Id ${child.pid} -ErrorAction Stop | Out-Null"`, { stdio: 'ignore' });
} catch {
  alive = false;
}
const dialogShown = execSync(`powershell -NoProfile -ExecutionPolicy Bypass -File scripts/find-dialog.ps1 -ProcessId ${child.pid}`).toString().trim();
check('close with unsaved changes shows native prompt and keeps window open', alive && dialogShown === '1', { alive, dialogShown });

console.log(results.join('\n'));
await browser.close().catch(() => undefined);
child.kill();
try {
  execSync(`taskkill /F /PID ${child.pid}`, { stdio: 'ignore' });
} catch {
  /* already gone */
}
// leave no test traces in the real app data: drop the temp file from Recent and the crash snapshot
await new Promise((r) => setTimeout(r, 500));
const appData = path.join(process.env.APPDATA, 'app.myexcel.desktop');
const sp = path.join(appData, 'settings.json');
if (fs.existsSync(sp)) {
  const st = JSON.parse(fs.readFileSync(sp, 'utf8'));
  st.recentFiles = (st.recentFiles || []).filter((r) => !(r.path || '').toLowerCase().includes('myexcel-desktop-'));
  fs.writeFileSync(sp, JSON.stringify(st, null, 2));
}
fs.rmSync(path.join(appData, 'recovery', 'snapshot.json'), { force: true });

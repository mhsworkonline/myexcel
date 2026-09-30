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
// The test build (builds/test/...) has its own identity and app data folder.
const isTest = /test/i.test(path.basename(exe));
const appId = isTest ? 'app.myexcel.test' : 'app.myexcel.desktop';
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
// Leave any MyExcel window that is already open alone: keep its crash snapshots safe and use a
// separate WebView2 profile (otherwise the new window would join the running browser process).
const recoveryDir = path.join(process.env.APPDATA, appId, 'recovery');
const keep = new Map();
if (fs.existsSync(recoveryDir)) for (const f of fs.readdirSync(recoveryDir)) if (f.endsWith('.json')) keep.set(f, fs.readFileSync(path.join(recoveryDir, f)));
const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'myexcel-wv2-'));
const env = { ...process.env, WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS: `--remote-debugging-port=${port}`, WEBVIEW2_USER_DATA_FOLDER: profile };
const child = spawn(exe, [file], { env, stdio: 'ignore' });
const childrenOf = (pid) => execSync(`powershell -NoProfile -Command "Get-CimInstance Win32_Process -Filter 'ParentProcessId=${pid}' | Where-Object { $_.Name -like 'MyExcel*' } | ForEach-Object { $_.ProcessId.ToString() + ' ' + $_.CommandLine }"`).toString().trim().split(/\r?\n/).filter((l) => /MyExcel/i.test(l));
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

// A snapshot left by a crashed window may be offered on launch; dismiss it (restored at the end)
await page.waitForTimeout(800);
const offered = await page.evaluate(() => window.__myexcel.S().dialog?.type === 'recovery');
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
const appDir = path.join(process.env.APPDATA, appId);
const settingsPath = path.join(appDir, 'settings.json');
const settings = fs.existsSync(settingsPath) ? JSON.parse(fs.readFileSync(settingsPath, 'utf8')) : null;
check('settings.json in app config dir', !!settings && Array.isArray(settings.recentFiles), settings && { theme: settings.theme, recent: settings.recentFiles.length });
check('recent files include launched file', !!settings && settings.recentFiles.some((r) => (r.path || '').toLowerCase() === file.toLowerCase()));
const recoveryPath = path.join(appDir, 'recovery', `snapshot-${child.pid}.json`);
let recovered = false;
for (let i = 0; i < 40 && !recovered; i++) {
  await new Promise((r) => setTimeout(r, 1000));
  recovered = fs.existsSync(recoveryPath);
}
check('recovery snapshot written within 30s while dirty (one file per window)', recovered, path.basename(recoveryPath));
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
if (isTest) check('test build marked in title', /MyExcel \(Test build \d+\)/.test(hwndTitle), hwndTitle);

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
// Test build: Help > Reload reloads the UI from disk and keeps the open, unsaved workbook
if (isTest) {
  const before = await page.evaluate(() => ({ name: window.__myexcel.S().file.name, handle: window.__myexcel.S().file.handle }));
  await page.evaluate(() => window.__myexcel.file.snapshotForReload().then(() => location.reload()));
  await page.waitForTimeout(1500);
  await page.waitForFunction(() => window.__myexcel?.file && window.__myexcel.S().file.name === 'Budget', null, { timeout: 20000 }).catch(() => undefined);
  const after = await page.evaluate(() => {
    const st = window.__myexcel.S();
    return { name: st.file.name, handle: st.file.handle, dirty: st.file.dirty, a6: st.wb.activeSheet.getCell(5, 0)?.v, dialog: st.dialog?.type ?? null };
  });
  check('reload keeps the workbook, path and unsaved edit', after.name === before.name && after.handle === before.handle && after.dirty && after.a6 === 'unsaved' && !after.dialog, after);
}
// One workbook per window: opening a file while this one is in use starts a new window with it
const second = path.join(dir, 'Second.xlsx');
fs.copyFileSync(file, second);
await page.evaluate((p) => window.__myexcel.file.openPathInNewWindow(p), second);
await new Promise((r) => setTimeout(r, 4000));
const extraPids = [];
let kids = childrenOf(child.pid);
check('opening a file while one is open starts a new window with it', kids.some((l) => l.includes('Second.xlsx') && l.includes('--pos=')), kids);
const secondTitle = kids.length ? execSync(`powershell -NoProfile -Command "(Get-Process -Id ${kids[0].split(' ')[0]}).MainWindowTitle"`).toString().trim() : '';
check('the new window opened the file', /^Second/.test(secondTitle), secondTitle);
for (const k of kids) {
  const kpid = k.split(' ')[0];
  execSync(`taskkill /F /PID ${kpid}`, { stdio: 'ignore' });
  extraPids.push(kpid);
}
await page.keyboard.press('Control+n');
await new Promise((r) => setTimeout(r, 3000));
kids = childrenOf(child.pid);
check('Ctrl+N opens a new window with a blank workbook', kids.length === 1 && !kids[0].includes('.xlsx'), kids);
for (const k of kids) {
  const kpid = k.split(' ')[0];
  execSync(`taskkill /F /PID ${kpid}`, { stdio: 'ignore' });
  extraPids.push(kpid);
}
check('this window still has its workbook', (await page.evaluate(() => window.__myexcel.S().file.name)) === 'Budget');

// Ctrl+W with unsaved changes: the native Save / Don't Save / Cancel prompt must keep the window open
await page.evaluate(() => document.querySelector('[data-testid="grid-canvas"]')?.focus());
await page.keyboard.press('Control+w');
await new Promise((r) => setTimeout(r, 2500));
let alive = true;
try {
  execSync(`powershell -NoProfile -Command "Get-Process -Id ${child.pid} -ErrorAction Stop | Out-Null"`, { stdio: 'ignore' });
} catch {
  alive = false;
}
const dialogShown = execSync(`powershell -NoProfile -ExecutionPolicy Bypass -File scripts/find-dialog.ps1 -ProcessId ${child.pid}`).toString().trim();
check('Ctrl+W with unsaved changes shows the native prompt and keeps the window open', alive && dialogShown === '1', { alive, dialogShown });

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
const appData = path.join(process.env.APPDATA, appId);
const sp = path.join(appData, 'settings.json');
if (fs.existsSync(sp)) {
  const st = JSON.parse(fs.readFileSync(sp, 'utf8'));
  st.recentFiles = (st.recentFiles || []).filter((r) => !(r.path || '').toLowerCase().includes('myexcel-desktop-'));
  fs.writeFileSync(sp, JSON.stringify(st, null, 2));
}
// remove this run's snapshot/lock and bring back snapshots that belonged to other windows
for (const f of [child.pid, ...extraPids].flatMap((id) => [`snapshot-${id}.json`, `owner-${id}.lock`])) fs.rmSync(path.join(recoveryDir, f), { force: true });
for (const [f, data] of keep) if (!fs.existsSync(path.join(recoveryDir, f))) fs.writeFileSync(path.join(recoveryDir, f), data);
fs.rmSync(profile, { recursive: true, force: true });

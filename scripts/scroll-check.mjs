// One-time performance check: open a 200k-row workbook and measure scroll frame times.
// Usage: node scripts/scroll-check.mjs   (server must be running on :3100)
import { chromium } from '@playwright/test';
import ExcelJS from 'exceljs';

const ROWS = 200000;
const wbx = new ExcelJS.Workbook();
const ws = wbx.addWorksheet('Big');
ws.addRow(['ID', 'Name', 'Region', 'Amount', 'Date', 'Score']);
for (let i = 1; i <= ROWS; i++) ws.addRow([i, 'Customer ' + i, ['East', 'West', 'North', 'South'][i % 4], Math.round(Math.random() * 100000) / 100, new Date(Date.UTC(2024, i % 12, (i % 28) + 1)), i % 97]);
const t0 = Date.now();
const buf = Buffer.from(await wbx.xlsx.writeBuffer());
console.log(`generated ${ROWS} rows (${(buf.length / 1e6).toFixed(1)} MB) in ${Date.now() - t0} ms`);

const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1440, height: 860 } });
await page.goto('http://localhost:3100/');
await page.waitForFunction(() => window.__myexcel);
const openMs = await page.evaluate(async (b64) => {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const t = performance.now();
  await window.__myexcel.openFromData('big.xlsx', bytes.buffer);
  return performance.now() - t;
}, buf.toString('base64'));
console.log(`opened in ${Math.round(openMs)} ms, rows in model: ${await page.evaluate(() => window.__myexcel.S().wb.activeSheet.rows.size)}`);
await page.waitForTimeout(500);
await page.mouse.move(700, 500);
// sample frame intervals while wheel-scrolling continuously
await page.evaluate(() => {
  window.__frames = [];
  let last = performance.now();
  const tick = (t) => { window.__frames.push(t - last); last = t; if (window.__frames.length < 400) requestAnimationFrame(tick); };
  requestAnimationFrame(tick);
});
for (let i = 0; i < 300; i++) {
  await page.mouse.wheel(0, i < 150 ? 2000 : -1200);
}
await page.keyboard.press('Control+End');
await page.waitForTimeout(300);
const frames = await page.evaluate(() => window.__frames.slice(5));
frames.sort((a, b) => a - b);
const p = (q) => frames[Math.floor(frames.length * q)].toFixed(1);
console.log(`frame intervals over ${frames.length} frames: median ${p(0.5)} ms, p95 ${p(0.95)} ms, max ${frames[frames.length - 1].toFixed(1)} ms`);
console.log('active cell after Ctrl+End:', await page.inputValue('[data-testid="name-box"]'));
await page.screenshot({ path: 'screenshots/big.png' });
await browser.close();

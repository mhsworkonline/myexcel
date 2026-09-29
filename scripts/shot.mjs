// Visual check helper: node scripts/shot.mjs <name> [actions-json]
import { chromium } from '@playwright/test';
const name = process.argv[2] || 'shot';
const actions = process.argv[3] ? JSON.parse(process.argv[3]) : [];
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1440, height: 860 }, deviceScaleFactor: 1 });
const errors = [];
page.on('pageerror', (e) => errors.push(String(e)));
page.on('console', (m) => { if (m.type() === 'error') errors.push(m.text()); });
await page.goto('http://localhost:3100/');
await page.waitForSelector('[data-testid="grid-canvas"]');
await page.waitForTimeout(500);
for (const a of actions) {
  if (a.click) await page.click(a.click);
  if (a.clickAt) await page.mouse.click(a.clickAt[0], a.clickAt[1]);
  if (a.dblAt) await page.mouse.dblclick(a.dblAt[0], a.dblAt[1]);
  if (a.type) await page.keyboard.type(a.type);
  if (a.press) await page.keyboard.press(a.press);
  if (a.wait) await page.waitForTimeout(a.wait);
  if (a.eval) await page.evaluate(a.eval);
}
await page.waitForTimeout(300);
await page.screenshot({ path: `screenshots/${name}.png` });
if (errors.length) console.log('ERRORS:\n' + errors.join('\n'));
await browser.close();

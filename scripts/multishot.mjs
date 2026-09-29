// Scenario screenshots for visual review
import { chromium } from '@playwright/test';
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1440, height: 860 } });
const errors = [];
page.on('pageerror', (e) => errors.push(String(e.stack || e)));
await page.goto('http://localhost:3100/');
await page.waitForSelector('[data-testid="grid-canvas"]');
await page.waitForTimeout(300);
const cell = (c, r) => [25 + 64 * c + 30, 200 + 20 + 20 * r + 10];
const click = async (c, r, opts) => page.mouse.click(...cell(c, r), opts);
// data
await click(0, 0);
const rows = [['Region', 'Rep', 'Sales', 'Date'], ['East', 'Ann', '1200', '1/5/2024'], ['West', 'Bob', '950', '2/9/2024'], ['East', 'Cy', '1810', '3/1/2024'], ['North', 'Dee', '400', '3/15/2024'], ['West', 'Eve', '1333', '4/2/2024']];
for (const r of rows) { for (let i = 0; i < r.length; i++) { await page.keyboard.type(r[i]); await page.keyboard.press(i === r.length - 1 ? 'Enter' : 'Tab'); } }
await click(0, 0);
await page.keyboard.press('Control+Shift+L');
await page.waitForTimeout(200);
// open filter dropdown on column C
const [fx, fy] = cell(2, 0);
await page.mouse.click(fx + 24, fy + 1);
await page.waitForTimeout(300);
await page.screenshot({ path: 'screenshots/filter.png' });
await page.keyboard.press('Escape');
await page.mouse.click(700, 600);
await page.waitForTimeout(100);
// context menu
await page.mouse.click(...cell(1, 2), { button: 'right' });
await page.waitForTimeout(300);
await page.screenshot({ path: 'screenshots/context.png' });
await page.keyboard.press('Escape');
// fill handle: select C2 and drag handle down
await click(4, 1); await page.keyboard.type('Q1'); await page.keyboard.press('Enter');
await click(4, 1);
const [hx, hy] = cell(4, 1);
await page.mouse.move(hx + 32, hy + 10);
await page.mouse.down();
await page.mouse.move(hx + 32, hy + 60, { steps: 5 });
await page.mouse.up();
await page.waitForTimeout(200);
// freeze top row + scroll
await click(0, 1);
await page.click('[data-testid="tab-View"]');
await page.click('[data-testid="btn-freeze"]');
await page.waitForTimeout(200);
await page.screenshot({ path: 'screenshots/viewtab.png' });
await page.getByText('Freeze Top Row').click();
await page.mouse.move(700, 500);
await page.mouse.wheel(0, 300);
await page.waitForTimeout(300);
await page.screenshot({ path: 'screenshots/frozen.png' });
// dark mode
await page.click('[data-testid="theme-toggle"]');
await page.click('[data-testid="tab-Home"]');
await page.waitForTimeout(300);
await page.screenshot({ path: 'screenshots/dark.png' });
if (errors.length) console.log('ERRORS\n' + errors.join('\n'));
await browser.close();

import { expect, test } from '@playwright/test';
import ExcelJS from 'exceljs';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

type Cell = { v?: unknown; f?: string; s?: number } | undefined;

async function cellAt(page: import('@playwright/test').Page, r: number, c: number): Promise<Cell> {
  return page.evaluate(([r, c]) => {
    const w = window as unknown as { __myexcel: { S: () => { wb: { activeSheet: { getCell: (r: number, c: number) => Cell } } } } };
    return w.__myexcel.S().wb.activeSheet.getCell(r, c);
  }, [r, c] as const);
}

async function isBold(page: import('@playwright/test').Page, r: number, c: number): Promise<boolean> {
  return page.evaluate(([r, c]) => {
    const w = window as unknown as { __myexcel: { S: () => { wb: { activeSheet: { styleIdAt: (r: number, c: number) => number }; styles: { get: (id: number) => { bold?: boolean } } } } } };
    const st = w.__myexcel.S();
    return !!st.wb.styles.get(st.wb.activeSheet.styleIdAt(r, c)).bold;
  }, [r, c] as const);
}

test('open, edit, format, sort, undo, save', async ({ page }) => {
  // Build a fixture workbook
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'myexcel-'));
  const src = path.join(dir, 'fruit.xlsx');
  const wbx = new ExcelJS.Workbook();
  const ws = wbx.addWorksheet('Fruit');
  ws.addRows([
    ['Name', 'Qty'],
    ['Pear', 3],
    ['Apple', 7],
    ['Mango', 5],
  ]);
  ws.getCell('D1').value = { formula: 'SUM(B2:B4)', result: 15 };
  await wbx.xlsx.writeFile(src);

  // Force the <input type=file> / download fallbacks so Playwright can drive them
  await page.addInitScript(() => {
    const w = window as unknown as Record<string, unknown>;
    delete w.showOpenFilePicker;
    delete w.showSaveFilePicker;
  });
  await page.goto('/');
  await page.waitForSelector('[data-testid="grid-canvas"]');

  // Open via File > Open > Browse
  await page.getByTestId('tab-File').click();
  await page.getByTestId('bs-open').click();
  const chooserPromise = page.waitForEvent('filechooser');
  await page.getByTestId('bs-browse').click();
  const chooser = await chooserPromise;
  await chooser.setFiles(src);
  await expect(page.getByTestId('sheet-tab-Fruit')).toBeVisible();
  expect((await cellAt(page, 1, 0))?.v).toBe('Pear');
  expect((await cellAt(page, 0, 3))?.f).toBe('=SUM(B2:B4)');

  // Edit: B2 = 10, formula recalculates
  await page.getByTestId('name-box').fill('B2');
  await page.getByTestId('name-box').press('Enter');
  await page.keyboard.type('10');
  await page.keyboard.press('Enter');
  expect((await cellAt(page, 1, 1))?.v).toBe(10);
  const total = await page.evaluate(() => {
    const w = window as unknown as { __myexcel: { S: () => { engine: { getValue: (s: unknown, r: number, c: number) => unknown }; wb: { activeSheet: unknown } } } };
    const st = w.__myexcel.S();
    return st.engine.getValue(st.wb.activeSheet, 0, 3);
  });
  expect(total).toBe(22);

  // Format: bold the header row via ribbon
  await page.getByTestId('name-box').fill('A1:B1');
  await page.getByTestId('name-box').press('Enter');
  await page.getByTestId('btn-bold').click();
  expect(await isBold(page, 0, 0)).toBe(true);
  expect(await isBold(page, 0, 1)).toBe(true);

  // Sort A→Z on the Name column
  await page.getByTestId('name-box').fill('A2');
  await page.getByTestId('name-box').press('Enter');
  await page.getByTestId('btn-sortfilter').click();
  await page.getByTestId('sort-asc').click();
  expect((await cellAt(page, 1, 0))?.v).toBe('Apple');
  expect((await cellAt(page, 2, 0))?.v).toBe('Mango');
  expect((await cellAt(page, 3, 0))?.v).toBe('Pear');

  // Undo the sort
  await page.getByTestId('qat-undo').click();
  expect((await cellAt(page, 1, 0))?.v).toBe('Pear');
  expect((await cellAt(page, 1, 1))?.v).toBe(10);

  // Save (Ctrl+S → Save As download fallback) and verify the file
  const downloadPromise = page.waitForEvent('download');
  await page.keyboard.press('Control+s');
  const download = await downloadPromise;
  const out = path.join(dir, 'saved.xlsx');
  await download.saveAs(out);
  const check = new ExcelJS.Workbook();
  await check.xlsx.readFile(out);
  const s = check.getWorksheet('Fruit')!;
  expect(s.getCell('A2').value).toBe('Pear');
  expect(s.getCell('B2').value).toBe(10);
  expect(s.getCell('A1').font?.bold).toBe(true);
  expect((s.getCell('D1').value as { formula: string }).formula).toBe('SUM(B2:B4)');
  await expect(page.getByTestId('doc-title')).toContainText('Saved');
});

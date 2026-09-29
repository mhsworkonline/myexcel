import { describe, expect, it } from 'vitest';
import { Engine, isErrorVal } from '@/engine/Engine';
import { readCsv, writeCsv } from '@/io/csv';
import { readXlsx, writeXlsx } from '@/io/xlsx';
import { formatValue } from '@/model/numfmt';
import { Sheet } from '@/model/sheet';
import { Workbook } from '@/model/workbook';

function buildWorkbook(): { wb: Workbook; engine: Engine } {
  const wb = Workbook.createDefault();
  const s = wb.sheets[0];
  s.name = 'Sales Data';
  const bold = wb.styles.intern({ bold: true, fillColor: '#FFFF00', hAlign: 'center', bBottom: { style: 'double', color: '#FF0000' } });
  const money = wb.styles.intern({ numFmt: '[$₹-4009] #,##0.00', fontColor: '#1F4E79', italic: true, fontSize: 14, fontName: 'Arial' });
  const date = wb.styles.intern({ numFmt: 'm/d/yyyy', wrap: true, vAlign: 'top' });
  s.setCellRaw(0, 0, { v: 'Item', s: bold });
  s.setCellRaw(0, 1, { v: 'Amount', s: bold });
  s.setCellRaw(0, 2, { v: 'Date', s: bold });
  s.setCellRaw(1, 0, { v: 'Apples' });
  s.setCellRaw(1, 1, { v: 1234567.5, s: money });
  s.setCellRaw(1, 2, { v: 45293, s: date });
  s.setCellRaw(2, 0, { v: 'Pears' });
  s.setCellRaw(2, 1, { v: 42, s: money });
  s.setCellRaw(2, 2, { v: true });
  s.setCellRaw(3, 0, { v: 'Total' });
  s.setCellRaw(3, 1, { f: '=SUM(B2:B3)' });
  s.setCellRaw(4, 1, { f: "=Other!A1*2" });
  s.setCellRaw(5, 0, { v: '#N/A', e: true });
  s.colWidths.set(0, 120);
  s.colWidths.set(3, 30);
  s.rowHeights.set(0, 30);
  s.hiddenRows.add(8);
  s.freeze = { rows: 1, cols: 1 };
  s.merges = [{ r1: 6, c1: 0, r2: 7, c2: 2 }];
  s.tabColor = '#FF0000';
  const other = new Sheet('Other');
  other.setCellRaw(0, 0, { v: 21 });
  wb.sheets.push(other);
  const engine = new Engine(wb);
  engine.attach();
  return { wb, engine };
}

describe('xlsx round-trip', () => {
  it('preserves values, formulas, styles, sizes, freezes, merges and sheet names', async () => {
    const { wb, engine } = buildWorkbook();
    const buf = await writeXlsx(wb, (sh, r, c) => {
      const v = engine.getValue(sh, r, c);
      return isErrorVal(v) ? { error: v.error } : v;
    });
    const back = await readXlsx(buf);
    expect(back.sheets.map((s) => s.name)).toEqual(['Sales Data', 'Other']);
    const s = back.sheets[0];
    expect(s.getCell(0, 0)?.v).toBe('Item');
    expect(s.getCell(1, 1)?.v).toBe(1234567.5);
    expect(s.getCell(1, 2)?.v).toBe(45293);
    expect(s.getCell(2, 2)?.v).toBe(true);
    expect(s.getCell(3, 1)?.f).toBe('=SUM(B2:B3)');
    expect(s.getCell(4, 1)?.f).toBe('=Other!A1*2');
    expect(s.getCell(5, 0)).toMatchObject({ v: '#N/A', e: true });
    // styles
    const hdr = back.styles.get(s.getCell(0, 0)?.s);
    expect(hdr).toMatchObject({ bold: true, fillColor: '#FFFF00', hAlign: 'center' });
    expect(hdr.bBottom).toEqual({ style: 'double', color: '#FF0000' });
    const money = back.styles.get(s.getCell(1, 1)?.s);
    expect(money).toMatchObject({ numFmt: '[$₹-4009] #,##0.00', fontColor: '#1F4E79', italic: true, fontSize: 14, fontName: 'Arial' });
    expect(formatValue(1234567.5, money.numFmt).text).toBe('₹ 12,34,567.50');
    expect(back.styles.get(s.getCell(1, 2)?.s)).toMatchObject({ numFmt: 'm/d/yyyy', wrap: true, vAlign: 'top' });
    // layout
    expect(s.colWidths.get(0)).toBe(120);
    expect(s.colWidths.get(3)).toBe(30);
    expect(s.rowHeights.get(0)).toBe(30);
    expect(s.hiddenRows.has(8)).toBe(true);
    expect(s.freeze).toEqual({ rows: 1, cols: 1 });
    expect(s.merges).toEqual([{ r1: 6, c1: 0, r2: 7, c2: 2 }]);
    expect(s.tabColor).toBe('#FF0000');
    // formulas still compute after reload
    const eng2 = new Engine(back);
    eng2.attach();
    expect(eng2.getValue(s, 3, 1)).toBe(1234609.5);
    expect(eng2.getValue(s, 4, 1)).toBe(42);
  });
});

describe('csv round-trip', () => {
  it('preserves text, numbers, quoting and parses types on open', () => {
    const csv = 'Name,Qty,Price,Note\r\n"Smith, J",3,$4.50,"said ""hi"""\r\nLee,10,12.25,\r\n';
    const wb = readCsv(csv, 'orders.csv');
    const s = wb.sheets[0];
    expect(s.name).toBe('orders');
    expect(s.getCell(1, 0)?.v).toBe('Smith, J');
    expect(s.getCell(1, 1)?.v).toBe(3);
    expect(s.getCell(1, 2)?.v).toBe(4.5);
    expect(s.getCell(1, 3)?.v).toBe('said "hi"');
    const out = writeCsv(s, (r, c) => {
      const cell = s.getCell(r, c);
      if (!cell || cell.v === undefined || cell.v === null) return '';
      return formatValue(cell.v, wb.styles.get(cell.s).numFmt).text;
    });
    expect(out).toBe('Name,Qty,Price,Note\r\n"Smith, J",3,$4.50,"said ""hi"""\r\nLee,10,12.25,\r\n');
  });
});

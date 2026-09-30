// Legacy .xls (BIFF) import via SheetJS CE — read only.
import * as XLSX from 'xlsx';
import { parseRange } from '../model/address';
import { Sheet } from '../model/sheet';
import type { Cell } from '../model/types';
import { Workbook } from '../model/workbook';

export async function readXls(data: ArrayBuffer): Promise<Workbook> {
  const x = XLSX.read(new Uint8Array(data), { type: 'array', cellFormula: true, cellNF: true, cellStyles: true, cellDates: false });
  const wb = new Workbook();
  for (const name of x.SheetNames) {
    const ws = x.Sheets[name];
    const sheet = new Sheet(name.slice(0, 31));
    const fmtIds = new Map<string, number>();
    for (const key of Object.keys(ws)) {
      if (key[0] === '!') continue;
      const rg = parseRange(key);
      if (!rg) continue;
      const c = ws[key] as XLSX.CellObject;
      const cell: Cell = {};
      if (c.f) cell.f = '=' + c.f;
      if (c.t === 'e') {
        cell.v = c.w ?? '#N/A';
        cell.e = true;
      } else if (c.t === 'n' || c.t === 's' || c.t === 'b') cell.v = c.v as number | string | boolean;
      else if (c.t === 'd' && c.v instanceof Date) cell.v = (c.v.getTime() - Date.UTC(1899, 11, 30)) / 86400000;
      const z = typeof c.z === 'string' ? c.z : undefined;
      if (z && z !== 'General') {
        let id = fmtIds.get(z);
        if (id === undefined) fmtIds.set(z, (id = wb.styles.intern({ numFmt: z })));
        cell.s = id;
      }
      if (c.l?.Target) cell.link = c.l.Target;
      if (c.c?.length) cell.note = { text: c.c.map((n) => n.t).join('\n'), author: c.c[0].a };
      sheet.setCellRaw(rg.r1, rg.c1, cell);
    }
    for (const m of ws['!merges'] ?? []) sheet.merges.push({ r1: m.s.r, c1: m.s.c, r2: m.e.r, c2: m.e.c });
    (ws['!cols'] ?? []).forEach((col, i) => {
      if (!col) return;
      if (col.hidden) sheet.hiddenCols.add(i);
      const px = col.wpx ?? (col.wch ? Math.round(col.wch * 7 + 5) : undefined);
      if (px) sheet.colWidths.set(i, Math.round(px));
    });
    (ws['!rows'] ?? []).forEach((row, i) => {
      if (!row) return;
      if (row.hidden) sheet.hiddenRows.add(i);
      const px = row.hpx ?? (row.hpt ? Math.round((row.hpt * 4) / 3) : undefined);
      if (px) sheet.rowHeights.set(i, Math.round(px));
    });
    sheet.touch();
    wb.sheets.push(sheet);
  }
  if (!wb.sheets.length) wb.sheets.push(new Sheet('Sheet1'));
  wb.activeSheetId = wb.sheets[0].id;
  for (const n of x.Workbook?.Names ?? []) {
    if (!n.Name || n.Name.startsWith('_')) continue;
    wb.names.push({ name: n.Name, ref: n.Ref, scope: n.Sheet !== undefined ? wb.sheets[n.Sheet]?.id : undefined });
  }
  return wb;
}

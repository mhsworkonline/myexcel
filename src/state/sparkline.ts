import { parseRange, Range } from '../model/address';
import type { Sheet } from '../model/sheet';
import type { CellValue } from '../model/types';
import type { Workbook } from '../model/workbook';
import { getScalar } from './values';

/** Resolve "Sheet!A1:B2" / "A1:B2" relative to `host`. */
export function resolveRef(wb: Workbook, host: Sheet, ref: string): { sheet: Sheet; range: Range } | null {
  let t = ref.trim().replace(/^=/, '');
  let sheet = host;
  const bang = t.lastIndexOf('!');
  if (bang > 0) {
    const name = t.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'");
    const s = wb.sheetByName(name);
    if (!s) return null;
    sheet = s;
    t = t.slice(bang + 1);
  }
  const rg = parseRange(t.replace(/\$/g, ''));
  return rg ? { sheet, range: rg } : null;
}

export function rangeValues(sheet: Sheet, rg: Range): CellValue[] {
  const out: CellValue[] = [];
  const r2 = Math.min(rg.r2, rg.r1 + 100000);
  const c2 = Math.min(rg.c2, rg.c1 + 1000);
  for (let r = rg.r1; r <= r2; r++) for (let c = rg.c1; c <= c2; c++) out.push(getScalar(sheet, r, c));
  return out;
}

export function sparklineValues(wb: Workbook, host: Sheet, ref: string): CellValue[] {
  const res = resolveRef(wb, host, ref);
  if (!res) return [];
  return rangeValues(res.sheet, res.range);
}

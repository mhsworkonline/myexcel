// Read path for cell values and display text. UI → here → engine/model.
import { ComputedValue, isErrorVal } from '../engine/Engine';
import { formatGeneral, formatValue, FormatResult } from '../model/numfmt';
import type { Sheet } from '../model/sheet';
import type { CellStyle } from '../model/styles';
import type { CellValue } from '../model/types';
import { S } from './store';

export function getComputed(sheet: Sheet, r: number, c: number): ComputedValue {
  const cell = sheet.getCell(r, c);
  if (!cell) return null;
  if (cell.f !== undefined) {
    const eng = S().engine;
    if (eng.attached) return eng.getValue(sheet, r, c);
    if (cell.e && typeof cell.v === 'string') return { error: cell.v };
    return cell.v ?? null;
  }
  if (cell.e && typeof cell.v === 'string') return { error: cell.v };
  return cell.v ?? null;
}

/** Plain scalar (errors → error text) for sorting/filtering/charts. */
export function getScalar(sheet: Sheet, r: number, c: number): CellValue {
  const v = getComputed(sheet, r, c);
  if (isErrorVal(v)) return v.circular ? 0 : v.error;
  return v;
}

export function getNumber(sheet: Sheet, r: number, c: number): number | null {
  const v = getComputed(sheet, r, c);
  return typeof v === 'number' ? v : null;
}

export interface Display extends FormatResult {
  hAlign: 'left' | 'center' | 'right';
  isErr: boolean;
  value: ComputedValue;
}

export function styleAt(sheet: Sheet, r: number, c: number): CellStyle {
  return S().wb.styles.get(sheet.styleIdAt(r, c));
}

export function displayOf(sheet: Sheet, r: number, c: number, style: CellStyle, generalChars = 11, showFormulas = false): Display {
  const cell = sheet.getCell(r, c);
  if (showFormulas && cell?.f) return { text: cell.f, hAlign: 'left', isErr: false, value: null };
  let value = getComputed(sheet, r, c);
  if (isErrorVal(value)) {
    if (value.circular) value = 0;
    else return { text: value.error, hAlign: 'center', isErr: true, value };
  }
  let res: FormatResult;
  if (typeof value === 'number' && style.numFmt === undefined) res = { text: formatGeneral(value, generalChars), isNumber: true };
  else res = formatValue(value, style.numFmt, generalChars);
  let hAlign: Display['hAlign'] = 'left';
  const h = style.hAlign;
  if (h === 'center' || h === 'centerContinuous') hAlign = 'center';
  else if (h === 'right') hAlign = 'right';
  else if (h === 'left' || h === 'fill' || h === 'justify' || h === 'distributed') hAlign = 'left';
  else if (typeof value === 'number') hAlign = 'right';
  else if (typeof value === 'boolean') hAlign = 'center';
  return { ...res, hAlign, isErr: false, value };
}

/** Text shown in the formula bar / edit box for a cell. */
export function editText(sheet: Sheet, r: number, c: number): string {
  const cell = sheet.getCell(r, c);
  if (!cell) return '';
  if (cell.f !== undefined) return cell.f;
  const v = cell.v;
  if (v === undefined || v === null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'string') {
    // Numeric-looking text needs the apostrophe to stay text
    const st = S().wb.styles.get(sheet.styleIdAt(r, c));
    if (st.numFmt !== '@' && !cell.e && /^[+-]?(\d|\.\d)|^(true|false)$|^=/i.test(v) && !isNaN(Number(v.replace(/[,$%]/g, '')))) return "'" + v;
    return v;
  }
  const st = S().wb.styles.get(sheet.styleIdAt(r, c));
  const fmt = st.numFmt;
  if (fmt && /[dmyhs]/i.test(fmt.replace(/"[^"]*"|\[[^\]]*\]/g, '')) && !/^[#0.,%E+\-?/ ]+$/i.test(fmt)) {
    // Dates/times edit as their date text
    const hasDate = /[dy]/i.test(fmt.replace(/"[^"]*"|\[[^\]]*\]/g, '')) || /m{3,}/i.test(fmt);
    const hasTime = /[hs]/i.test(fmt.replace(/"[^"]*"|\[[^\]]*\]/g, ''));
    const f = hasDate && hasTime ? 'm/d/yyyy h:mm:ss AM/PM' : hasDate ? 'm/d/yyyy' : 'h:mm:ss AM/PM';
    return formatValue(v, f).text;
  }
  if (fmt && fmt.includes('%')) {
    const pct = v * 100;
    return String(+pct.toPrecision(15)) + '%';
  }
  return String(+v.toPrecision(15));
}

import { isErrorVal } from '../engine/Engine';
import { shiftFormula } from '../model/formula';
import { formatValue } from '../model/numfmt';
import { parseInput } from '../model/parseInput';
import type { Sheet } from '../model/sheet';
import type { DataValidation } from '../model/types';
import { S } from './store';

export function validationAt(sheet: Sheet, r: number, c: number): DataValidation | undefined {
  for (const dv of sheet.validations) for (const rg of dv.ranges) if (r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2) return dv;
  return undefined;
}

function topLeft(dv: DataValidation): { r: number; c: number } {
  let r = Infinity;
  let c = Infinity;
  for (const rg of dv.ranges) {
    r = Math.min(r, rg.r1);
    c = Math.min(c, rg.c1);
  }
  return { r, c };
}

function evalOperand(sheet: Sheet, dv: DataValidation, f: string | undefined, r: number, c: number): number | string | null {
  if (f === undefined || f === '') return null;
  const t = f.trim();
  if (!t.startsWith('=')) {
    const p = parseInput(t);
    if (p.kind === 'number') return p.value as number;
    return t;
  }
  const tl = topLeft(dv);
  const shifted = shiftFormula(t, r - tl.r, c - tl.c);
  const v = S().engine.evaluate(shifted, sheet);
  if (isErrorVal(v)) return null;
  if (typeof v === 'boolean') return v ? 1 : 0;
  return v;
}

export function listItems(sheet: Sheet, dv: DataValidation): string[] {
  const f = (dv.formula1 ?? '').trim();
  if (!f) return [];
  if (!f.startsWith('=')) {
    const raw = f.replace(/^"|"$/g, '');
    return raw.split(',').map((s) => s.trim()).filter((s) => s !== '');
  }
  const arr = S().engine.evaluateArray(f, sheet);
  const out: string[] = [];
  for (const row of arr)
    for (const v of row) {
      if (v === null || v === '' || isErrorVal(v)) continue;
      out.push(typeof v === 'number' ? formatValue(v, undefined).text : String(v));
    }
  return out;
}

function compare(op: string | undefined, v: number, a: number | null, b: number | null): boolean {
  if (a === null) return true;
  switch (op ?? 'between') {
    case 'between': return b === null ? v >= a : v >= Math.min(a, b) && v <= Math.max(a, b);
    case 'notBetween': return b === null ? v < a : v < Math.min(a, b) || v > Math.max(a, b);
    case 'equal': return v === a;
    case 'notEqual': return v !== a;
    case 'greaterThan': return v > a;
    case 'lessThan': return v < a;
    case 'greaterThanOrEqual': return v >= a;
    case 'lessThanOrEqual': return v <= a;
  }
  return true;
}

export interface ValidationFailure {
  style: 'stop' | 'warning' | 'information';
  title: string;
  message: string;
}

/** Returns true when `text` is acceptable for the cell; otherwise the configured error. */
export function checkValidation(sheet: Sheet, r: number, c: number, text: string): true | ValidationFailure {
  const dv = validationAt(sheet, r, c);
  if (!dv || dv.type === 'any') return true;
  if (text === '') return dv.allowBlank ? true : fail(dv);
  if (!validValue(sheet, dv, r, c, text)) return dv.showError ? fail(dv) : true;
  return true;
}

export function validValue(sheet: Sheet, dv: DataValidation, r: number, c: number, text: string): boolean {
  const p = text.startsWith('=') ? null : parseInput(text);
  const num = p && p.kind === 'number' ? (p.value as number) : null;
  const a = () => evalOperand(sheet, dv, dv.formula1, r, c);
  const b = () => evalOperand(sheet, dv, dv.formula2, r, c);
  const asNum = (x: number | string | null) => (typeof x === 'number' ? x : x === null ? null : isNaN(Number(x)) ? null : Number(x));
  switch (dv.type) {
    case 'whole':
      if (num === null || !Number.isInteger(num)) return false;
      return compare(dv.operator, num, asNum(a()), asNum(b()));
    case 'decimal':
    case 'date':
    case 'time':
      if (num === null) return false;
      return compare(dv.operator, num, asNum(a()), asNum(b()));
    case 'textLength':
      return compare(dv.operator, text.length, asNum(a()), asNum(b()));
    case 'list': {
      const items = listItems(sheet, dv);
      return items.some((i) => i.toLowerCase() === text.toLowerCase());
    }
    case 'custom': {
      const v = a();
      return v === 1 || v === 'TRUE' || (typeof v === 'number' && v !== 0);
    }
  }
  return true;
}

function fail(dv: DataValidation): ValidationFailure {
  return {
    style: dv.errorStyle,
    title: dv.errorTitle || 'MyExcel',
    message: dv.error || 'This value doesn\'t match the data validation restrictions defined for this cell.',
  };
}

/** Cells in the sheet currently violating validation (for Circle Invalid Data). */
export function invalidCells(sheet: Sheet): { r: number; c: number }[] {
  const out: { r: number; c: number }[] = [];
  for (const dv of sheet.validations) {
    if (dv.type === 'any') continue;
    for (const rg of dv.ranges) {
      sheet.forEachInRange(rg, (r, c, cell) => {
        if (cell.v === undefined && cell.f === undefined) return;
        const v = cell.f ? S().engine.getValue(sheet, r, c) : cell.v;
        const text = v === null || v === undefined || isErrorVal(v) ? '' : typeof v === 'boolean' ? String(v).toUpperCase() : String(v);
        if (!validValue(sheet, dv, r, c, text)) out.push({ r, c });
      });
    }
  }
  return out;
}

import { isErrorVal } from '../../engine/Engine';
import { isFullCols, isFullRows, Range } from '../../model/address';
import { shiftFormula } from '../../model/formula';
import { primaryRange, rangeSel } from '../../model/selection';
import type { Sheet } from '../../model/sheet';
import type { AutoFilter, Cell, ColumnFilter, FilterOp } from '../../model/types';
import { alertBox, bump, S, setState, transact } from '../store';
import { displayOf, getComputed } from '../values';
import { guardSheet } from './edit';

// ---------- regions ----------

function hasContent(sheet: Sheet, r: number, c: number): boolean {
  const cell = sheet.getCell(r, c);
  return !!cell && (cell.f !== undefined || (cell.v !== undefined && cell.v !== null && cell.v !== ''));
}

/** Excel's CurrentRegion: the block of data around a cell bounded by empty rows/columns. */
export function currentRegion(sheet: Sheet, r: number, c: number): Range {
  let rg: Range = { r1: r, r2: r, c1: c, c2: c };
  let changed = true;
  let guard = 0;
  while (changed && guard++ < 10000) {
    changed = false;
    const tryExpand = (nr: Range, test: () => boolean) => {
      if (test()) {
        rg = nr;
        changed = true;
      }
    };
    if (rg.r1 > 0)
      tryExpand({ ...rg, r1: rg.r1 - 1 }, () => {
        for (let cc = Math.max(0, rg.c1 - 1); cc <= rg.c2 + 1; cc++) if (hasContent(sheet, rg.r1 - 1, cc)) return true;
        return false;
      });
    tryExpand({ ...rg, r2: rg.r2 + 1 }, () => {
      for (let cc = Math.max(0, rg.c1 - 1); cc <= rg.c2 + 1; cc++) if (hasContent(sheet, rg.r2 + 1, cc)) return true;
      return false;
    });
    if (rg.c1 > 0)
      tryExpand({ ...rg, c1: rg.c1 - 1 }, () => {
        for (let rr = Math.max(0, rg.r1 - 1); rr <= rg.r2 + 1; rr++) if (hasContent(sheet, rr, rg.c1 - 1)) return true;
        return false;
      });
    tryExpand({ ...rg, c2: rg.c2 + 1 }, () => {
      for (let rr = Math.max(0, rg.r1 - 1); rr <= rg.r2 + 1; rr++) if (hasContent(sheet, rr, rg.c2 + 1)) return true;
      return false;
    });
  }
  return rg;
}

/** Heuristic header detection (Excel's "My data has headers" guess). */
export function guessHeader(sheet: Sheet, rg: Range): boolean {
  if (rg.r2 <= rg.r1) return false;
  let textTop = 0;
  let numsBelow = 0;
  for (let c = rg.c1; c <= rg.c2; c++) {
    const top = getComputed(sheet, rg.r1, c);
    const below = getComputed(sheet, rg.r1 + 1, c);
    if (typeof top === 'string') textTop++;
    if (typeof below === 'number') numsBelow++;
    const ts = S().wb.styles.get(sheet.styleIdAt(rg.r1, c));
    const bs = S().wb.styles.get(sheet.styleIdAt(rg.r1 + 1, c));
    if (ts.bold && !bs.bold) return true;
  }
  if (textTop === rg.c2 - rg.c1 + 1 && numsBelow > 0) return true;
  // all-text first row and all-text second row: Excel says header if first-row values are unique strings
  if (textTop === rg.c2 - rg.c1 + 1) {
    let allTextBelow = true;
    for (let c = rg.c1; c <= rg.c2; c++) if (typeof getComputed(sheet, rg.r1 + 1, c) !== 'string') allTextBelow = false;
    return !allTextBelow ? true : false;
  }
  return false;
}

/** Range for data commands: selection if multi-cell, else current region. */
export function dataRangeForCommand(): Range | null {
  const st = S();
  const sheet = st.wb.activeSheet;
  let rg = primaryRange(st.sel);
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) {
    if (!hasContent(sheet, rg.r1, rg.c1)) {
      const cr = currentRegion(sheet, rg.r1, rg.c1);
      if (cr.r1 === cr.r2 && cr.c1 === cr.c2) return null;
      return cr;
    }
    rg = currentRegion(sheet, rg.r1, rg.c1);
  } else if (isFullCols(rg) || isFullRows(rg)) {
    const used = sheet.usedRange();
    if (!used) return null;
    rg = { r1: Math.max(rg.r1, used.r1), r2: Math.min(rg.r2, used.r2), c1: Math.max(rg.c1, used.c1), c2: Math.min(rg.c2, used.c2) };
  }
  return rg;
}

// ---------- sort ----------

export interface SortKey {
  col: number; // absolute column (or row when sorting left-to-right)
  desc: boolean;
  by: 'value' | 'cellColor' | 'fontColor';
  color?: string;
  customList?: string[];
}

function typeRank(v: unknown): number {
  if (v === null || v === undefined || v === '') return 5;
  if (typeof v === 'number') return 1;
  if (typeof v === 'string') return 2;
  if (typeof v === 'boolean') return 3;
  return 4; // errors
}

export function compareValues(a: unknown, b: unknown, desc: boolean, caseSensitive = false, customList?: string[]): number {
  const ra = typeRank(a);
  const rb = typeRank(b);
  if (ra === 5 || rb === 5) return ra === rb ? 0 : ra === 5 ? 1 : -1; // blanks always last
  let cmp: number;
  if (customList) {
    const ia = customList.findIndex((x) => x.toLowerCase() === String(a).toLowerCase());
    const ib = customList.findIndex((x) => x.toLowerCase() === String(b).toLowerCase());
    cmp = (ia < 0 ? 1e9 : ia) - (ib < 0 ? 1e9 : ib);
    if (cmp === 0 && ia < 0) cmp = String(a).localeCompare(String(b));
  } else if (ra !== rb) cmp = ra - rb;
  else if (ra === 1) cmp = (a as number) - (b as number);
  else if (ra === 2) {
    const sa = a as string;
    const sb = b as string;
    cmp = caseSensitive ? sa.localeCompare(sb, undefined, { sensitivity: 'case' }) : sa.localeCompare(sb, undefined, { sensitivity: 'base', numeric: false });
  } else if (ra === 3) cmp = (a ? 1 : 0) - (b ? 1 : 0);
  else cmp = 0;
  return desc ? -cmp : cmp;
}

function cellSortValue(sheet: Sheet, r: number, c: number): unknown {
  const v = getComputed(sheet, r, c);
  if (isErrorVal(v)) return v.circular ? 0 : { err: v.error };
  return v;
}

export function sortRange(rg: Range, keys: SortKey[], hasHeader: boolean, opts: { leftToRight?: boolean; caseSensitive?: boolean } = {}): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  if (!guardSheet('sort')) return;
  const merged = sheet.merges.some((m) => m.r1 <= rg.r2 && rg.r1 <= m.r2 && m.c1 <= rg.c2 && rg.c1 <= m.c2 && (m.r1 !== m.r2 || !opts.leftToRight) && m.r2 - m.r1 + 1 !== rg.r2 - rg.r1 + 1);
  if (merged) {
    alertBox('To do this, all the merged cells need to be the same size.');
    return;
  }
  const ltr = !!opts.leftToRight;
  const start = ltr ? rg.c1 + (hasHeader ? 1 : 0) : rg.r1 + (hasHeader ? 1 : 0);
  const end = ltr ? rg.c2 : rg.r2;
  if (end <= start) return;
  const idx: number[] = [];
  for (let i = start; i <= end; i++) idx.push(i);
  const wb = st.wb;
  const keyVal = (i: number, k: SortKey): unknown => {
    const r = ltr ? k.col : i;
    const c = ltr ? i : k.col;
    if (k.by === 'value') return cellSortValue(sheet, r, c);
    const s = wb.styles.get(sheet.styleIdAt(r, c));
    const col = k.by === 'cellColor' ? s.fillColor : s.fontColor;
    return (col ?? '').toUpperCase() === (k.color ?? '').toUpperCase() ? 0 : 1;
  };
  const cache = keys.map((k) => new Map(idx.map((i) => [i, keyVal(i, k)])));
  const sorted = [...idx].sort((a, b) => {
    for (let ki = 0; ki < keys.length; ki++) {
      const k = keys[ki];
      const va = cache[ki].get(a);
      const vb = cache[ki].get(b);
      let cmp: number;
      if (k.by !== 'value') cmp = (va as number) - (vb as number);
      else if (typeof va === 'object' && va && typeof vb === 'object' && vb) cmp = 0;
      else cmp = compareValues(typeof va === 'object' && va ? { e: 1 } : va, typeof vb === 'object' && vb ? { e: 1 } : vb, k.desc, opts.caseSensitive, k.customList);
      if (cmp !== 0) return cmp;
    }
    return a - b; // stable
  });
  if (sorted.every((v, i) => v === idx[i])) {
    bump();
    return;
  }
  transact('Sort', (tx) => {
    // snapshot lines
    const snap = new Map<number, Map<number, Cell | undefined>>();
    for (const i of idx) {
      const line = new Map<number, Cell | undefined>();
      if (ltr) for (let r = rg.r1; r <= rg.r2; r++) line.set(r, sheet.getCell(r, i));
      else for (let c = rg.c1; c <= rg.c2; c++) line.set(c, sheet.getCell(i, c));
      snap.set(i, line);
    }
    sorted.forEach((srcI, k) => {
      const dstI = idx[k];
      if (srcI === dstI) return;
      const line = snap.get(srcI)!;
      for (const [o, cell] of line) {
        const r = ltr ? o : dstI;
        const c = ltr ? dstI : o;
        let next = cell;
        if (cell?.f) next = { ...cell, f: shiftFormula(cell.f, ltr ? 0 : dstI - srcI, ltr ? dstI - srcI : 0) };
        tx.setCell(sheet, r, c, next);
      }
    });
    if (sheet.autoFilter && rangesEq(sheet.autoFilter.range, rg) && keys.length === 1) {
      tx.setMeta(sheet, 'autoFilter', { ...sheet.autoFilter, sort: { col: keys[0].col, desc: keys[0].desc } });
    }
  });
}

function rangesEq(a: Range, b: Range): boolean {
  return a.r1 === b.r1 && a.r2 === b.r2 && a.c1 === b.c1 && a.c2 === b.c2;
}

/** Ribbon Sort A→Z / Z→A. */
export function quickSort(desc: boolean): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  let rg = dataRangeForCommand();
  if (!rg) {
    alertBox('This can\'t be applied to the selected range. Select a single cell in a range and try again.');
    return;
  }
  const af = sheet.autoFilter;
  if (af && st.sel.active.r >= af.range.r1 && st.sel.active.r <= af.range.r2 && st.sel.active.c >= af.range.c1 && st.sel.active.c <= af.range.c2) rg = af.range;
  const table = sheet.tables.find((t) => st.sel.active.r >= t.range.r1 && st.sel.active.r <= t.range.r2 && st.sel.active.c >= t.range.c1 && st.sel.active.c <= t.range.c2);
  if (table) rg = { ...table.range, r2: table.totalRow ? table.range.r2 - 1 : table.range.r2 };
  const header = af && rangesEq(af.range, rg) ? true : table ? table.headerRow : guessHeader(sheet, rg);
  const col = Math.max(rg.c1, Math.min(rg.c2, st.sel.active.c));
  sortRange(rg, [{ col, desc, by: 'value' }], header);
  if (!table && !(af && rangesEq(af.range, rg))) setState({ sel: rangeSel(rg, st.sel.active) });
}

// ---------- AutoFilter ----------

export function toggleAutoFilter(): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  if (!guardSheet('autoFilter')) return;
  if (sheet.autoFilter) {
    transact('Filter', (tx) => {
      tx.setMeta(sheet, 'autoFilter', null);
      tx.setMeta(sheet, 'filteredRows', new Set());
    });
    return;
  }
  const rg = dataRangeForCommand();
  if (!rg) {
    alertBox("This can't be applied to the selected range. Select a single cell in a range and try again.");
    return;
  }
  transact('Filter', (tx) => tx.setMeta(sheet, 'autoFilter', { range: rg, filters: {} }));
}

export function displayTextAt(sheet: Sheet, r: number, c: number): string {
  const style = S().wb.styles.get(sheet.styleIdAt(r, c));
  const d = displayOf(sheet, r, c, style, 30);
  return d.left !== undefined ? (d.left + (d.right ?? '')).trim() : d.text;
}

function testOp(op: FilterOp, val: string, text: string, v: unknown): boolean {
  const num = Number(val);
  const isNum = val.trim() !== '' && !isNaN(num) && typeof v === 'number';
  const wild = (pattern: string) => new RegExp('^' + pattern.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
  switch (op) {
    case 'eq': return isNum ? v === num : wild(val).test(text);
    case 'ne': return isNum ? v !== num : !wild(val).test(text);
    case 'gt': return isNum ? (v as number) > num : text.toLowerCase() > val.toLowerCase();
    case 'gte': return isNum ? (v as number) >= num : text.toLowerCase() >= val.toLowerCase();
    case 'lt': return isNum ? (v as number) < num : text.toLowerCase() < val.toLowerCase();
    case 'lte': return isNum ? (v as number) <= num : text.toLowerCase() <= val.toLowerCase();
    case 'beginsWith': return text.toLowerCase().startsWith(val.toLowerCase());
    case 'endsWith': return text.toLowerCase().endsWith(val.toLowerCase());
    case 'contains': return wild('*' + val + '*').test(text);
    case 'notContains': return !wild('*' + val + '*').test(text);
  }
  return true;
}

function rowPasses(sheet: Sheet, r: number, af: AutoFilter, stats: Map<number, { thresh?: number; avg?: number }>): boolean {
  for (const [k, f] of Object.entries(af.filters)) {
    const c = +k;
    const v = getComputed(sheet, r, c);
    const text = displayTextAt(sheet, r, c);
    switch (f.type) {
      case 'values':
        if (text === '') {
          if (!f.blanks) return false;
        } else if (!f.values.includes(text)) return false;
        break;
      case 'custom': {
        const a = testOp(f.c1.op, f.c1.val, text, v);
        const b = f.c2 ? testOp(f.c2.op, f.c2.val, text, v) : f.and;
        if (f.and ? !(a && b) : !(a || (f.c2 ? b : false))) return false;
        break;
      }
      case 'color': {
        const s = S().wb.styles.get(sheet.styleIdAt(r, c));
        const col = (f.font ? s.fontColor ?? '#000000' : s.fillColor ?? '') || '';
        if (col.toUpperCase() !== f.color.toUpperCase()) return false;
        break;
      }
      case 'top10': {
        const t = stats.get(c)?.thresh;
        if (typeof v !== 'number' || t === undefined) return false;
        if (f.top ? v < t : v > t) return false;
        break;
      }
      case 'dynamic': {
        const avg = stats.get(c)?.avg;
        if (typeof v !== 'number' || avg === undefined) return false;
        if (f.kind === 'aboveAverage' ? v <= avg : v >= avg) return false;
        break;
      }
    }
  }
  return true;
}

export function computeFilteredRows(sheet: Sheet, af: AutoFilter): Set<number> {
  const rows = new Set<number>();
  const stats = new Map<number, { thresh?: number; avg?: number }>();
  for (const [k, f] of Object.entries(af.filters)) {
    if (f.type !== 'top10' && f.type !== 'dynamic') continue;
    const c = +k;
    const nums: number[] = [];
    for (let r = af.range.r1 + 1; r <= af.range.r2; r++) {
      const v = getComputed(sheet, r, c);
      if (typeof v === 'number') nums.push(v);
    }
    if (!nums.length) continue;
    if (f.type === 'dynamic') stats.set(c, { avg: nums.reduce((a, b) => a + b, 0) / nums.length });
    else {
      nums.sort((a, b) => (f.top ? b - a : a - b));
      const n = f.percent ? Math.max(1, Math.floor((nums.length * f.n) / 100)) : Math.min(f.n, nums.length);
      stats.set(c, { thresh: nums[n - 1] });
    }
  }
  for (let r = af.range.r1 + 1; r <= af.range.r2; r++) if (!rowPasses(sheet, r, af, stats)) rows.add(r);
  return rows;
}

export function setColumnFilter(col: number, filter: ColumnFilter | null): void {
  const sheet = S().wb.activeSheet;
  const af = sheet.autoFilter;
  if (!af) return;
  const filters = { ...af.filters };
  if (filter) filters[col] = filter;
  else delete filters[col];
  const next: AutoFilter = { ...af, filters, range: extendFilterRange(sheet, af.range) };
  transact('Filter', (tx) => {
    tx.setMeta(sheet, 'autoFilter', next);
    tx.setMeta(sheet, 'filteredRows', computeFilteredRows(sheet, next));
  });
}

/** Filters grow downward to include newly appended data (Excel does this on reapply). */
function extendFilterRange(sheet: Sheet, rg: Range): Range {
  let r2 = rg.r2;
  while (true) {
    let any = false;
    for (let c = rg.c1; c <= rg.c2; c++) if (hasContent(sheet, r2 + 1, c)) any = true;
    if (!any) break;
    r2++;
  }
  return { ...rg, r2 };
}

export function reapplyFilter(): void {
  const sheet = S().wb.activeSheet;
  const af = sheet.autoFilter;
  if (!af) return;
  const next = { ...af, range: extendFilterRange(sheet, af.range) };
  transact('Reapply', (tx) => {
    tx.setMeta(sheet, 'autoFilter', next);
    tx.setMeta(sheet, 'filteredRows', computeFilteredRows(sheet, next));
  });
}

export function clearAllFilters(): void {
  const sheet = S().wb.activeSheet;
  const af = sheet.autoFilter;
  if (!af) return;
  transact('Clear Filter', (tx) => {
    tx.setMeta(sheet, 'autoFilter', { ...af, filters: {}, sort: undefined });
    tx.setMeta(sheet, 'filteredRows', new Set());
  });
}

/** Distinct display values for a filter dropdown (respecting other columns' filters like Excel). */
export function columnValues(sheet: Sheet, col: number): { values: { text: string; count: number }[]; hasBlanks: boolean; numeric: boolean; dates: boolean } {
  const af = sheet.autoFilter;
  if (!af) return { values: [], hasBlanks: false, numeric: false, dates: false };
  const others: AutoFilter = { ...af, filters: { ...af.filters } };
  delete others.filters[col];
  const hidden = computeFilteredRows(sheet, others);
  const map = new Map<string, { count: number; sortKey: unknown }>();
  let hasBlanks = false;
  let nums = 0;
  let total = 0;
  for (let r = af.range.r1 + 1; r <= af.range.r2; r++) {
    if (hidden.has(r) || sheet.hiddenRows.has(r)) continue;
    const text = displayTextAt(sheet, r, col);
    if (text === '') {
      hasBlanks = true;
      continue;
    }
    total++;
    const v = getComputed(sheet, r, col);
    if (typeof v === 'number') nums++;
    const e = map.get(text);
    if (e) e.count++;
    else map.set(text, { count: 1, sortKey: v });
  }
  const values = [...map.entries()].sort((a, b) => compareValues(a[1].sortKey, b[1].sortKey, false)).map(([text, e]) => ({ text, count: e.count }));
  return { values, hasBlanks, numeric: total > 0 && nums === total, dates: false };
}

export function sortByColumn(col: number, desc: boolean): void {
  const sheet = S().wb.activeSheet;
  const af = sheet.autoFilter;
  if (!af) return;
  sortRange(af.range, [{ col, desc, by: 'value' }], true);
  // sorting changes which rows are hidden
  const next = sheet.autoFilter!;
  transact('Filter', (tx) => tx.setMeta(sheet, 'filteredRows', computeFilteredRows(sheet, next)));
}

export function sortByColor(col: number, color: string, font: boolean): void {
  const sheet = S().wb.activeSheet;
  const af = sheet.autoFilter;
  if (!af) return;
  sortRange(af.range, [{ col, desc: false, by: font ? 'fontColor' : 'cellColor', color }], true);
}

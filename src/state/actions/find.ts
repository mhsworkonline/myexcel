import { isErrorVal } from '../../engine/Engine';
import { CellAddr, MAX_COLS, MAX_ROWS, parseRange, Range } from '../../model/address';
import { refToRange, tokenize } from '../../model/formula';
import { Selection, singleSel } from '../../model/selection';
import type { Sheet } from '../../model/sheet';
import { alertBox, bump, requestScroll, S, setState, setStatus, transact } from '../store';
import { editText, getComputed } from '../values';
import { writeInput } from './edit';
import { currentRegion, displayTextAt } from './sortFilter';
import { activateSheet } from './structure';

export interface FindOptions {
  what: string;
  within: 'sheet' | 'workbook';
  searchBy: 'rows' | 'cols';
  lookIn: 'formulas' | 'values' | 'notes';
  matchCase: boolean;
  entireCell: boolean;
}

export const defaultFind: FindOptions = { what: '', within: 'sheet', searchBy: 'rows', lookIn: 'formulas', matchCase: false, entireCell: false };

export interface FindHit {
  sheetId: string;
  sheetName: string;
  r: number;
  c: number;
  value: string;
  formula: string;
}

function patternFor(o: FindOptions): RegExp {
  // Excel wildcards: * ? with ~ escape
  let src = '';
  const w = o.what;
  for (let i = 0; i < w.length; i++) {
    const ch = w[i];
    if (ch === '~' && i + 1 < w.length) {
      src += w[i + 1].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      i++;
    } else if (ch === '*') src += '.*';
    else if (ch === '?') src += '.';
    else src += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(o.entireCell ? `^${src}$` : src, o.matchCase ? '' : 'i');
}

function cellText(sheet: Sheet, r: number, c: number, o: FindOptions): string {
  const cell = sheet.getCell(r, c);
  if (!cell) return '';
  if (o.lookIn === 'notes') return cell.note?.text ?? '';
  if (o.lookIn === 'formulas') return editText(sheet, r, c);
  return displayTextAt(sheet, r, c);
}

function sheetHits(sheet: Sheet, o: FindOptions, re: RegExp): FindHit[] {
  const hits: FindHit[] = [];
  const keys = sheet.sortedRowKeys();
  for (const r of keys) {
    const row = sheet.rows.get(r)!;
    const cols = [...row.keys()].sort((a, b) => a - b);
    for (const c of cols) {
      const t = cellText(sheet, r, c, o);
      if (t && re.test(t)) hits.push({ sheetId: sheet.id, sheetName: sheet.name, r, c, value: displayTextAt(sheet, r, c), formula: sheet.getCell(r, c)?.f ?? '' });
    }
  }
  if (o.searchBy === 'cols') hits.sort((a, b) => a.c - b.c || a.r - b.r);
  return hits;
}

export function findAll(o: FindOptions): FindHit[] {
  if (!o.what) return [];
  const re = patternFor(o);
  const wb = S().wb;
  const sheets = o.within === 'workbook' ? wb.sheets : [wb.activeSheet];
  return sheets.flatMap((s) => sheetHits(s, o, re));
}

/** Restrict search to the selection when more than one cell is selected (Excel behaviour). */
function selectionScope(): Range | null {
  const st = S();
  const rg = st.sel.ranges[st.sel.ranges.length - 1];
  if (st.sel.ranges.length === 1 && rg.r1 === rg.r2 && rg.c1 === rg.c2) return null;
  return rg;
}

export function findNext(o: FindOptions, backwards = false): FindHit | null {
  const st = S();
  let hits = findAll(o);
  const scope = selectionScope();
  if (scope && o.within === 'sheet') hits = hits.filter((h) => h.r >= scope.r1 && h.r <= scope.r2 && h.c >= scope.c1 && h.c <= scope.c2);
  if (!hits.length) {
    alertBox("We couldn't find what you were looking for. Select Options for more ways to search.", 'MyExcel', 'info');
    return null;
  }
  const wb = st.wb;
  const sheetOrder = (id: string) => wb.sheets.findIndex((s) => s.id === id);
  const cur = { s: sheetOrder(wb.activeSheetId), r: st.sel.active.r, c: st.sel.active.c };
  const key = (h: { s: number; r: number; c: number }) => (o.searchBy === 'rows' ? [h.s, h.r, h.c] : [h.s, h.c, h.r]);
  const cmp = (a: number[], b: number[]) => a[0] - b[0] || a[1] - b[1] || a[2] - b[2];
  const ck = key(cur);
  const withKeys = hits.map((h) => ({ h, k: key({ s: sheetOrder(h.sheetId), r: h.r, c: h.c }) }));
  let target = backwards ? [...withKeys].reverse().find((x) => cmp(x.k, ck) < 0) : withKeys.find((x) => cmp(x.k, ck) > 0);
  if (!target) target = backwards ? withKeys[withKeys.length - 1] : withKeys[0];
  gotoHit(target.h, scope ?? undefined);
  return target.h;
}

export function gotoHit(h: FindHit, keepScope?: Range): void {
  const st = S();
  if (h.sheetId !== st.wb.activeSheetId) activateSheet(h.sheetId);
  if (keepScope) setState({ sel: { ...S().sel, active: { r: h.r, c: h.c } } });
  else setState({ sel: singleSel(h.r, h.c) });
  requestScroll(h.r, h.c);
  bump();
}

function replaceIn(text: string, o: FindOptions, withText: string): string {
  const re = patternFor(o);
  if (o.entireCell) return re.test(text) ? withText : text;
  return text.replace(new RegExp(re.source, o.matchCase ? 'g' : 'gi'), withText.replace(/\$/g, '$$$$'));
}

export function replaceCurrent(o: FindOptions, withText: string): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  const { r, c } = st.sel.active;
  const t = cellText(sheet, r, c, { ...o, lookIn: 'formulas' });
  if (t && patternFor(o).test(t)) {
    transact('Replace', (tx) => writeInput(tx, sheet, r, c, replaceIn(t, o, withText)));
  }
  findNext(o);
}

export function replaceAll(o: FindOptions, withText: string): number {
  const hits = findAll({ ...o, lookIn: 'formulas' });
  const wb = S().wb;
  let n = 0;
  transact('Replace All', (tx) => {
    for (const h of hits) {
      const sheet = wb.sheetById(h.sheetId)!;
      if (sheet.protection) continue;
      const t = editText(sheet, h.r, h.c);
      const nt = replaceIn(t, o, withText);
      if (nt !== t) {
        writeInput(tx, sheet, h.r, h.c, nt);
        n++;
      }
    }
  });
  alertBox(n ? `All done. We made ${n} replacement${n === 1 ? '' : 's'}.` : "We couldn't find anything to replace. Click Options for more ways to search.", 'MyExcel', 'info');
  return n;
}

// ---------- Go To ----------

/** Resolve a reference or name typed into the Name Box / Go To dialog. */
export function resolveReference(text: string): { sheetId: string; range: Range } | null {
  const wb = S().wb;
  let t = text.trim();
  if (!t) return null;
  let sheet = wb.activeSheet;
  const bang = t.lastIndexOf('!');
  if (bang > 0) {
    const sn = t.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'");
    const s = wb.sheetByName(sn);
    if (!s) return null;
    sheet = s;
    t = t.slice(bang + 1);
  }
  const rg = parseRange(t.replace(/\$/g, ''));
  if (rg) return { sheetId: sheet.id, range: rg };
  // single column/row like "C" or "5"?
  if (/^[A-Za-z]{1,3}$/.test(t)) {
    const r2 = parseRange(`${t}:${t}`);
    if (r2) return { sheetId: sheet.id, range: r2 };
  }
  if (/^\d+$/.test(t)) {
    const r2 = parseRange(`${t}:${t}`);
    if (r2) return { sheetId: sheet.id, range: r2 };
  }
  // defined name
  const nm = wb.names.find((n) => n.name.toLowerCase() === t.toLowerCase() && (!n.scope || n.scope === wb.activeSheetId));
  if (nm) return resolveReference(nm.ref);
  // table name
  for (const s of wb.sheets) {
    const tb = s.tables.find((x) => x.name.toLowerCase() === t.toLowerCase());
    if (tb) return { sheetId: s.id, range: { ...tb.range, r1: tb.range.r1 + (tb.headerRow ? 1 : 0), r2: tb.range.r2 - (tb.totalRow ? 1 : 0) } };
  }
  return null;
}

export function goTo(text: string): boolean {
  const res = resolveReference(text);
  if (!res) return false;
  if (res.sheetId !== S().wb.activeSheetId) activateSheet(res.sheetId);
  const rg = res.range;
  setState({ sel: { ranges: [rg], active: { r: rg.r1, c: rg.c1 }, anchor: { r: rg.r1, c: rg.c1 } } });
  requestScroll(rg.r1, rg.c1);
  bump();
  return true;
}

// ---------- Go To Special ----------

export type SpecialKind =
  | 'notes'
  | 'constants'
  | 'formulas'
  | 'blanks'
  | 'currentRegion'
  | 'rowDiffs'
  | 'colDiffs'
  | 'precedents'
  | 'dependents'
  | 'lastCell'
  | 'visible'
  | 'conditionalFormats'
  | 'validation';

export interface SpecialOpts {
  kind: SpecialKind;
  numbers?: boolean;
  text?: boolean;
  logicals?: boolean;
  errors?: boolean;
}

function valueKind(v: unknown): 'numbers' | 'text' | 'logicals' | 'errors' | null {
  if (isErrorVal(v)) return 'errors';
  if (typeof v === 'number') return 'numbers';
  if (typeof v === 'string') return v === '' ? null : 'text';
  if (typeof v === 'boolean') return 'logicals';
  return null;
}

function cellsToRanges(cells: CellAddr[]): Range[] {
  // merge horizontally adjacent cells into row runs, then vertically identical runs
  const byRow = new Map<number, number[]>();
  for (const { r, c } of cells) {
    let l = byRow.get(r);
    if (!l) byRow.set(r, (l = []));
    l.push(c);
  }
  const runs: Range[] = [];
  for (const [r, cols] of [...byRow.entries()].sort((a, b) => a[0] - b[0])) {
    cols.sort((a, b) => a - b);
    let s = cols[0];
    let p = cols[0];
    for (let i = 1; i <= cols.length; i++) {
      if (i < cols.length && cols[i] === p + 1) {
        p = cols[i];
        continue;
      }
      runs.push({ r1: r, r2: r, c1: s, c2: p });
      if (i < cols.length) s = p = cols[i];
    }
  }
  const out: Range[] = [];
  for (const run of runs) {
    const prev = out.find((o) => o.r2 === run.r1 - 1 && o.c1 === run.c1 && o.c2 === run.c2);
    if (prev) prev.r2 = run.r2;
    else out.push({ ...run });
  }
  return out.slice(0, 5000);
}

export function goToSpecial(o: SpecialOpts): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  const sel = st.sel;
  const a = sel.active;
  const multi = sel.ranges.length > 1 || sel.ranges[0].r1 !== sel.ranges[0].r2 || sel.ranges[0].c1 !== sel.ranges[0].c2;
  const used = sheet.usedRange();
  const scope: Range[] = multi ? sel.ranges : used ? [used] : [];
  let found: CellAddr[] = [];
  const inScope = (r: number, c: number) => scope.some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2);
  switch (o.kind) {
    case 'currentRegion': {
      const rg = currentRegion(sheet, a.r, a.c);
      setSel({ ranges: [rg], active: a, anchor: a });
      return;
    }
    case 'lastCell': {
      if (!used) return;
      setSel(singleSel(used.r2, used.c2));
      requestScroll(used.r2, used.c2);
      return;
    }
    case 'blanks':
      for (const rg of scope)
        for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 100000); r++)
          for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 1000); c++) {
            const cell = sheet.getCell(r, c);
            if (!cell || (cell.v === undefined && cell.f === undefined) || cell.v === '') found.push({ r, c });
          }
      break;
    case 'constants':
    case 'formulas':
      for (const [r, row] of sheet.rows)
        for (const [c, cell] of row) {
          if (!inScope(r, c)) continue;
          const isF = cell.f !== undefined;
          if ((o.kind === 'formulas') !== isF) continue;
          const v = isF ? getComputed(sheet, r, c) : cell.e ? { error: String(cell.v) } : cell.v;
          const k = valueKind(v);
          if (k && o[k] !== false) found.push({ r, c });
        }
      break;
    case 'notes':
      for (const [r, row] of sheet.rows) for (const [c, cell] of row) if (cell.note) found.push({ r, c });
      break;
    case 'conditionalFormats':
      for (const cf of sheet.conditionalFormats) for (const rg of cf.ranges) for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 10000); r++) for (let c = rg.c1; c <= rg.c2; c++) found.push({ r, c });
      break;
    case 'validation':
      for (const dv of sheet.validations) for (const rg of dv.ranges) for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 10000); r++) for (let c = rg.c1; c <= rg.c2; c++) found.push({ r, c });
      break;
    case 'visible': {
      const out: CellAddr[] = [];
      for (const rg of sel.ranges) {
        const r2 = Math.min(rg.r2, used?.r2 ?? rg.r2);
        const c2 = Math.min(rg.c2, used?.c2 ?? rg.c2);
        for (let r = rg.r1; r <= r2; r++) {
          if (sheet.isRowHidden(r)) continue;
          for (let c = rg.c1; c <= c2; c++) if (!sheet.isColHidden(c)) out.push({ r, c });
        }
      }
      found = out;
      break;
    }
    case 'rowDiffs':
    case 'colDiffs': {
      for (const rg of scope) {
        if (o.kind === 'rowDiffs') {
          for (let r = rg.r1; r <= rg.r2; r++) {
            const ref = getComputed(sheet, r, a.c);
            for (let c = rg.c1; c <= rg.c2; c++) if (c !== a.c && JSON.stringify(getComputed(sheet, r, c)) !== JSON.stringify(ref)) found.push({ r, c });
          }
        } else {
          for (let c = rg.c1; c <= rg.c2; c++) {
            const ref = getComputed(sheet, a.r, c);
            for (let r = rg.r1; r <= rg.r2; r++) if (r !== a.r && JSON.stringify(getComputed(sheet, r, c)) !== JSON.stringify(ref)) found.push({ r, c });
          }
        }
      }
      break;
    }
    case 'precedents': {
      const f = sheet.getCell(a.r, a.c)?.f;
      if (f) {
        const ranges = tokenize(f)
          .filter((t) => t.type === 'ref' && t.ref && (!t.ref.sheet || t.ref.sheet.toLowerCase() === sheet.name.toLowerCase()))
          .map((t) => refToRange(t.ref!));
        if (ranges.length) {
          setSel({ ranges, active: { r: ranges[0].r1, c: ranges[0].c1 }, anchor: { r: ranges[0].r1, c: ranges[0].c1 } });
          return;
        }
      }
      break;
    }
    case 'dependents': {
      for (const [r, row] of sheet.rows)
        for (const [c, cell] of row) {
          if (!cell.f) continue;
          const deps = tokenize(cell.f).filter((t) => t.type === 'ref' && t.ref && (!t.ref.sheet || t.ref.sheet.toLowerCase() === sheet.name.toLowerCase()));
          if (deps.some((t) => {
            const rg = refToRange(t.ref!);
            return a.r >= rg.r1 && a.r <= rg.r2 && a.c >= rg.c1 && a.c <= rg.c2;
          }))
            found.push({ r, c });
        }
      break;
    }
  }
  if (!found.length) {
    alertBox('No cells were found.', 'MyExcel', 'info');
    return;
  }
  const ranges = cellsToRanges(found);
  const first = ranges[0];
  setSel({ ranges, active: { r: first.r1, c: first.c1 }, anchor: { r: first.r1, c: first.c1 } });
  setStatus(null);
}

function setSel(sel: Selection): void {
  setState({ sel });
  requestScroll(sel.active.r, sel.active.c);
  bump();
}

export function selectAllCells(): void {
  setState({ sel: { ranges: [{ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }], active: S().sel.active, anchor: { r: 0, c: 0 } } });
}

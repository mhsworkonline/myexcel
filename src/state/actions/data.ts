import { isErrorVal } from '../../engine/Engine';
import { addrToA1, colToName, isFullCols, Range, rangeToA1 } from '../../model/address';
import { primaryRange, rangeSel } from '../../model/selection';
import { newId } from '../../model/sheet';
import type { CFRule, DataValidation, Note, SheetProtection, TableDef } from '../../model/types';
import { alertBox, bump, openDialog, S, setState, transact } from '../store';
import { getComputed } from '../values';
import { applyStyleFn } from './format';
import { guardRanges, guardSheet, writeInput } from './edit';
import { currentRegion, dataRangeForCommand, guessHeader } from './sortFilter';

// ---------- conditional formatting ----------

export function selectionRanges(): Range[] {
  const st = S();
  const sheet = st.wb.activeSheet;
  const used = sheet.usedRange();
  return st.sel.ranges.map((rg) => (isFullCols(rg) && used ? { ...rg, r2: Math.max(used.r2, rg.r1) } : rg));
}

export function addCFRule(rule: Omit<CFRule, 'id' | 'priority' | 'ranges'> & { ranges?: Range[] }): void {
  const sheet = S().wb.activeSheet;
  if (!guardSheet('formatCells')) return;
  const ranges = rule.ranges ?? selectionRanges();
  // new rules get top priority (1) like Excel
  const existing = sheet.conditionalFormats.map((r) => ({ ...r, priority: r.priority + 1 }));
  const nr: CFRule = { ...rule, id: newId('cf'), priority: 1, ranges } as CFRule;
  transact('Conditional Formatting', (tx) => tx.setMeta(sheet, 'conditionalFormats', [nr, ...existing]));
}

export function setCFRules(rules: CFRule[]): void {
  const sheet = S().wb.activeSheet;
  transact('Conditional Formatting', (tx) => tx.setMeta(sheet, 'conditionalFormats', rules));
}

export function clearCFRules(scope: 'selection' | 'sheet'): void {
  const sheet = S().wb.activeSheet;
  if (scope === 'sheet') {
    transact('Clear Rules', (tx) => tx.setMeta(sheet, 'conditionalFormats', []));
    return;
  }
  const sel = selectionRanges();
  const rules = sheet.conditionalFormats
    .map((r) => ({ ...r, ranges: r.ranges.flatMap((rg) => subtractRanges(rg, sel)) }))
    .filter((r) => r.ranges.length);
  transact('Clear Rules', (tx) => tx.setMeta(sheet, 'conditionalFormats', rules));
}

/** Remove `cut` rectangles from `rg`, returning the remaining pieces. */
export function subtractRanges(rg: Range, cuts: Range[]): Range[] {
  let pieces: Range[] = [rg];
  for (const c of cuts) {
    const next: Range[] = [];
    for (const p of pieces) {
      if (c.r1 > p.r2 || c.r2 < p.r1 || c.c1 > p.c2 || c.c2 < p.c1) {
        next.push(p);
        continue;
      }
      if (c.r1 > p.r1) next.push({ ...p, r2: c.r1 - 1 });
      if (c.r2 < p.r2) next.push({ ...p, r1: c.r2 + 1 });
      const r1 = Math.max(p.r1, c.r1);
      const r2 = Math.min(p.r2, c.r2);
      if (c.c1 > p.c1) next.push({ r1, r2, c1: p.c1, c2: c.c1 - 1 });
      if (c.c2 < p.c2) next.push({ r1, r2, c1: c.c2 + 1, c2: p.c2 });
    }
    pieces = next;
  }
  return pieces;
}

export const CF_PRESETS: Record<string, { fillColor?: string; fontColor?: string; borderColor?: string; bold?: boolean }> = {
  'Light Red Fill with Dark Red Text': { fillColor: '#FFC7CE', fontColor: '#9C0006' },
  'Yellow Fill with Dark Yellow Text': { fillColor: '#FFEB9C', fontColor: '#9C5700' },
  'Green Fill with Dark Green Text': { fillColor: '#C6EFCE', fontColor: '#006100' },
  'Light Red Fill': { fillColor: '#FFC7CE' },
  'Red Text': { fontColor: '#9C0006' },
  'Red Border': { borderColor: '#9C0006' },
};

export const DATA_BAR_COLORS = ['#638EC6', '#63BE7B', '#FF555A', '#FFB628', '#008AEF', '#D6007B'];

export const COLOR_SCALES: { name: string; colors: string[] }[] = [
  { name: 'Green - Yellow - Red', colors: ['#F8696B', '#FFEB84', '#63BE7B'] },
  { name: 'Red - Yellow - Green', colors: ['#63BE7B', '#FFEB84', '#F8696B'] },
  { name: 'Green - White - Red', colors: ['#F8696B', '#FCFCFF', '#63BE7B'] },
  { name: 'Red - White - Green', colors: ['#63BE7B', '#FCFCFF', '#F8696B'] },
  { name: 'Blue - White - Red', colors: ['#F8696B', '#FCFCFF', '#5A8AC6'] },
  { name: 'Red - White - Blue', colors: ['#5A8AC6', '#FCFCFF', '#F8696B'] },
  { name: 'White - Red', colors: ['#F8696B', '#FCFCFF'] },
  { name: 'Red - White', colors: ['#FCFCFF', '#F8696B'] },
  { name: 'Green - White', colors: ['#FCFCFF', '#63BE7B'] },
  { name: 'White - Green', colors: ['#63BE7B', '#FCFCFF'] },
  { name: 'Green - Yellow', colors: ['#FFEF9C', '#63BE7B'] },
  { name: 'Yellow - Green', colors: ['#63BE7B', '#FFEF9C'] },
];

export const ICON_SETS = [
  '3Arrows',
  '3ArrowsGray',
  '3Triangles',
  '4Arrows',
  '4ArrowsGray',
  '5Arrows',
  '5ArrowsGray',
  '3TrafficLights1',
  '3TrafficLights2',
  '3Signs',
  '4TrafficLights',
  '3Symbols',
  '3Symbols2',
  '3Flags',
  '3Stars',
  '4Rating',
  '5Rating',
  '5Quarters',
  '5Boxes',
];

export function addDataBar(color: string, gradient: boolean): void {
  addCFRule({ type: 'dataBar', dataBar: { color, gradient, min: { type: 'autoMin' }, max: { type: 'autoMax' }, showValue: true } });
}

export function addColorScale(colors: string[]): void {
  const cfvos =
    colors.length === 3
      ? [{ type: 'min' as const }, { type: 'percentile' as const, value: 50 }, { type: 'max' as const }]
      : [{ type: 'min' as const }, { type: 'max' as const }];
  // gallery lists high→low; rules are stored low→high
  addCFRule({ type: 'colorScale', colorScale: { cfvos, colors: [...colors] } });
}

export function addIconSet(set: string): void {
  const n = parseInt(set, 10) || 3;
  const cfvos = Array.from({ length: n }, (_, i) => ({ type: 'percent' as const, value: Math.round((i * 100) / n) }));
  addCFRule({ type: 'iconSet', iconSet: { set, cfvos, showValue: true } });
}

// ---------- data validation ----------

export function setValidation(dv: Omit<DataValidation, 'id' | 'ranges'> | null, ranges?: Range[]): void {
  const sheet = S().wb.activeSheet;
  const target = ranges ?? selectionRanges();
  const kept = sheet.validations
    .map((v) => ({ ...v, ranges: v.ranges.flatMap((rg) => subtractRanges(rg, target)) }))
    .filter((v) => v.ranges.length);
  const next = dv ? kept.concat([{ ...dv, id: newId('dv'), ranges: target }]) : kept;
  transact('Data Validation', (tx) => tx.setMeta(sheet, 'validations', next));
}

// ---------- tables ----------

export function uniqueTableName(): string {
  const all = new Set(S().wb.sheets.flatMap((s) => s.tables.map((t) => t.name.toLowerCase())));
  let i = 1;
  while (all.has(`table${i}`)) i++;
  return `Table${i}`;
}

export function createTable(range: Range, hasHeaders: boolean, style = 'TableStyleMedium2'): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  if (!guardRanges(sheet, [range])) return;
  if (sheet.tables.some((t) => t.range.r1 <= range.r2 && range.r1 <= t.range.r2 && t.range.c1 <= range.c2 && range.c1 <= t.range.c2)) {
    alertBox("A table can't overlap another table.");
    return;
  }
  if (sheet.merges.some((m) => m.r1 <= range.r2 && range.r1 <= m.r2 && m.c1 <= range.c2 && range.c1 <= m.c2)) {
    alertBox("A table can't overlap a range that contains merged cells.");
    return;
  }
  let rg = range;
  transact('Create Table', (tx) => {
    if (!hasHeaders) {
      // insert a header row by shifting the data down one row within the table columns
      const cells: { r: number; c: number; cell: import('../../model/types').Cell }[] = [];
      sheet.forEachInRange(rg, (r, c, cell) => cells.push({ r, c, cell }));
      for (const x of cells) tx.setCell(sheet, x.r, x.c, undefined);
      for (const x of cells) tx.setCell(sheet, x.r + 1, x.c, x.cell);
      for (let c = rg.c1; c <= rg.c2; c++) writeInput(tx, sheet, rg.r1, c, `Column${c - rg.c1 + 1}`);
      rg = { ...rg, r2: rg.r2 + 1 };
    }
    const names: string[] = [];
    for (let c = rg.c1; c <= rg.c2; c++) {
      let n = String(getComputed(sheet, rg.r1, c) ?? '').trim() || `Column${c - rg.c1 + 1}`;
      let k = 2;
      const base = n;
      while (names.includes(n)) n = `${base}${k++}`;
      names.push(n);
      if (String(sheet.getCell(rg.r1, c)?.v ?? '') !== n) writeInput(tx, sheet, rg.r1, c, n);
    }
    const t: TableDef = {
      id: newId('tb'),
      name: uniqueTableName(),
      range: rg,
      headerRow: true,
      totalRow: false,
      style,
      bandedRows: true,
      bandedCols: false,
      firstCol: false,
      lastCol: false,
      showFilterButton: true,
      columns: names.map((name) => ({ name, totalFunction: 'none' })),
    };
    tx.setMeta(sheet, 'tables', sheet.tables.concat([t]));
    if (!sheet.autoFilter) tx.setMeta(sheet, 'autoFilter', { range: rg, filters: {} });
  }, rangeSel(rg));
}

export function updateTable(id: string, patch: Partial<TableDef>): void {
  const sheet = S().wb.activeSheet;
  const t = sheet.tables.find((x) => x.id === id);
  if (!t) return;
  transact('Table', (tx) => {
    let next = { ...t, ...patch };
    if (patch.totalRow !== undefined && patch.totalRow !== t.totalRow) {
      if (patch.totalRow) {
        next = { ...next, range: { ...next.range, r2: next.range.r2 + 1 } };
        const tr = next.range.r2;
        next.columns = next.columns.map((c, i) => (i === 0 ? { ...c, totalLabel: 'Total', totalFunction: 'none' } : i === next.columns.length - 1 ? { ...c, totalFunction: 'sum' } : c));
        writeTotals(tx, next, tr);
      } else {
        const tr = t.range.r2;
        for (let c = t.range.c1; c <= t.range.c2; c++) tx.setCell(sheet, tr, c, undefined);
        next = { ...next, range: { ...next.range, r2: next.range.r2 - 1 } };
      }
    }
    tx.setMeta(sheet, 'tables', sheet.tables.map((x) => (x.id === id ? next : x)));
    if (sheet.autoFilter && sheet.autoFilter.range.r1 === t.range.r1 && sheet.autoFilter.range.c1 === t.range.c1) {
      if (patch.showFilterButton === false) tx.setMeta(sheet, 'autoFilter', null);
      else tx.setMeta(sheet, 'autoFilter', { ...sheet.autoFilter, range: { ...next.range, r2: next.totalRow ? next.range.r2 - 1 : next.range.r2 } });
    }
  });
}

const SUBTOTAL: Record<string, number> = { average: 101, count: 103, countNums: 102, max: 104, min: 105, sum: 109, stdDev: 107, var: 110 };

function writeTotals(tx: import('../../model/commands').Tx, t: TableDef, tr: number): void {
  const sheet = S().wb.activeSheet;
  t.columns.forEach((col, i) => {
    const c = t.range.c1 + i;
    if (col.totalLabel) writeInput(tx, sheet, tr, c, col.totalLabel);
    else if (col.totalFunction && col.totalFunction !== 'none' && col.totalFunction !== 'custom') {
      const data = { r1: t.range.r1 + 1, r2: tr - 1, c1: c, c2: c };
      writeInput(tx, sheet, tr, c, `=SUBTOTAL(${SUBTOTAL[col.totalFunction]},${rangeToA1(data)})`);
    } else tx.setCell(sheet, tr, c, undefined);
  });
}

export function setTotalFunction(id: string, colIdx: number, fn: TableDef['columns'][0]['totalFunction']): void {
  const sheet = S().wb.activeSheet;
  const t = sheet.tables.find((x) => x.id === id);
  if (!t || !t.totalRow) return;
  const next = { ...t, columns: t.columns.map((c, i) => (i === colIdx ? { ...c, totalFunction: fn, totalLabel: undefined } : c)) };
  transact('Total Row', (tx) => {
    tx.setMeta(sheet, 'tables', sheet.tables.map((x) => (x.id === id ? next : x)));
    writeTotals(tx, next, next.range.r2);
  });
}

export function convertTableToRange(id: string): void {
  const sheet = S().wb.activeSheet;
  const t = sheet.tables.find((x) => x.id === id);
  if (!t) return;
  transact('Convert to Range', (tx) => {
    tx.setMeta(sheet, 'tables', sheet.tables.filter((x) => x.id !== id));
    if (sheet.autoFilter && sheet.autoFilter.range.r1 === t.range.r1 && sheet.autoFilter.range.c1 === t.range.c1) tx.setMeta(sheet, 'autoFilter', null);
  });
}

export function tableAtActive(): TableDef | undefined {
  const st = S();
  const { r, c } = st.sel.active;
  return st.wb.activeSheet.tables.find((t) => r >= t.range.r1 && r <= t.range.r2 && c >= t.range.c1 && c <= t.range.c2);
}

export function defaultTableRange(): Range | null {
  const rg = dataRangeForCommand();
  return rg;
}

// ---------- remove duplicates ----------

export function removeDuplicates(range: Range, cols: number[], hasHeader: boolean): void {
  const sheet = S().wb.activeSheet;
  const seen = new Set<string>();
  const keep: number[] = [];
  const start = range.r1 + (hasHeader ? 1 : 0);
  for (let r = start; r <= range.r2; r++) {
    const key = cols
      .map((c) => {
        const v = getComputed(sheet, r, c);
        return isErrorVal(v) ? v.error : typeof v === 'string' ? v.toLowerCase() : String(v ?? '');
      })
      .join('\u0001');
    if (seen.has(key)) continue;
    seen.add(key);
    keep.push(r);
  }
  const removed = range.r2 - start + 1 - keep.length;
  transact('Remove Duplicates', (tx) => {
    const rows = keep.map((r) => {
      const line: (import('../../model/types').Cell | undefined)[] = [];
      for (let c = range.c1; c <= range.c2; c++) line.push(sheet.getCell(r, c));
      return line;
    });
    for (let r = start; r <= range.r2; r++) for (let c = range.c1; c <= range.c2; c++) if (sheet.getCell(r, c)) tx.setCell(sheet, r, c, undefined);
    rows.forEach((line, i) => line.forEach((cell, j) => cell && tx.setCell(sheet, start + i, range.c1 + j, cell)));
  });
  alertBox(removed ? `${removed} duplicate values found and removed; ${keep.length} unique values remain.` : `No duplicate values found.`, 'MyExcel', 'info');
}

// ---------- text to columns ----------

export interface TextToColumnsOpts {
  mode: 'delimited' | 'fixed';
  delimiters: { tab: boolean; semicolon: boolean; comma: boolean; space: boolean; other: string };
  consecutive: boolean;
  qualifier: '"' | "'" | '';
  widths: number[]; // break positions for fixed width
  dest?: { r: number; c: number };
  skipCols?: number[];
}

export function splitText(text: string, o: TextToColumnsOpts): string[] {
  if (o.mode === 'fixed') {
    const out: string[] = [];
    let prev = 0;
    for (const b of [...o.widths].sort((a, b) => a - b)) {
      out.push(text.slice(prev, b).trim());
      prev = b;
    }
    out.push(text.slice(prev).trim());
    return out;
  }
  const ds = new Set<string>();
  if (o.delimiters.tab) ds.add('\t');
  if (o.delimiters.semicolon) ds.add(';');
  if (o.delimiters.comma) ds.add(',');
  if (o.delimiters.space) ds.add(' ');
  if (o.delimiters.other) ds.add(o.delimiters.other[0]);
  const out: string[] = [];
  let cur = '';
  let q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (o.qualifier && ch === o.qualifier) {
      q = !q;
      continue;
    }
    if (!q && ds.has(ch)) {
      out.push(cur);
      cur = '';
      if (o.consecutive) while (i + 1 < text.length && ds.has(text[i + 1])) i++;
      continue;
    }
    cur += ch;
  }
  out.push(cur);
  return out;
}

export function textToColumns(o: TextToColumnsOpts): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  const rg = primaryRange(st.sel);
  if (rg.c1 !== rg.c2) {
    alertBox('Text to Columns can only convert one column at a time. The range can be many rows tall but no more than one column wide. Try again after selecting cells in one column only.');
    return;
  }
  const used = sheet.usedRange();
  const r2 = Math.min(rg.r2, used?.r2 ?? rg.r2);
  const dest = o.dest ?? { r: rg.r1, c: rg.c1 };
  transact('Text to Columns', (tx) => {
    for (let r = rg.r1; r <= r2; r++) {
      const v = getComputed(sheet, r, rg.c1);
      if (v === null || v === undefined || v === '') continue;
      const parts = splitText(String(v), o);
      let out = 0;
      parts.forEach((p, i) => {
        if (o.skipCols?.includes(i)) return;
        writeInput(tx, sheet, dest.r + (r - rg.r1), dest.c + out, p);
        out++;
      });
    }
  });
}

// ---------- protection ----------

export async function hashPassword(pw: string): Promise<string> {
  if (!pw) return '';
  const data = new TextEncoder().encode(pw);
  const buf = await crypto.subtle.digest('SHA-256', data);
  return btoa(String.fromCharCode(...new Uint8Array(buf)));
}

export async function protectSheet(password: string, allow: SheetProtection['allow']): Promise<void> {
  const sheet = S().wb.activeSheet;
  const passwordHash = password ? await hashPassword(password) : undefined;
  transact('Protect Sheet', (tx) => tx.setMeta(sheet, 'protection', { passwordHash, allow }));
}

export async function unprotectSheet(password: string): Promise<boolean> {
  const sheet = S().wb.activeSheet;
  const p = sheet.protection;
  if (!p) return true;
  if (p.passwordHash && (await hashPassword(password)) !== p.passwordHash) {
    alertBox('The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization.');
    return false;
  }
  transact('Unprotect Sheet', (tx) => tx.setMeta(sheet, 'protection', null));
  return true;
}

export const DEFAULT_ALLOW: SheetProtection['allow'] = {
  selectLocked: true,
  selectUnlocked: true,
  formatCells: false,
  formatColumns: false,
  formatRows: false,
  insertColumns: false,
  insertRows: false,
  insertHyperlinks: false,
  deleteColumns: false,
  deleteRows: false,
  sort: false,
  autoFilter: false,
  pivotTables: false,
  editObjects: false,
  editScenarios: false,
};

export function toggleLockCell(): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  const cur = st.wb.styles.get(sheet.styleIdAt(st.sel.active.r, st.sel.active.c)).locked !== false;
  transact('Lock Cell', (tx) => applyStyleFn(tx, sheet, st.sel.ranges, () => ({ locked: cur ? false : null })));
}

// ---------- notes / comments ----------

export function setNote(r: number, c: number, note: Note | undefined): void {
  const sheet = S().wb.activeSheet;
  transact(note ? 'Edit Note' : 'Delete Note', (tx) => tx.patchCell(sheet, r, c, { note }));
}

export function deleteNotesInSelection(): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  transact('Delete Comment', (tx) => {
    for (const rg of st.sel.ranges) sheet.forEachInRange(rg, (r, c, cell) => cell.note && tx.patchCell(sheet, r, c, { note: undefined }));
  });
}

export function nextNote(dir: 1 | -1): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  const list: { r: number; c: number }[] = [];
  for (const [r, row] of sheet.rows) for (const [c, cell] of row) if (cell.note) list.push({ r, c });
  if (!list.length) return;
  list.sort((a, b) => a.r - b.r || a.c - b.c);
  const a = st.sel.active;
  const idx = list.findIndex((x) => x.r > a.r || (x.r === a.r && x.c > a.c));
  let t;
  if (dir > 0) t = list[idx < 0 ? 0 : idx];
  else {
    const prev = [...list].reverse().find((x) => x.r < a.r || (x.r === a.r && x.c < a.c));
    t = prev ?? list[list.length - 1];
  }
  setState({ sel: rangeSel({ r1: t.r, c1: t.c, r2: t.r, c2: t.c }) });
  openDialog('note', { edit: true });
}

// ---------- hyperlinks ----------

export function setHyperlink(link: string | undefined, text?: string): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  const { r, c } = st.sel.active;
  transact(link ? 'Insert Hyperlink' : 'Remove Hyperlink', (tx) => {
    if (text !== undefined && text !== '') writeInput(tx, sheet, r, c, text);
    else if (link && !sheet.getCell(r, c)?.v) writeInput(tx, sheet, r, c, link);
    const cur = sheet.getCell(r, c) ?? {};
    const s = link ? st.wb.styles.merge(cur.s ?? sheet.styleIdAt(r, c), { fontColor: '#0563C1', underline: 'single' }) : cur.s;
    tx.setCell(sheet, r, c, { ...cur, link, s });
  });
}

// ---------- current region helpers for dialogs ----------

export function tableSourceGuess(): { range: Range; header: boolean } | null {
  const rg = dataRangeForCommand();
  if (!rg) return null;
  const sheet = S().wb.activeSheet;
  return { range: rg, header: guessHeader(sheet, rg) };
}

export function rangeLabel(rg: Range): string {
  return `$${colToName(rg.c1)}$${rg.r1 + 1}:$${colToName(rg.c2)}$${rg.r2 + 1}`;
}

export { addrToA1, currentRegion, bump };

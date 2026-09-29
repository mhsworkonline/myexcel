import { Range } from '../../model/address';
import type { Tx } from '../../model/commands';
import { shiftFormula, tokenize } from '../../model/formula';
import { parseInput } from '../../model/parseInput';
import { cycleActive, primaryRange, Selection, singleSel, step } from '../../model/selection';
import type { Sheet } from '../../model/sheet';
import type { Cell } from '../../model/types';
import { alertBox, bump, EditState, openDialog, requestScroll, S, setState, transact } from '../store';
import { checkValidation } from '../validation';
import { editText } from '../values';

// ---------- protection ----------

export function isCellLocked(sheet: Sheet, r: number, c: number): boolean {
  if (!sheet.protection) return false;
  const st = S().wb.styles.get(sheet.styleIdAt(r, c));
  return st.locked !== false;
}

export function guardRanges(sheet: Sheet, ranges: Range[]): boolean {
  if (!sheet.protection) return true;
  for (const rg of ranges) {
    const rows = Math.min(rg.r2, rg.r1 + 5000);
    const cols = Math.min(rg.c2, rg.c1 + 500);
    for (let r = rg.r1; r <= rows; r++)
      for (let c = rg.c1; c <= cols; c++)
        if (isCellLocked(sheet, r, c)) {
          protectedAlert();
          return false;
        }
  }
  return true;
}

export function protectedAlert(): void {
  alertBox(
    "The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the sheet. You might be requested to enter a password.",
    'MyExcel',
    'warning',
  );
}

export function guardSheet(what: keyof NonNullable<Sheet['protection']>['allow'] | null): boolean {
  const sheet = S().wb.activeSheet;
  if (!sheet.protection) return true;
  if (what && sheet.protection.allow[what]) return true;
  protectedAlert();
  return false;
}

// ---------- formula normalisation ----------

export function normalizeFormula(f: string): string {
  const toks = tokenize(f);
  let depth = 0;
  let inArr = 0;
  let out = '';
  for (const t of toks) {
    if (t.type === 'lparen') depth++;
    if (t.type === 'rparen') depth--;
    if (t.type === 'other' && t.text === '{') inArr++;
    if (t.type === 'other' && t.text === '}') inArr--;
    if (t.type === 'func' || t.type === 'bool') out += t.text.toUpperCase();
    else if (t.type === 'ref') {
      const bang = t.text.lastIndexOf('!');
      out += bang >= 0 ? t.text.slice(0, bang + 1) + t.text.slice(bang + 1).toUpperCase() : t.text.toUpperCase();
    } else out += t.text;
  }
  while (depth-- > 0) out += ')';
  while (inArr-- > 0) out += '}';
  return out;
}

// ---------- writing input ----------

export function buildCellFromInput(sheet: Sheet, r: number, c: number, text: string): { cell: Cell | undefined; autoFormat?: string } {
  const cur = sheet.getCell(r, c);
  const base: Cell = { ...(cur ?? {}) };
  delete base.v;
  delete base.f;
  delete base.e;
  const style = S().wb.styles.get(sheet.styleIdAt(r, c));
  if (style.numFmt === '@' && !text.startsWith('=')) {
    if (text !== '') base.v = text;
    return { cell: base };
  }
  const p = parseInput(text);
  switch (p.kind) {
    case 'empty':
      break;
    case 'formula':
      base.f = normalizeFormula(p.formula!);
      break;
    case 'error':
      base.v = p.value;
      base.e = true;
      break;
    default:
      base.v = p.value;
  }
  const autoFormat = p.autoFormat && (style.numFmt === undefined || style.numFmt === 'General') ? p.autoFormat : undefined;
  return { cell: base, autoFormat };
}

export function writeInput(tx: Tx, sheet: Sheet, r: number, c: number, text: string): void {
  const { cell, autoFormat } = buildCellFromInput(sheet, r, c, text);
  if (!cell) return;
  const wb = S().wb;
  if (autoFormat) cell.s = wb.styles.merge(sheet.styleIdAt(r, c), { numFmt: autoFormat });
  if (text.includes('\n') && !wb.styles.get(cell.s ?? sheet.styleIdAt(r, c)).wrap) {
    cell.s = wb.styles.merge(cell.s ?? sheet.styleIdAt(r, c), { wrap: true });
  }
  tx.setCell(sheet, r, c, cell);
}

/** After formulas are written, give General cells returning dates/percent a matching format (Excel behaviour). */
export function autoFormatFormulaResults(tx: Tx, sheet: Sheet, cells: { r: number; c: number }[]): void {
  tx.flush();
  const eng = S().engine;
  const wb = S().wb;
  if (!eng.attached) return;
  for (const { r, c } of cells.slice(0, 2000)) {
    const cell = sheet.getCell(r, c);
    if (!cell?.f) continue;
    const st = wb.styles.get(sheet.styleIdAt(r, c));
    if (st.numFmt !== undefined) continue;
    const t = eng.valueType(sheet, r, c);
    let fmt: string | undefined;
    if (t === 'NUMBER_DATE') fmt = 'm/d/yyyy';
    else if (t === 'NUMBER_DATETIME') fmt = 'm/d/yyyy h:mm';
    else if (t === 'NUMBER_TIME') fmt = 'h:mm AM/PM';
    else if (t === 'NUMBER_PERCENT') fmt = '0%';
    else if (t === 'NUMBER_CURRENCY') fmt = '"$"#,##0.00';
    if (fmt) tx.setCell(sheet, r, c, { ...cell, s: wb.styles.merge(sheet.styleIdAt(r, c), { numFmt: fmt }) });
  }
}

// ---------- edit session ----------

export function beginEdit(mode: 'enter' | 'edit', initial?: string, fromBar = false): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  const { r, c } = st.sel.active;
  if (isCellLocked(sheet, r, c)) {
    protectedAlert();
    return;
  }
  const text = initial !== undefined ? initial : editText(sheet, r, c);
  const ed: EditState = { sheetId: sheet.id, hostSheetId: sheet.id, r, c, text, caret: text.length, mode, fromBar };
  setState({ edit: ed, menu: null, listMenu: null });
}

export function updateEdit(text: string, caret: number, mode?: 'enter' | 'edit'): void {
  const ed = S().edit;
  if (!ed) return;
  setState({ edit: { ...ed, text, caret, mode: mode ?? ed.mode, point: undefined } });
}

export function cancelEdit(): void {
  const ed = S().edit;
  if (!ed) return;
  const st = S();
  if (st.wb.activeSheetId !== ed.hostSheetId) {
    st.wb.activeSheetId = ed.hostSheetId;
  }
  setState({ edit: null, sel: st.wb.activeSheetId === ed.hostSheetId ? st.sel : singleSel(ed.r, ed.c) });
  bump();
}

export type MoveDir = 'down' | 'up' | 'right' | 'left' | 'none';

let warnedCircular = false;

/** Commits the in-cell edit. Returns false if the edit stays open (invalid formula / validation). */
export function commitEdit(move: MoveDir = 'down', opts: { fillSelection?: boolean; array?: boolean } = {}): boolean {
  const st = S();
  const ed = st.edit;
  if (!ed) return true;
  const wb = st.wb;
  const sheet = wb.sheetById(ed.hostSheetId)!;
  if (wb.activeSheetId !== sheet.id) wb.activeSheetId = sheet.id;
  let text = ed.text;
  const hostSel = st.sel.ranges.length && st.wb.activeSheetId === ed.hostSheetId ? st.sel : singleSel(ed.r, ed.c);
  // Validate formula syntax
  if (text.startsWith('=') && text.length > 1) {
    text = normalizeFormula(text);
    if (!st.engine.validateFormula(text)) {
      alertBox("There's a problem with this formula.\n\nNot trying to type a formula? When the first character is an equal (=) or minus (-) sign, MyExcel thinks it's a formula. To get around this, type an apostrophe (') first.", 'MyExcel', 'error');
      return false;
    }
  }
  // Data validation
  const verdict = checkValidation(sheet, ed.r, ed.c, text);
  if (verdict !== true) {
    openDialog('validationError', { ...verdict, retry: true, text, move, fillSelection: !!opts.fillSelection });
    return false;
  }
  applyCommittedText(sheet, ed.r, ed.c, text, hostSel, opts);
  setState({ edit: null });
  moveAfterCommit(move, hostSel);
  if (text.startsWith('=')) {
    const v = st.engine.getValue(sheet, ed.r, ed.c);
    if (v && typeof v === 'object' && 'circular' in v && v.circular && !warnedCircular) {
      warnedCircular = true;
      alertBox('There are one or more circular references where a formula refers to its own cell either directly or indirectly. This might cause them to calculate incorrectly.\n\nTry removing or changing these references, or moving the formulas to different cells.', 'MyExcel', 'warning');
    }
  }
  return true;
}

/** Writes committed text (used by commitEdit and the validation "warning → Yes" path). */
export function applyCommittedText(sheet: Sheet, r0: number, c0: number, text: string, sel: Selection, opts: { fillSelection?: boolean } = {}): void {
  // Re-committing identical text is a no-op (no undo entry).
  if (!opts.fillSelection && editText(sheet, r0, c0) === text) return;
  transact('Typing', (tx) => {
    const written: { r: number; c: number }[] = [];
    if (opts.fillSelection) {
      for (const rg of sel.ranges) {
        for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 100000); r++) {
          for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 1000); c++) {
            const m = sheet.mergeAt(r, c);
            if (m && (m.r1 !== r || m.c1 !== c)) continue;
            const t = text.startsWith('=') ? shiftFormula(text, r - r0, c - c0) : text;
            writeInput(tx, sheet, r, c, t);
            written.push({ r, c });
          }
        }
      }
    } else {
      writeInput(tx, sheet, r0, c0, text);
      written.push({ r: r0, c: c0 });
    }
    if (text.startsWith('=')) autoFormatFormulaResults(tx, sheet, written);
    autoExtendTables(tx, sheet, r0, c0);
  });
}

// Excel: after Tab, Tab, Enter the active cell returns to the column where tabbing started.
let tabAnchor: { r: number; c: number } | null = null;
export function resetTabAnchor(): void {
  tabAnchor = null;
}

/** Returns the Enter target when a Tab sequence is in progress, updating the anchor. */
export function tabAwareMove(from: { r: number; c: number }, move: MoveDir): { r: number; c: number } | null {
  if (move === 'right') {
    if (!tabAnchor || tabAnchor.r !== from.r) tabAnchor = { ...from };
    return null;
  }
  if (move === 'down' && tabAnchor && tabAnchor.r === from.r) {
    const t = { r: from.r + 1, c: tabAnchor.c };
    tabAnchor = null;
    return t;
  }
  if (move !== 'left') tabAnchor = null;
  return null;
}

function moveAfterCommit(move: MoveDir, sel: Selection): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  if (move === 'none') {
    setState({ sel });
    return;
  }
  const dr = move === 'down' ? 1 : move === 'up' ? -1 : 0;
  const dc = move === 'right' ? 1 : move === 'left' ? -1 : 0;
  const rg = primaryRange(sel);
  const multi = sel.ranges.length > 1 || rg.r1 !== rg.r2 || rg.c1 !== rg.c2;
  let next: Selection;
  if (multi) next = cycleActive(sheet, sel, dr, dc);
  else {
    const back = tabAwareMove(sel.active, move);
    const a = back ?? step(sheet, sel.active, dr, dc);
    next = singleSel(a.r, a.c);
  }
  setState({ sel: next });
  requestScroll(next.active.r, next.active.c);
}

/** Typing directly below or right of a table extends it (Excel AutoExpansion). */
function autoExtendTables(tx: Tx, sheet: Sheet, r: number, c: number): void {
  for (const t of sheet.tables) {
    const rg = t.range;
    if (t.totalRow) continue;
    if (r === rg.r2 + 1 && c >= rg.c1 && c <= rg.c2) {
      const tables = sheet.tables.map((x) => (x.id === t.id ? { ...x, range: { ...x.range, r2: r } } : x));
      tx.setMeta(sheet, 'tables', tables);
      return;
    }
    if (c === rg.c2 + 1 && r >= rg.r1 && r <= rg.r2 && t.headerRow && r === rg.r1) {
      const name = String(sheet.getCell(r, c)?.v ?? `Column${t.columns.length + 1}`);
      const tables = sheet.tables.map((x) =>
        x.id === t.id ? { ...x, range: { ...x.range, c2: c }, columns: [...x.columns, { name }] } : x,
      );
      tx.setMeta(sheet, 'tables', tables);
      return;
    }
  }
}

// ---------- clearing ----------

export type ClearKind = 'all' | 'contents' | 'formats' | 'comments' | 'hyperlinks';

export function clearSelection(kind: ClearKind = 'contents'): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  if (!guardRanges(sheet, st.sel.ranges)) return;
  transact(kind === 'contents' ? 'Clear' : 'Clear ' + kind, (tx) => {
    for (const rg of st.sel.ranges) {
      const full = rg.r1 === 0 && rg.r2 >= 1048575;
      const fullRows = rg.c1 === 0 && rg.c2 >= 16383;
      sheet.forEachInRange(rg, (r, c, cell) => {
        let next: Cell | undefined;
        switch (kind) {
          case 'all':
            next = undefined;
            break;
          case 'contents':
            next = { ...cell };
            delete next.v;
            delete next.f;
            delete next.e;
            break;
          case 'formats':
            next = { ...cell };
            delete next.s;
            break;
          case 'comments':
            next = { ...cell };
            delete next.note;
            break;
          case 'hyperlinks':
            next = { ...cell };
            delete next.link;
            if (next.s !== undefined) {
              const s = st.wb.styles.get(next.s);
              if (s.underline === 'single' && s.fontColor === '#0563C1') next.s = st.wb.styles.merge(next.s, { underline: null, fontColor: null });
            }
            break;
        }
        tx.setCell(sheet, r, c, next);
      });
      if ((kind === 'all' || kind === 'formats') && (full || fullRows)) {
        if (full) {
          const m = new Map(sheet.colStyles);
          for (let c = rg.c1; c <= rg.c2; c++) m.delete(c);
          if (m.size !== sheet.colStyles.size) tx.setMeta(sheet, 'colStyles', m);
        }
        if (fullRows) {
          const m = new Map(sheet.rowStyles);
          for (let r = rg.r1; r <= rg.r2; r++) m.delete(r);
          if (m.size !== sheet.rowStyles.size) tx.setMeta(sheet, 'rowStyles', m);
        }
      }
      if (kind === 'all' || kind === 'formats') {
        const merges = sheet.merges.filter((m) => !(m.r1 >= rg.r1 && m.r2 <= rg.r2 && m.c1 >= rg.c1 && m.c2 <= rg.c2));
        if (merges.length !== sheet.merges.length) tx.setMeta(sheet, 'merges', merges);
      }
    }
  });
}

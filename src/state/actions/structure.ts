import { isFullCols, isFullRows, MAX_COLS, MAX_ROWS, Range } from '../../model/address';
import { FnCommand, StructCommand, Tx } from '../../model/commands';
import { refToRange, refToString, RefPart, tokenize } from '../../model/formula';
import { primaryRange, rangeSel, Selection, singleSel } from '../../model/selection';
import { Sheet } from '../../model/sheet';
import type { Cell } from '../../model/types';
import { cloneSheet } from '../../model/workbook';
import { alertBox, bump, openDialog, requestScroll, S, setState, transact } from '../store';
import { guardSheet } from './edit';

// ---------- rows / columns ----------

function rowSpans(sel: Selection): [number, number][] {
  return sel.ranges.map((rg) => [rg.r1, rg.r2] as [number, number]).sort((a, b) => b[0] - a[0]);
}
function colSpans(sel: Selection): [number, number][] {
  return sel.ranges.map((rg) => [rg.c1, rg.c2] as [number, number]).sort((a, b) => b[0] - a[0]);
}

function checkTablesAndMerges(sheet: Sheet, axis: 'row' | 'col', at: number, n: number, insert: boolean): boolean {
  if (insert) {
    const lastUsed = axis === 'row' ? sheet.usedRange()?.r2 ?? -1 : sheet.usedRange()?.c2 ?? -1;
    const max = axis === 'row' ? MAX_ROWS : MAX_COLS;
    if (lastUsed + n >= max) {
      alertBox("To prevent possible loss of data, MyExcel can't shift nonblank cells off of the worksheet.");
      return false;
    }
  }
  void at;
  return true;
}

export function insertRows(before?: number, count?: number): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  if (!guardSheet('insertRows')) return;
  const spans = before !== undefined ? ([[before, before + (count ?? 1) - 1]] as [number, number][]) : rowSpans(st.sel);
  transact('Insert Rows', (tx) => {
    for (const [r1, r2] of spans) {
      const n = r2 - r1 + 1;
      if (!checkTablesAndMerges(sheet, 'row', r1, n, true)) return false;
      tx.run(new StructCommand('Insert Rows', sheet.id, { axis: 'row', at: r1, delta: n }));
      copyFormatFromNeighbour(tx, sheet, 'row', r1, n);
    }
  });
}

export function deleteRows(): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  if (!guardSheet('deleteRows')) return;
  const spans = rowSpans(st.sel);
  transact('Delete Rows', (tx) => {
    for (const [r1, r2] of spans) {
      tx.run(new StructCommand('Delete Rows', sheet.id, { axis: 'row', at: r1, delta: -(Math.min(r2, MAX_ROWS - 1) - r1 + 1) }));
    }
  });
}

export function insertCols(before?: number, count?: number): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  if (!guardSheet('insertColumns')) return;
  const spans = before !== undefined ? ([[before, before + (count ?? 1) - 1]] as [number, number][]) : colSpans(st.sel);
  transact('Insert Columns', (tx) => {
    for (const [c1, c2] of spans) {
      const n = c2 - c1 + 1;
      if (!checkTablesAndMerges(sheet, 'col', c1, n, true)) return false;
      tx.run(new StructCommand('Insert Columns', sheet.id, { axis: 'col', at: c1, delta: n }));
      copyFormatFromNeighbour(tx, sheet, 'col', c1, n);
    }
  });
}

export function deleteCols(): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  if (!guardSheet('deleteColumns')) return;
  const spans = colSpans(st.sel);
  transact('Delete Columns', (tx) => {
    for (const [c1, c2] of spans) tx.run(new StructCommand('Delete Columns', sheet.id, { axis: 'col', at: c1, delta: -(Math.min(c2, MAX_COLS - 1) - c1 + 1) }));
  });
}

/** Inserted rows/cols take formatting from the row above / column left (Excel default). */
function copyFormatFromNeighbour(tx: Tx, sheet: Sheet, axis: 'row' | 'col', at: number, n: number): void {
  if (at === 0) return;
  const src = at - 1;
  if (axis === 'row') {
    const rs = sheet.rowStyles.get(src);
    if (rs !== undefined) {
      const m = new Map(sheet.rowStyles);
      for (let i = 0; i < n; i++) m.set(at + i, rs);
      tx.setMeta(sheet, 'rowStyles', m);
    }
    const h = sheet.rowHeights.get(src);
    if (h !== undefined) {
      const m = new Map(sheet.rowHeights);
      for (let i = 0; i < n; i++) m.set(at + i, h);
      tx.setMeta(sheet, 'rowHeights', m);
    }
    const row = sheet.rows.get(src);
    if (row) for (const [c, cell] of row) if (cell.s) for (let i = 0; i < n; i++) tx.setCell(sheet, at + i, c, { s: cell.s });
  } else {
    const cs = sheet.colStyles.get(src);
    if (cs !== undefined) {
      const m = new Map(sheet.colStyles);
      for (let i = 0; i < n; i++) m.set(at + i, cs);
      tx.setMeta(sheet, 'colStyles', m);
    }
    const w = sheet.colWidths.get(src);
    if (w !== undefined) {
      const m = new Map(sheet.colWidths);
      for (let i = 0; i < n; i++) m.set(at + i, w);
      tx.setMeta(sheet, 'colWidths', m);
    }
    for (const [r, row] of sheet.rows) {
      const cell = row.get(src);
      if (cell?.s) for (let i = 0; i < n; i++) tx.setCell(sheet, r, at + i, { s: cell.s });
    }
  }
}

// ---------- insert / delete cells with shift ----------

export type ShiftDir = 'right' | 'down' | 'left' | 'up' | 'row' | 'col';

export function insertCellsDialog(): void {
  const rg = primaryRange(S().sel);
  if (isFullRows(rg)) return insertRows();
  if (isFullCols(rg)) return insertCols();
  openDialog('insertCells');
}

export function deleteCellsDialog(): void {
  const rg = primaryRange(S().sel);
  if (isFullRows(rg)) return deleteRows();
  if (isFullCols(rg)) return deleteCols();
  openDialog('deleteCells');
}

export function insertCells(dir: ShiftDir): void {
  if (dir === 'row') return insertRows();
  if (dir === 'col') return insertCols();
  const st = S();
  const sheet = st.wb.activeSheet;
  const rg = primaryRange(st.sel);
  transact('Insert Cells', (tx) => {
    if (dir === 'right') shiftBlock(tx, sheet, { r1: rg.r1, r2: rg.r2, c1: rg.c1, c2: MAX_COLS - 1 }, 0, rg.c2 - rg.c1 + 1);
    else shiftBlock(tx, sheet, { r1: rg.r1, r2: MAX_ROWS - 1, c1: rg.c1, c2: rg.c2 }, rg.r2 - rg.r1 + 1, 0);
  });
}

export function deleteCells(dir: ShiftDir): void {
  if (dir === 'row') return deleteRows();
  if (dir === 'col') return deleteCols();
  const st = S();
  const sheet = st.wb.activeSheet;
  const rg = primaryRange(st.sel);
  transact('Delete Cells', (tx) => {
    // clear the deleted block, invalidate refs into it, then shift the rest in
    markRefsDeleted(tx, sheet, rg);
    sheet.forEachInRange(rg, (r, c) => tx.setCell(sheet, r, c, undefined));
    if (dir === 'left') shiftBlock(tx, sheet, { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: MAX_COLS - 1 }, 0, -(rg.c2 - rg.c1 + 1));
    else shiftBlock(tx, sheet, { r1: rg.r2 + 1, r2: MAX_ROWS - 1, c1: rg.c1, c2: rg.c2 }, -(rg.r2 - rg.r1 + 1), 0);
  });
}

/** Move every cell inside `block` by (dr, dc), retargeting references to moved cells. */
export function shiftBlock(tx: Tx, sheet: Sheet, block: Range, dr: number, dc: number): void {
  const moved: { r: number; c: number; cell: Cell }[] = [];
  sheet.forEachInRange(block, (r, c, cell) => moved.push({ r, c, cell }));
  for (const m of moved) tx.setCell(sheet, m.r, m.c, undefined);
  retargetRefs(tx, sheet, block, dr, dc);
  for (const m of moved) {
    const nr = m.r + dr;
    const nc = m.c + dc;
    if (nr < 0 || nc < 0 || nr >= MAX_ROWS || nc >= MAX_COLS) continue;
    // take the possibly-retargeted formula of the moved cell
    tx.setCell(sheet, nr, nc, m.cell.f ? { ...m.cell, f: retargetFormula(m.cell.f, sheet, sheet, block, dr, dc) } : m.cell);
  }
  const merges = sheet.merges.map((mg) =>
    mg.r1 >= block.r1 && mg.r2 <= block.r2 && mg.c1 >= block.c1 && mg.c2 <= block.c2 ? { r1: mg.r1 + dr, r2: mg.r2 + dr, c1: mg.c1 + dc, c2: mg.c2 + dc } : mg,
  );
  tx.setMeta(sheet, 'merges', merges);
}

function retargetFormula(f: string, host: Sheet, target: Sheet, block: Range, dr: number, dc: number): string {
  const wb = S().wb;
  const toks = tokenize(f);
  let changed = false;
  const out = toks
    .map((t) => {
      if (t.type !== 'ref' || !t.ref) return t.text;
      const rs = t.ref.sheet ? wb.sheetByName(t.ref.sheet) : host;
      if (rs !== target) return t.text;
      const rr = refToRange(t.ref);
      if (rr.r1 < block.r1 || rr.r2 > block.r2 || rr.c1 < block.c1 || rr.c2 > block.c2) return t.text;
      const mv = (p: RefPart): RefPart => ({ ...p, r: p.r !== undefined ? p.r + dr : undefined, c: p.c !== undefined ? p.c + dc : undefined });
      changed = true;
      return refToString({ ...t.ref, a: mv(t.ref.a), b: t.ref.b ? mv(t.ref.b) : undefined });
    })
    .join('');
  return changed ? out : f;
}

function retargetRefs(tx: Tx, target: Sheet, block: Range, dr: number, dc: number): void {
  for (const s of S().wb.sheets) {
    for (const [r, row] of [...s.rows]) {
      for (const [c, cell] of [...row]) {
        if (!cell.f) continue;
        if (s === target && r >= block.r1 && r <= block.r2 && c >= block.c1 && c <= block.c2) continue; // handled when moved
        const nf = retargetFormula(cell.f, s, target, block, dr, dc);
        if (nf !== cell.f) tx.setCell(s, r, c, { ...cell, f: nf });
      }
    }
  }
}

function markRefsDeleted(tx: Tx, target: Sheet, block: Range): void {
  const wb = S().wb;
  for (const s of wb.sheets) {
    for (const [r, row] of [...s.rows]) {
      for (const [c, cell] of [...row]) {
        if (!cell.f) continue;
        const toks = tokenize(cell.f);
        let changed = false;
        const out = toks
          .map((t) => {
            if (t.type !== 'ref' || !t.ref) return t.text;
            const rs = t.ref.sheet ? wb.sheetByName(t.ref.sheet) : s;
            if (rs !== target) return t.text;
            const rr = refToRange(t.ref);
            if (rr.r1 >= block.r1 && rr.r2 <= block.r2 && rr.c1 >= block.c1 && rr.c2 <= block.c2) {
              changed = true;
              return '#REF!';
            }
            return t.text;
          })
          .join('');
        if (changed) tx.setCell(s, r, c, { ...cell, f: out });
      }
    }
  }
}

// ---------- sheets ----------

export function activateSheet(id: string): void {
  const st = S();
  if (st.wb.activeSheetId === id) return;
  const sheetSel = { ...st.sheetSel, [st.wb.activeSheetId]: st.sel };
  if (st.edit) {
    // pointing into another sheet while editing a formula
    if (st.edit.text.startsWith('=')) {
      st.wb.activeSheetId = id;
      setState({ sheetSel, sel: sheetSel[id] ?? singleSel(0, 0), rev: st.rev + 1 });
      return;
    }
    return;
  }
  st.wb.activeSheetId = id;
  setState({ sheetSel, sel: sheetSel[id] ?? singleSel(0, 0), clip: st.clip, selectedChartId: null, rev: st.rev + 1 });
}

export function validSheetName(name: string, except?: Sheet): string | null {
  const n = name.trim();
  if (!n) return 'You typed an invalid name for a sheet or chart.';
  if (n.length > 31) return 'Sheet names cannot exceed 31 characters.';
  if (/[\\/?*[\]:]/.test(n) || n.startsWith("'") || n.endsWith("'")) return 'You typed an invalid name for a sheet or chart. Make sure that:\n• The name that you type does not exceed 31 characters.\n• The name does not contain any of the following characters: \\ / ? * [ or ]\n• You did not leave the name blank.';
  if (n.toLowerCase() === 'history') return '"History" is a reserved name.';
  const ex = S().wb.sheetByName(n);
  if (ex && ex !== except) return 'That name is already taken. Try a different one.';
  return null;
}

export function addSheet(afterIndex?: number): void {
  const st = S();
  const wb = st.wb;
  const idx = afterIndex ?? wb.sheets.indexOf(wb.activeSheet) + 1;
  const sheet = new Sheet(wb.uniqueSheetName());
  const prevActive = wb.activeSheetId;
  transact('Insert Sheet', (tx) => {
    tx.run(
      new FnCommand(
        'Insert Sheet',
        (w) => {
          w.insertSheet(sheet, idx);
          w.activeSheetId = sheet.id;
        },
        (w) => {
          w.removeSheet(sheet);
          w.activeSheetId = prevActive;
        },
      ),
    );
  }, singleSel(0, 0));
  setState({ sheetSel: { ...st.sheetSel, [prevActive]: st.sel } });
}

export function deleteSheet(id?: string, confirmed = false): void {
  const st = S();
  const wb = st.wb;
  const sheet = wb.sheetById(id ?? wb.activeSheetId);
  if (!sheet) return;
  const visible = wb.sheets.filter((s) => s.visibility === 'visible');
  if (visible.length <= 1 && sheet.visibility === 'visible') {
    alertBox('A workbook must contain at least one visible worksheet.\n\nTo hide, delete, or move the selected sheet(s), you must first insert a new sheet or unhide a sheet that is already hidden.');
    return;
  }
  if (sheet.rows.size && !confirmed) {
    openDialog('confirm', {
      message: 'MyExcel will permanently delete this sheet. Do you want to continue?',
      okLabel: 'Delete',
      onOk: () => deleteSheet(sheet.id, true),
    });
    return;
  }
  let index = -1;
  const prevActive = wb.activeSheetId;
  transact('Delete Sheet', (tx) => {
    tx.run(
      new FnCommand(
        'Delete Sheet',
        (w) => {
          index = w.removeSheet(sheet);
        },
        (w) => {
          w.insertSheet(sheet, index);
          w.activeSheetId = prevActive;
        },
      ),
    );
  });
  setState({ sel: S().sheetSel[wb.activeSheetId] ?? singleSel(0, 0) });
}

export function renameSheet(id: string, name: string): boolean {
  const st = S();
  const sheet = st.wb.sheetById(id);
  if (!sheet) return false;
  if (name === sheet.name) return true;
  const err = validSheetName(name, sheet);
  if (err) {
    alertBox(err);
    return false;
  }
  const old = sheet.name;
  transact('Rename Sheet', (tx) => {
    tx.run(new FnCommand('Rename Sheet', (w) => w.renameSheet(w.sheetById(id)!, name.trim()), (w) => w.renameSheet(w.sheetById(id)!, old)));
  });
  return true;
}

export function moveSheet(id: string, toIndex: number): void {
  const wb = S().wb;
  const sheet = wb.sheetById(id);
  if (!sheet) return;
  const from = wb.sheets.indexOf(sheet);
  const to = toIndex > from ? toIndex - 1 : toIndex;
  if (to === from) return;
  transact('Move Sheet', (tx) => {
    tx.run(new FnCommand('Move Sheet', (w) => w.moveSheet(w.sheetById(id)!, to), (w) => w.moveSheet(w.sheetById(id)!, from)));
  });
}

export function copySheet(id: string, toIndex: number): void {
  const wb = S().wb;
  const src = wb.sheetById(id);
  if (!src) return;
  let base = src.name.replace(/ \(\d+\)$/, '');
  let n = 2;
  let name = `${base} (${n})`;
  while (wb.sheetByName(name)) name = `${base} (${++n})`;
  if (name.length > 31) {
    base = base.slice(0, 31 - ` (${n})`.length);
    name = `${base} (${n})`;
  }
  const copy = cloneSheet(src, name);
  const prev = wb.activeSheetId;
  transact('Copy Sheet', (tx) => {
    tx.run(
      new FnCommand(
        'Copy Sheet',
        (w) => {
          w.insertSheet(copy, toIndex);
          w.activeSheetId = copy.id;
        },
        (w) => {
          w.removeSheet(copy);
          w.activeSheetId = prev;
        },
      ),
    );
  }, singleSel(0, 0));
}

export function setTabColor(id: string, color: string | undefined): void {
  const sheet = S().wb.sheetById(id);
  if (!sheet) return;
  transact('Tab Color', (tx) => tx.setMeta(sheet, 'tabColor', color));
}

export function hideSheet(id: string): void {
  const wb = S().wb;
  const sheet = wb.sheetById(id);
  if (!sheet) return;
  if (wb.sheets.filter((s) => s.visibility === 'visible').length <= 1) {
    alertBox('A workbook must contain at least one visible worksheet.');
    return;
  }
  const idx = wb.sheets.indexOf(sheet);
  const next = wb.sheets.slice(idx + 1).concat(wb.sheets.slice(0, idx).reverse()).find((s) => s.visibility === 'visible');
  transact('Hide Sheet', (tx) => {
    tx.setMeta(sheet, 'visibility', 'hidden');
  });
  if (next) activateSheet(next.id);
}

export function unhideSheet(id: string): void {
  const sheet = S().wb.sheetById(id);
  if (!sheet) return;
  transact('Unhide Sheet', (tx) => tx.setMeta(sheet, 'visibility', 'visible'));
  activateSheet(id);
}

// ---------- freeze / split ----------

export function freezePanes(kind: 'panes' | 'topRow' | 'firstCol' | 'unfreeze'): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  let freeze = { rows: 0, cols: 0 };
  if (kind === 'topRow') freeze = { rows: 1, cols: 0 };
  else if (kind === 'firstCol') freeze = { rows: 0, cols: 1 };
  else if (kind === 'panes') {
    const a = st.sel.active;
    freeze = { rows: a.r, cols: a.c };
    if (freeze.rows === 0 && freeze.cols === 0) freeze = { rows: 10, cols: 5 };
  }
  transact('Freeze Panes', (tx) => {
    tx.setMeta(sheet, 'freeze', freeze);
    if (sheet.split) tx.setMeta(sheet, 'split', null);
  });
  requestScroll(0, 0);
}

export function toggleSplit(): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  transact('Split', (tx) => {
    if (sheet.split) tx.setMeta(sheet, 'split', null);
    else {
      tx.setMeta(sheet, 'split', { x: Math.max(1, st.sel.active.c), y: Math.max(1, st.sel.active.r) });
      if (sheet.freeze.rows || sheet.freeze.cols) tx.setMeta(sheet, 'freeze', { rows: 0, cols: 0 });
    }
  });
}

export function selectRangeAndScroll(rg: Range, sheetId?: string): void {
  if (sheetId && sheetId !== S().wb.activeSheetId) activateSheet(sheetId);
  setState({ sel: rangeSel(rg) });
  requestScroll(rg.r1, rg.c1);
  bump();
}

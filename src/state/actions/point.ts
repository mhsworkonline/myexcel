// "Point" mode: building references into a formula by clicking cells or using arrow keys.
import { CellAddr, colToName, normRange, quoteSheetName, rangeToA1 } from '../../model/address';
import { canInsertRefAt } from '../../model/formula';
import { step } from '../../model/selection';
import { S, setState } from '../store';

export function isFormulaEdit(): boolean {
  const ed = S().edit;
  return !!ed && ed.text.startsWith('=');
}

/** Whether a click/arrow should insert a reference rather than commit. */
export function canPoint(): boolean {
  const ed = S().edit;
  if (!ed || !ed.text.startsWith('=')) return false;
  if (ed.point) return ed.point.end === ed.caret;
  return canInsertRefAt(ed.text, ed.caret);
}

function refText(anchor: CellAddr, corner: CellAddr, sheetId: string, fullCols = false, fullRows = false): string {
  const st = S();
  const ed = st.edit!;
  const rg = normRange({ r1: anchor.r, c1: anchor.c, r2: corner.r, c2: corner.c });
  let t: string;
  if (fullCols) t = `${colToName(rg.c1)}:${colToName(rg.c2)}`;
  else if (fullRows) t = `${rg.r1 + 1}:${rg.r2 + 1}`;
  else {
    const sheet = st.wb.sheetById(sheetId)!;
    const m = sheet.mergeAt(rg.r1, rg.c1);
    if (m && rg.r1 === rg.r2 && rg.c1 === rg.c2) t = rangeToA1(m);
    else t = rangeToA1(rg);
  }
  if (sheetId !== ed.hostSheetId) t = quoteSheetName(st.wb.sheetById(sheetId)!.name) + '!' + t;
  return t;
}

/** Insert or replace the pointed reference. */
export function pointTo(anchor: CellAddr, corner: CellAddr, opts: { fullCols?: boolean; fullRows?: boolean } = {}): void {
  const st = S();
  const ed = st.edit;
  if (!ed) return;
  const sheetId = st.wb.activeSheetId;
  const ref = refText(anchor, corner, sheetId, opts.fullCols, opts.fullRows);
  let start: number;
  let end: number;
  if (ed.point && ed.point.end === ed.caret) {
    start = ed.point.start;
    end = ed.point.end;
  } else {
    start = ed.caret;
    end = ed.caret;
  }
  const text = ed.text.slice(0, start) + ref + ed.text.slice(end);
  const caret = start + ref.length;
  setState({ edit: { ...ed, text, caret, point: { start, end: caret, anchor, corner, sheetId } } });
}

/** Arrow keys in point mode move (or with shift extend) the pointed reference. */
export function pointMove(dr: number, dc: number, extend: boolean): void {
  const st = S();
  const ed = st.edit;
  if (!ed) return;
  const sheet = st.wb.activeSheet;
  let anchor: CellAddr;
  let corner: CellAddr;
  if (ed.point && ed.point.end === ed.caret) {
    anchor = ed.point.anchor;
    corner = ed.point.corner;
    corner = step(sheet, corner, dr, dc);
    if (!extend) anchor = corner;
  } else {
    const base = sheet.id === ed.hostSheetId ? { r: ed.r, c: ed.c } : st.sel.active;
    corner = step(sheet, base, dr, dc);
    anchor = corner;
  }
  pointTo(anchor, corner);
}

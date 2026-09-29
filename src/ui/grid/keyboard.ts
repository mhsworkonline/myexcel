// Ready-mode keyboard handling (edit-mode keys live in CellEditor).
import { MAX_COLS, MAX_ROWS } from '../../model/address';
import { FMT } from '../../model/numfmt';
import {
  cycleActive,
  dataEdge,
  expandForMerges,
  extendTo,
  movingCorner,
  primaryRange,
  Selection,
  singleSel,
  step,
} from '../../model/selection';
import { sheetLayout } from '../../model/layout';
import { beginEdit, clearSelection, guardRanges, resetTabAnchor, tabAwareMove, writeInput } from '../../state/actions/edit';
import { clearClip } from '../../state/actions/clipboard';
import { fillDirection, flashFill } from '../../state/actions/fill';
import { currentRegion, toggleAutoFilter } from '../../state/actions/sortFilter';
import {
  applyBorder,
  applyStyle,
  hideRowsCols,
  setNumFmt,
  startFormatPainter,
  toggleStyle,
  toggleUnderline,
} from '../../state/actions/format';
import { activateSheet, addSheet, deleteCellsDialog, insertCellsDialog } from '../../state/actions/structure';
import { newWorkbook, openFile, saveFile } from '../../state/actions/file';
import { autoSum } from '../../state/actions/formulas';
import { bump, openDialog, redo, requestScroll, S, setState, transact, undo } from '../../state/store';
import { editText } from '../../state/values';
import { getScroll } from './geometry';

let lastViewportRows = 30;
let lastViewportCols = 12;
export function setViewportSize(rows: number, cols: number): void {
  lastViewportRows = Math.max(1, rows);
  lastViewportCols = Math.max(1, cols);
}

function setSel(sel: Selection): void {
  setState({ sel });
  const tgt = movingCornerOrActive(sel);
  requestScroll(tgt.r, tgt.c);
}

function movingCornerOrActive(sel: Selection) {
  const rg = primaryRange(sel);
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) return sel.active;
  return movingCorner(sel);
}

function moveBy(dr: number, dc: number, extend: boolean, ctrl: boolean): void {
  resetTabAnchor();
  const st = S();
  const sheet = st.wb.activeSheet;
  const sel = st.sel;
  const ext = extend || st.extendMode;
  if (ext) {
    const corner = movingCorner(sel);
    const to = ctrl ? dataEdge(sheet, corner, dr, dc) : step(sheet, corner, dr, dc);
    // when the anchor sits in a merge, step from the far side
    setSel(extendTo(sheet, sel, to));
    return;
  }
  const from = sel.active;
  const to = ctrl ? dataEdge(sheet, from, dr, dc) : step(sheet, from, dr, dc);
  const m = sheet.mergeAt(to.r, to.c);
  const s = m ? { ranges: [m], active: { r: m.r1, c: m.c1 }, anchor: { r: m.r1, c: m.c1 } } : singleSel(to.r, to.c);
  setSel(s);
}

function pageMove(dir: 1 | -1, horizontal: boolean, extend: boolean): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  const n = horizontal ? lastViewportCols : lastViewportRows;
  const { rows, cols } = sheetLayout(sheet);
  const from = extend ? movingCorner(st.sel) : st.sel.active;
  let r = from.r;
  let c = from.c;
  for (let i = 0; i < n; i++) {
    if (horizontal) c = cols.nextVisible(c, dir);
    else r = rows.nextVisible(r, dir);
  }
  const sc = getScroll(sheet.id);
  if (horizontal) sc.x[1] = Math.max(0, sc.x[1] + dir * cols.def * n * (sheet.zoom / 100));
  else sc.y[1] = Math.max(0, sc.y[1] + dir * rows.def * n * (sheet.zoom / 100));
  if (extend) setSel(extendTo(sheet, st.sel, { r, c }));
  else setSel(singleSel(r, c));
}

function selectAllSmart(): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  const rg = primaryRange(st.sel);
  const region = currentRegion(sheet, st.sel.active.r, st.sel.active.c);
  const isRegion = rg.r1 === region.r1 && rg.r2 === region.r2 && rg.c1 === region.c1 && rg.c2 === region.c2;
  const regionEmpty = region.r1 === region.r2 && region.c1 === region.c2;
  if (!isRegion && !regionEmpty) setState({ sel: { ranges: [region], active: st.sel.active, anchor: { r: region.r1, c: region.c1 } } });
  else setState({ sel: { ranges: [{ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }], active: st.sel.active, anchor: { r: 0, c: 0 } } });
}

function insertNow(kind: 'date' | 'time'): void {
  const d = new Date();
  const text =
    kind === 'date'
      ? `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`
      : `${((d.getHours() + 11) % 12) + 1}:${String(d.getMinutes()).padStart(2, '0')} ${d.getHours() >= 12 ? 'PM' : 'AM'}`;
  const ed = S().edit;
  if (ed) return;
  beginEdit('enter', text);
}

function copyFromAbove(value: boolean): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  const { r, c } = st.sel.active;
  if (r === 0) return;
  if (value) {
    const v = sheet.getCell(r - 1, c);
    beginEdit('edit', v?.f ? String((S().engine.getValue(sheet, r - 1, c) as unknown) ?? '') : editText(sheet, r - 1, c));
  } else beginEdit('edit', editText(sheet, r - 1, c));
}

/** Returns true when the key was handled. */
export function handleGridKey(e: KeyboardEvent): boolean {
  const st = S();
  const ctrl = e.ctrlKey || e.metaKey;
  const shift = e.shiftKey;
  const alt = e.altKey;
  const key = e.key;
  const code = e.code;
  const sheet = st.wb.activeSheet;

  if (st.dialog) return false;

  // ---- navigation ----
  switch (key) {
    case 'ArrowUp':
    case 'ArrowDown':
    case 'ArrowLeft':
    case 'ArrowRight': {
      if (alt && key === 'ArrowDown') {
        // Alt+Down opens validation list / filter dropdown
        setState({ listMenu: { ...st.sel.active } });
        return true;
      }
      const dr = key === 'ArrowUp' ? -1 : key === 'ArrowDown' ? 1 : 0;
      const dc = key === 'ArrowLeft' ? -1 : key === 'ArrowRight' ? 1 : 0;
      moveBy(dr, dc, shift, ctrl);
      return true;
    }
    case 'Tab': {
      if (ctrl) return false;
      const rgT = primaryRange(st.sel);
      if (rgT.r1 === rgT.r2 && rgT.c1 === rgT.c2) tabAwareMove(st.sel.active, shift ? 'left' : 'right');
      setSel(cycleActive(sheet, st.sel, 0, shift ? -1 : 1));
      return true;
    }
    case 'Enter': {
      if (st.painter) return true;
      if (st.clip && !ctrl && !shift) {
        // Enter pastes the clipboard (Excel behaviour)
        import('../../state/actions/clipboard').then((m) => {
          const p = m.getInternalClip();
          if (p) m.pasteFrom({ html: '', text: p.text }, m.DEFAULT_PASTE);
          clearClip();
        });
        return true;
      }
      if (alt) return true;
      const rgE = primaryRange(st.sel);
      if (!shift && rgE.r1 === rgE.r2 && rgE.c1 === rgE.c2) {
        const back = tabAwareMove(st.sel.active, 'down');
        if (back) {
          setSel(singleSel(back.r, back.c));
          return true;
        }
      }
      setSel(cycleActive(sheet, st.sel, shift ? -1 : 1, 0));
      return true;
    }
    case 'Home': {
      const fr = sheet.freeze.rows;
      const fc = sheet.freeze.cols;
      const target = ctrl ? { r: fr, c: fc } : { r: (shift ? movingCorner(st.sel) : st.sel.active).r, c: fc };
      if (shift) setSel(extendTo(sheet, st.sel, target));
      else setSel(singleSel(target.r, target.c));
      return true;
    }
    case 'End': {
      if (ctrl) {
        const used = sheet.usedRange();
        const t = used ? { r: used.r2, c: used.c2 } : { r: 0, c: 0 };
        if (shift) setSel(extendTo(sheet, st.sel, t));
        else setSel(singleSel(t.r, t.c));
      } else setState({ extendMode: false });
      return true;
    }
    case 'PageDown':
    case 'PageUp':
      if (ctrl) {
        const vis = st.wb.sheets.filter((s) => s.visibility === 'visible');
        const i = vis.indexOf(sheet);
        const n = vis[i + (key === 'PageDown' ? 1 : -1)];
        if (n) activateSheet(n.id);
        return true;
      }
      pageMove(key === 'PageDown' ? 1 : -1, alt, shift);
      return true;
    case 'F2':
      if (shift) {
        openDialog('note', { edit: true });
        return true;
      }
      beginEdit('edit');
      return true;
    case 'F3':
      if (shift) {
        openDialog('insertFunction');
        return true;
      }
      openDialog('pasteName');
      return true;
    case 'F4':
      if (ctrl) return false;
      redo();
      return true;
    case 'F5':
      openDialog(ctrl ? 'goto' : 'goto');
      return true;
    case 'F8':
      if (shift) setState({ addMode: !st.addMode, extendMode: false });
      else setState({ extendMode: !st.extendMode, addMode: false });
      return true;
    case 'F9':
      st.engine.recalc();
      bump();
      return true;
    case 'F10':
      if (shift) {
        const el = document.querySelector('[data-grid-root]') as HTMLElement | null;
        const rc = el?.getBoundingClientRect();
        setState({ menu: { x: (rc?.left ?? 0) + 120, y: (rc?.top ?? 0) + 80, kind: 'cell' } });
        return true;
      }
      return false;
    case 'F11':
      if (shift) {
        addSheet();
        return true;
      }
      if (alt) return false;
      openDialog('insertChart', { quick: true });
      return true;
    case 'F12':
      openDialog('saveAs');
      return true;
    case 'Escape':
      if (st.painter) setState({ painter: null });
      clearClip();
      setState({ extendMode: false, addMode: false, listMenu: null, filterMenu: null });
      return true;
    case 'Delete':
      clearSelection('contents');
      return true;
    case 'Backspace': {
      if (!guardRanges(sheet, [primaryRange(st.sel)])) return true;
      beginEdit('enter', '');
      return true;
    }
  }

  if (alt && (key === '=' || code === 'Equal')) {
    autoSum('SUM');
    return true;
  }
  if (shift && !ctrl && key === ' ') {
    const r = primaryRange(st.sel);
    setState({ sel: { ranges: [{ r1: r.r1, r2: r.r2, c1: 0, c2: MAX_COLS - 1 }], active: st.sel.active, anchor: { r: r.r1, c: 0 } } });
    return true;
  }

  if (!ctrl) return false;
  // ---- Ctrl shortcuts ----
  const k = key.toLowerCase();
  if (shift) {
    switch (code) {
      case 'Digit1': setNumFmt('#,##0.00'); return true;
      case 'Digit2': setNumFmt('h:mm AM/PM'); return true;
      case 'Digit3': setNumFmt('d-mmm-yy'); return true;
      case 'Digit4': setNumFmt('"$"#,##0.00_);\\("$"#,##0.00\\)'); return true;
      case 'Digit5': setNumFmt('0%'); return true;
      case 'Digit6': setNumFmt('0.00E+00'); return true;
      case 'Digit7': applyBorder('outside'); return true;
      case 'Digit9': hideRowsCols('row', false); return true;
      case 'Digit0': hideRowsCols('col', false); return true;
      case 'Backquote': setNumFmt('General'); return true;
      case 'Minus': applyBorder('none'); return true;
      case 'Equal': insertCellsDialog(); return true;
      case 'Semicolon': insertNow('time'); return true;
      case 'Quote': copyFromAbove(true); return true;
      case 'KeyL': toggleAutoFilter(); return true;
      case 'KeyF': openDialog('formatCells', { tab: 'Font' }); return true;
      case 'KeyP': openDialog('formatCells', { tab: 'Font' }); return true;
      case 'KeyU': setState({ formulaBarExpanded: !st.formulaBarExpanded }); return true;
      case 'KeyA': openDialog('insertFunction'); return true;
      case 'KeyO': import('../../state/actions/find').then((m) => m.goToSpecial({ kind: 'notes' })); return true;
      case 'Space': selectAllSmart(); return true;
      case 'KeyZ': redo(); return true;
    }
  }
  switch (k) {
    case 'z': undo(); return true;
    case 'y': redo(); return true;
    case 'b': case '2': toggleStyle('bold'); return true;
    case 'i': case '3': toggleStyle('italic'); return true;
    case 'u': case '4': toggleUnderline('single'); return true;
    case '5': toggleStyle('strike'); return true;
    case '1': openDialog('formatCells'); return true;
    case '9': hideRowsCols('row', true); return true;
    case '0': hideRowsCols('col', true); return true;
    case 's': saveFile(); return true;
    case 'o': openFile(); return true;
    case 'n': newWorkbook(); return true;
    case 'p': openDialog('print'); return true;
    case 'f': openDialog('findReplace', { tab: 'find' }); return true;
    case 'h': openDialog('findReplace', { tab: 'replace' }); return true;
    case 'g': openDialog('goto'); return true;
    case 'a': selectAllSmart(); return true;
    case 'd': fillDirection('down'); return true;
    case 'r': fillDirection('right'); return true;
    case 'e': flashFill(); return true;
    case 'k': openDialog('hyperlink'); return true;
    case 't': case 'l': openDialog('createTable'); return true;
    case ';': insertNow('date'); return true;
    case '-': deleteCellsDialog(); return true;
    case '+': case '=': if (shift || key === '+') { insertCellsDialog(); return true; } return false;
    case '`': setState({ showFormulas: !st.showFormulas }); bump(); return true;
    case "'": copyFromAbove(false); return true;
    case ' ': {
      const r = primaryRange(st.sel);
      setState({ sel: { ranges: [{ r1: 0, r2: MAX_ROWS - 1, c1: r.c1, c2: r.c2 }], active: st.sel.active, anchor: { r: 0, c: r.c1 } } });
      return true;
    }
    case 'enter':
      return true;
    case 'f1':
      setState({ ribbonCollapsed: !st.ribbonCollapsed });
      return true;
    case 'c':
    case 'x':
    case 'v':
      return false; // handled by native clipboard events
  }
  return false;
}

/** Typing a printable character in Ready mode starts Enter mode. */
export function startTyping(ch: string): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  const rg = primaryRange(st.sel);
  if (!guardRanges(sheet, [{ r1: st.sel.active.r, c1: st.sel.active.c, r2: st.sel.active.r, c2: st.sel.active.c }])) return;
  void rg;
  beginEdit('enter', ch);
}

export { FMT, applyStyle, startFormatPainter, writeInput, transact, expandForMerges };

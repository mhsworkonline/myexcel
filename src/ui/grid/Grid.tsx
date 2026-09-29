'use client';
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';
import { CellAddr, MAX_COLS, MAX_ROWS, normRange, Range } from '../../model/address';
import { extractRefs, REF_COLORS_EXCEL, shiftFormula } from '../../model/formula';
import { expandForMerges, primaryRange, Selection, singleSel } from '../../model/selection';
import { sheetLayout } from '../../model/layout';
import { beginEdit, cancelEdit, commitEdit, resetTabAnchor } from '../../state/actions/edit';
import { onCopyEvent, onPasteEvent } from '../../state/actions/clipboard';
import { autoFill } from '../../state/actions/fill';
import { applyFormatPainter, autoFitColumns, autoFitRows, setColumnWidth, setRowHeight } from '../../state/actions/format';
import { canPoint, pointTo } from '../../state/actions/point';
import { shiftBlock } from '../../state/actions/structure';
import { bump, S, setState, transact, useStore } from '../../state/store';
import { validationAt } from '../../state/validation';
import { DARK, LIGHT } from '../theme';
import { CellEditor } from './CellEditor';
import { COL_HEADER_CURSOR, COL_RESIZE_CURSOR, CELL_CURSOR, FILL_CURSOR, MOVE_CURSOR, ROW_HEADER_CURSOR, ROW_RESIZE_CURSOR } from './cursors';
import { buildViewport, colX, ensureVisible, getScroll, hitTest, Hit, rangeRect, rowY, Viewport } from './geometry';
import { handleGridKey, setViewportSize, startTyping } from './keyboard';
import { handleEditKey } from './editKeys';
import { renderGrid } from './render';
import { emitScroll, setGridApi } from './scrollBus';
import { ChartLayer } from '../charts/ChartLayer';
import { FilterMenu } from './FilterMenu';
import { ListDropdown } from './ListDropdown';
import { AutoFillOptions } from './AutoFillOptions';
import { NotePopup } from './NotePopup';
import { invalidCells } from '../../state/validation';

type Drag =
  | { kind: 'select'; add: boolean; anchor: CellAddr }
  | { kind: 'colSel'; anchor: number; add: boolean }
  | { kind: 'rowSel'; anchor: number; add: boolean }
  | { kind: 'colResize'; col: number; startX: number; startW: number; cols: number[] }
  | { kind: 'rowResize'; row: number; startY: number; startH: number; rows: number[] }
  | { kind: 'fill'; src: Range; target: Range; ctrl: boolean }
  | { kind: 'point'; anchor: CellAddr; fullCols?: boolean; fullRows?: boolean }
  | { kind: 'painter'; anchor: CellAddr }
  | { kind: 'move'; src: Range; offset: CellAddr; target: Range; copy: boolean }
  | { kind: 'split'; axis: 'x' | 'y' };

let currentVp: Viewport | null = null;
export function getViewport(): Viewport | null {
  return currentVp;
}

let gridRootEl: HTMLElement | null = null;
/** Viewport for event handling; built on demand if no frame has been drawn yet (first load). */
function vpNow(): Viewport | null {
  if (currentVp) return currentVp;
  if (!gridRootEl) return null;
  const r = gridRootEl.getBoundingClientRect();
  const sheet = S().wb.activeSheet;
  currentVp = buildViewport(sheet, Math.max(50, Math.floor(r.width)), Math.max(50, Math.floor(r.height)), getScroll(sheet.id));
  return currentVp;
}

export let invalidOverlay: Set<string> | null = null;
export function toggleInvalidCircles(on: boolean): void {
  if (!on) invalidOverlay = null;
  else invalidOverlay = new Set(invalidCells(S().wb.activeSheet).map((x) => `${x.r},${x.c}`));
  bump();
}

export function Grid() {
  const rootRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const [size, setSize] = useState({ w: 800, h: 400 });
  const sizeRef = useRef(size);
  sizeRef.current = size;
  const dragRef = useRef<Drag | null>(null);
  const fillPreview = useRef<Range | null>(null);
  const lastPointer = useRef<{ x: number; y: number } | null>(null);
  const autoScrollTimer = useRef<number | null>(null);
  const rafRef = useRef<number | null>(null);
  const phaseRef = useRef(0);
  const [editorRect, setEditorRect] = useState<{ x: number; y: number; w: number; h: number } | null>(null);
  const [fillOpts, setFillOpts] = useState<{ x: number; y: number; src: Range; target: Range } | null>(null);
  const [notePos, setNotePos] = useState<{ r: number; c: number; x: number; y: number } | null>(null);
  const [cursor, setCursor] = useState<string>(CELL_CURSOR);
  const [resizeTip, setResizeTip] = useState<{ x: number; y: number; text: string } | null>(null);
  const noteTimer = useRef<number | null>(null);

  const theme = useStore((s) => s.theme);
  const edit = useStore((s) => s.edit);
  const filterMenu = useStore((s) => s.filterMenu);
  const listMenu = useStore((s) => s.listMenu);
  const dialog = useStore((s) => s.dialog);

  // ---------- drawing ----------
  const draw = useCallback(() => {
    rafRef.current = null;
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const st = S();
    const sheet = st.wb.activeSheet;
    const scroll = getScroll(sheet.id);
    const size = sizeRef.current;
    const vp = buildViewport(sheet, size.w, size.h, scroll);
    currentVp = vp;
    const dpr = window.devicePixelRatio || 1;
    if (canvas.width !== Math.round(size.w * dpr) || canvas.height !== Math.round(size.h * dpr)) {
      canvas.width = Math.round(size.w * dpr);
      canvas.height = Math.round(size.h * dpr);
    }
    const dark = st.theme === 'dark';
    // formula reference highlights while editing
    const refs: { range: Range; color: string }[] = [];
    const ed = st.edit;
    if (ed && ed.text.startsWith('=')) {
      for (const ref of extractRefs(ed.text)) {
        const target = ref.sheet ? st.wb.sheetByName(ref.sheet) : st.wb.sheetById(ed.hostSheetId);
        if (target && target.id === sheet.id) refs.push({ range: ref.range, color: REF_COLORS_EXCEL[ref.colorIndex % REF_COLORS_EXCEL.length] });
      }
    }
    const a = st.sel.active;
    const dv = validationAt(sheet, a.r, a.c);
    const listArrow = dv && dv.type === 'list' && dv.showDropdown && !ed ? { r: a.r, c: a.c } : null;
    const pb = st.viewMode === 'pageBreak' || sheet.rowBreaks.length || sheet.colBreaks.length ? pageBreaks(sheet) : null;
    const clip = st.clip && st.clip.sheetId === sheet.id ? { range: st.clip.range, phase: phaseRef.current } : null;
    renderGrid(ctx, {
      sheet,
      wb: st.wb,
      vp,
      pal: dark ? DARK : LIGHT,
      dark,
      dpr,
      sel: st.sel,
      editing: ed && ed.hostSheetId === sheet.id ? { r: ed.r, c: ed.c } : null,
      refs,
      clip,
      fillPreview: fillPreview.current,
      showFormulas: st.showFormulas,
      pageBreaks: pb,
      invalid: invalidOverlay,
      filterButtons: true,
      listArrow,
      painterRange: st.painter && st.painter.sheetId === sheet.id ? st.painter.range : null,
    });
    // editor rectangle
    if (ed && ed.hostSheetId === sheet.id) {
      const m = sheet.mergeAt(ed.r, ed.c);
      const rr = m ? rangeRect(vp, m.r1, m.c1, m.r2, m.c2) : rangeRect(vp, ed.r, ed.c, ed.r, ed.c);
      setEditorRect((old) => (old && rr && old.x === rr.x && old.y === rr.y && old.w === rr.w && old.h === rr.h ? old : rr));
    } else setEditorRect((old) => (old ? null : old));
    const mainRows = vp.rowPanes[vp.rowPanes.length - 1];
    const mainCols = vp.colPanes[vp.colPanes.length - 1];
    setViewportSize(Math.floor(mainRows.size / (20 * vp.z)), Math.floor(mainCols.size / (64 * vp.z)));
  }, []);

  // Redraw synchronously when the element is resized so the bitmap never shows stretched.
  useLayoutEffect(() => {
    if (rafRef.current !== null) {
      cancelAnimationFrame(rafRef.current);
      rafRef.current = null;
    }
    draw();
  }, [size, draw]);

  const schedule = useCallback(() => {
    if (rafRef.current === null) rafRef.current = requestAnimationFrame(draw);
  }, [draw]);

  useEffect(() => {
    schedule();
    const unsub = useStore.subscribe(() => schedule());
    return () => {
      unsub();
    };
  }, [schedule]);

  // marching ants
  useEffect(() => {
    const id = window.setInterval(() => {
      const st = S();
      if (st.clip || st.painter) {
        phaseRef.current = (phaseRef.current + 1) % 8;
        schedule();
      }
    }, 90);
    return () => clearInterval(id);
  }, [schedule]);

  // size
  useLayoutEffect(() => {
    const el = rootRef.current;
    if (!el) return;
    gridRootEl = el;
    const ro = new ResizeObserver(() => {
      const r = el.getBoundingClientRect();
      setSize({ w: Math.max(50, Math.floor(r.width)), h: Math.max(50, Math.floor(r.height)) });
    });
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  // ---------- scrolling ----------
  const extent = useCallback(
    (axis: 'x' | 'y') => {
      const st = S();
      const sheet = st.wb.activeSheet;
      const sc = getScroll(sheet.id);
      const vp = currentVp ?? buildViewport(sheet, sizeRef.current.w, sizeRef.current.h, sc);
      const z = vp.z;
      const used = sheet.usedRange();
      const pane = axis === 'x' ? vp.colPanes[vp.colPanes.length - 1] : vp.rowPanes[vp.rowPanes.length - 1];
      const L = axis === 'x' ? vp.cols : vp.rows;
      const base = vp.frozen ? L.offset(pane.first) * z : 0;
      const usedEnd = used ? L.offset((axis === 'x' ? used.c2 : used.r2) + 1) * z - base : 0;
      const sel = st.sel.active;
      const selEnd = L.offset((axis === 'x' ? sel.c : sel.r) + 1) * z - base;
      const pos = axis === 'x' ? sc.x[1] : sc.y[1];
      const view = pane.size;
      const hardMax = L.offset(L.count) * z - base;
      const total = Math.min(hardMax, Math.max(usedEnd + view * 0.5, selEnd + view * 0.5, pos + view * 1.2, view * 2));
      return { pos, view, total };
    },
    [],
  );

  const setScrollAxis = useCallback(
    (axis: 'x' | 'y', value: number) => {
      const sheet = S().wb.activeSheet;
      const sc = getScroll(sheet.id);
      const vp = vpNow();
      const L = axis === 'x' ? sheetLayout(sheet).cols : sheetLayout(sheet).rows;
      const z = sheet.zoom / 100;
      const max = L.offset(L.count) * z - (vp ? (axis === 'x' ? vp.colPanes[vp.colPanes.length - 1].size : vp.rowPanes[vp.rowPanes.length - 1].size) : 0);
      const v = Math.max(0, Math.min(max, value));
      if (axis === 'x') sc.x[1] = v;
      else sc.y[1] = v;
      schedule();
      emitScroll();
    },
    [schedule],
  );

  useEffect(() => {
    setGridApi({ setScroll: setScrollAxis, extent, redraw: schedule });
    return () => setGridApi(null);
  }, [setScrollAxis, extent, schedule]);

  // scroll requests from actions
  const scrollTo = useStore((s) => s.scrollTo);
  useEffect(() => {
    if (!scrollTo) return;
    const sheet = S().wb.activeSheet;
    const sc = getScroll(sheet.id);
    const vp = buildViewport(sheet, size.w, size.h, sc);
    if (ensureVisible(vp, sc, scrollTo.r, scrollTo.c)) {
      schedule();
      emitScroll();
    }
  }, [scrollTo, size, schedule]);

  const onWheel = useCallback(
    (e: WheelEvent) => {
      if (e.ctrlKey) {
        e.preventDefault();
        const sheet = S().wb.activeSheet;
        const z = Math.max(10, Math.min(400, sheet.zoom + (e.deltaY < 0 ? 10 : -10)));
        sheet.zoom = z;
        sheet.touch();
        bump();
        return;
      }
      e.preventDefault();
      const sheet = S().wb.activeSheet;
      const sc = getScroll(sheet.id);
      let dx = e.deltaX;
      let dy = e.deltaY;
      if (e.deltaMode === 1) {
        dx *= 20;
        dy *= 20 * 1;
      }
      if (e.shiftKey && !dx) {
        dx = dy;
        dy = 0;
      }
      // Excel scrolls 3 rows per notch
      if (e.deltaMode === 0 && Math.abs(dy) >= 100 && Number.isInteger(dy)) dy = Math.sign(dy) * 60 * (sheet.zoom / 100);
      const vp = vpNow();
      const splitPane = vp && vp.split && lastPointer.current ? vp : null;
      if (splitPane && lastPointer.current) {
        const { x, y } = lastPointer.current;
        const inTop = vp!.rowPanes.length > 1 && y < vp!.rowPanes[1].start;
        const inLeft = vp!.colPanes.length > 1 && x < vp!.colPanes[1].start;
        if (dy) {
          const i = inTop ? 0 : 1;
          sc.y[i] = Math.max(0, sc.y[i] + dy);
        }
        if (dx) {
          const i = inLeft ? 0 : 1;
          sc.x[i] = Math.max(0, sc.x[i] + dx);
        }
        schedule();
        emitScroll();
        return;
      }
      if (dy) setScrollAxis('y', sc.y[1] + dy);
      if (dx) setScrollAxis('x', sc.x[1] + dx);
    },
    [schedule, setScrollAxis],
  );

  useEffect(() => {
    const el = rootRef.current;
    if (!el) return;
    el.addEventListener('wheel', onWheel, { passive: false });
    return () => el.removeEventListener('wheel', onWheel);
  }, [onWheel]);

  // ---------- helpers ----------
  const localXY = (e: { clientX: number; clientY: number }) => {
    const r = rootRef.current!.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  };

  const focusInput = () => {
    if (!S().edit && !S().dialog) inputRef.current?.focus({ preventScroll: true });
  };

  const fillHandleHit = (x: number, y: number): boolean => {
    const vp = vpNow();
    const st = S();
    if (!vp || st.sel.ranges.length !== 1) return false;
    const rg = primaryRange(st.sel);
    const R = rangeRect(vp, rg.r1, rg.c1, Math.min(rg.r2, MAX_ROWS - 1), Math.min(rg.c2, MAX_COLS - 1));
    if (!R) return false;
    const hx = R.x + R.w - 1;
    const hy = R.y + R.h - 1;
    return Math.abs(x - hx) <= 5 && Math.abs(y - hy) <= 5;
  };

  const selectionBorderHit = (x: number, y: number, hit: Hit): boolean => {
    const vp = vpNow();
    const st = S();
    if (!vp || st.sel.ranges.length !== 1 || hit.area !== 'cell') return false;
    const rg = primaryRange(st.sel);
    if (rg.r2 - rg.r1 > 100000 || rg.c2 - rg.c1 > 10000) return false;
    const R = rangeRect(vp, rg.r1, rg.c1, rg.r2, rg.c2);
    if (!R) return false;
    const near = (a: number, b: number) => Math.abs(a - b) <= 3;
    const inY = y >= R.y - 3 && y <= R.y + R.h + 3;
    const inX = x >= R.x - 3 && x <= R.x + R.w + 3;
    return (inY && (near(x, R.x) || near(x, R.x + R.w))) || (inX && (near(y, R.y) || near(y, R.y + R.h)));
  };

  const filterButtonHit = (x: number, y: number, hit: Hit): number | null => {
    const st = S();
    const sheet = st.wb.activeSheet;
    const vp = vpNow();
    if (!vp || hit.area !== 'cell') return null;
    const check = (rg: Range, c: number) => {
      const R = rangeRect(vp, rg.r1, c, rg.r1, c);
      if (!R) return false;
      const s = Math.round(16 * vp.z);
      return x >= R.x + R.w - s - 2 && x <= R.x + R.w && y >= R.y + R.h - s - 2 && y <= R.y + R.h;
    };
    const af = sheet.autoFilter;
    if (af && hit.r === af.range.r1 && hit.c >= af.range.c1 && hit.c <= af.range.c2 && check(af.range, hit.c)) return hit.c;
    for (const t of sheet.tables)
      if (t.headerRow && t.showFilterButton && hit.r === t.range.r1 && hit.c >= t.range.c1 && hit.c <= t.range.c2 && check(t.range, hit.c)) {
        if (!af || af.range.r1 !== t.range.r1 || af.range.c1 !== t.range.c1) {
          // attach an autofilter to the table on first use
          transact('Filter', (tx) => tx.setMeta(sheet, 'autoFilter', { range: { ...t.range, r2: t.totalRow ? t.range.r2 - 1 : t.range.r2 }, filters: {} }));
        }
        return hit.c;
      }
    return null;
  };

  const listArrowHit = (x: number, y: number): boolean => {
    const st = S();
    const sheet = st.wb.activeSheet;
    const vp = vpNow();
    if (!vp) return false;
    const a = st.sel.active;
    const dv = validationAt(sheet, a.r, a.c);
    if (!dv || dv.type !== 'list' || !dv.showDropdown) return false;
    const m = sheet.mergeAt(a.r, a.c);
    const R = m ? rangeRect(vp, m.r1, m.c1, m.r2, m.c2) : rangeRect(vp, a.r, a.c, a.r, a.c);
    if (!R) return false;
    const s = Math.min(R.h, 17 * vp.z);
    return x >= R.x + R.w && x <= R.x + R.w + s + 1 && y >= R.y + R.h - s && y <= R.y + R.h;
  };

  const updateDrag = useCallback(
    (x: number, y: number, e?: { ctrlKey?: boolean; altKey?: boolean }) => {
      const d = dragRef.current;
      const vp = vpNow();
      if (!d || !vp) return;
      const st = S();
      const sheet = st.wb.activeSheet;
      const hit = hitTest(vp, Math.max(vp.hw + 1, Math.min(vp.width - 1, x)), Math.max(vp.hh + 1, Math.min(vp.height - 1, y)));
      switch (d.kind) {
        case 'select': {
          const rg = expandForMerges(sheet, normRange({ r1: d.anchor.r, c1: d.anchor.c, r2: hit.r, c2: hit.c }));
          const ranges = d.add ? st.sel.ranges.slice(0, -1).concat([rg]) : [rg];
          setState({ sel: { ranges, active: st.sel.active, anchor: d.anchor } });
          break;
        }
        case 'painter': {
          const rg = normRange({ r1: d.anchor.r, c1: d.anchor.c, r2: hit.r, c2: hit.c });
          setState({ sel: { ranges: [rg], active: d.anchor, anchor: d.anchor } });
          break;
        }
        case 'colSel': {
          const c1 = Math.min(d.anchor, hit.c);
          const c2 = Math.max(d.anchor, hit.c);
          const rg = { r1: 0, r2: MAX_ROWS - 1, c1, c2 };
          const ranges = d.add ? st.sel.ranges.slice(0, -1).concat([rg]) : [rg];
          setState({ sel: { ranges, active: st.sel.active, anchor: { r: 0, c: d.anchor } } });
          break;
        }
        case 'rowSel': {
          const r1 = Math.min(d.anchor, hit.r);
          const r2 = Math.max(d.anchor, hit.r);
          const rg = { r1, r2, c1: 0, c2: MAX_COLS - 1 };
          const ranges = d.add ? st.sel.ranges.slice(0, -1).concat([rg]) : [rg];
          setState({ sel: { ranges, active: st.sel.active, anchor: { r: d.anchor, c: 0 } } });
          break;
        }
        case 'colResize': {
          const w = Math.max(0, d.startW + (x - d.startX) / vp.z);
          const tmp = new Map(sheet.colWidths);
          tmp.set(d.col, Math.round(w));
          sheet.colWidths = tmp;
          sheet.touch();
          setResizeTip({ x, y: 4, text: `Width: ${((w - 5) / 7).toFixed(2)} (${Math.round(w)} pixels)` });
          schedule();
          break;
        }
        case 'rowResize': {
          const h = Math.max(0, d.startH + (y - d.startY) / vp.z);
          const tmp = new Map(sheet.rowHeights);
          tmp.set(d.row, Math.round(h));
          sheet.rowHeights = tmp;
          sheet.touch();
          setResizeTip({ x: 4, y, text: `Height: ${(h * 0.75).toFixed(2)} (${Math.round(h)} pixels)` });
          schedule();
          break;
        }
        case 'fill': {
          const src = d.src;
          // choose the axis with the larger displacement
          const dRow = hit.r > src.r2 ? hit.r - src.r2 : hit.r < src.r1 ? hit.r - src.r1 : 0;
          const dCol = hit.c > src.c2 ? hit.c - src.c2 : hit.c < src.c1 ? hit.c - src.c1 : 0;
          const rowPx = Math.abs(dRow) * 20;
          const colPx = Math.abs(dCol) * 64;
          let target: Range = { ...src };
          if (rowPx >= colPx && dRow !== 0) target = dRow > 0 ? { ...src, r2: hit.r } : { ...src, r1: hit.r };
          else if (dCol !== 0) target = dCol > 0 ? { ...src, c2: hit.c } : { ...src, c1: hit.c };
          else if (hit.r < src.r2 && hit.r >= src.r1) target = { ...src, r2: hit.r };
          else if (hit.c < src.c2 && hit.c >= src.c1) target = { ...src, c2: hit.c };
          d.target = target;
          d.ctrl = !!e?.ctrlKey;
          fillPreview.current = target;
          schedule();
          break;
        }
        case 'point': {
          if (d.fullCols) pointTo({ r: 0, c: d.anchor.c }, { r: MAX_ROWS - 1, c: hit.c }, { fullCols: true });
          else if (d.fullRows) pointTo({ r: d.anchor.r, c: 0 }, { r: hit.r, c: MAX_COLS - 1 }, { fullRows: true });
          else pointTo(d.anchor, { r: hit.r, c: hit.c });
          break;
        }
        case 'move': {
          const h = d.src.r2 - d.src.r1;
          const w = d.src.c2 - d.src.c1;
          const r1 = Math.max(0, hit.r - d.offset.r);
          const c1 = Math.max(0, hit.c - d.offset.c);
          d.target = { r1, c1, r2: r1 + h, c2: c1 + w };
          d.copy = !!e?.ctrlKey;
          fillPreview.current = d.target;
          schedule();
          break;
        }
        case 'split': {
          const sheetNow = S().wb.activeSheet;
          const sp = sheetNow.split!;
          const next = d.axis === 'x' ? { ...sp, x: hit.c } : { ...sp, y: hit.r };
          sheetNow.split = next;
          sheetNow.touch();
          schedule();
          break;
        }
      }
    },
    [schedule],
  );

  const stopAutoScroll = () => {
    if (autoScrollTimer.current !== null) {
      clearInterval(autoScrollTimer.current);
      autoScrollTimer.current = null;
    }
  };

  const startAutoScroll = () => {
    if (autoScrollTimer.current !== null) return;
    autoScrollTimer.current = window.setInterval(() => {
      const p = lastPointer.current;
      const vp = vpNow();
      const d = dragRef.current;
      if (!p || !vp || !d || d.kind === 'colResize' || d.kind === 'rowResize' || d.kind === 'split') return;
      const sheet = S().wb.activeSheet;
      const sc = getScroll(sheet.id);
      let moved = false;
      const main = vp.rowPanes[vp.rowPanes.length - 1];
      const mainC = vp.colPanes[vp.colPanes.length - 1];
      if (p.y > vp.height - 2) {
        sc.y[1] += Math.min(200, (p.y - vp.height + 10) * 2);
        moved = true;
      } else if (p.y < main.start && sc.y[1] > 0 && d.kind !== 'colSel') {
        sc.y[1] = Math.max(0, sc.y[1] - Math.min(200, (main.start - p.y + 10) * 2));
        moved = true;
      }
      if (p.x > vp.width - 2) {
        sc.x[1] += Math.min(200, (p.x - vp.width + 10) * 2);
        moved = true;
      } else if (p.x < mainC.start && sc.x[1] > 0 && d.kind !== 'rowSel') {
        sc.x[1] = Math.max(0, sc.x[1] - Math.min(200, (mainC.start - p.x + 10) * 2));
        moved = true;
      }
      if (moved) {
        draw();
        updateDrag(p.x, p.y);
        emitScroll();
      }
    }, 50);
  };

  // ---------- pointer events ----------
  const onPointerDown = (e: React.PointerEvent) => {
    if (e.button === 2) return; // context menu handles
    const vp = vpNow();
    if (!vp) return;
    const { x, y } = localXY(e);
    lastPointer.current = { x, y };
    const st = S();
    const sheet = st.wb.activeSheet;
    const hit = hitTest(vp, x, y);
    setFillOpts(null);
    setState({ menu: null, selectedChartId: null });
    if (e.button !== 0) return;
    const capture = () => (e.target as HTMLElement).setPointerCapture?.(e.pointerId);

    if (hit.area === 'splitV' || hit.area === 'splitH') {
      dragRef.current = { kind: 'split', axis: hit.area === 'splitV' ? 'x' : 'y' };
      capture();
      return;
    }

    // Editing a formula: clicks insert references
    if (st.edit) {
      if (canPoint() && (hit.area === 'cell' || hit.area === 'colHeader' || hit.area === 'rowHeader')) {
        e.preventDefault();
        if (hit.area === 'colHeader') {
          dragRef.current = { kind: 'point', anchor: { r: 0, c: hit.c }, fullCols: true };
          pointTo({ r: 0, c: hit.c }, { r: MAX_ROWS - 1, c: hit.c }, { fullCols: true });
        } else if (hit.area === 'rowHeader') {
          dragRef.current = { kind: 'point', anchor: { r: hit.r, c: 0 }, fullRows: true };
          pointTo({ r: hit.r, c: 0 }, { r: hit.r, c: MAX_COLS - 1 }, { fullRows: true });
        } else {
          dragRef.current = { kind: 'point', anchor: { r: hit.r, c: hit.c } };
          pointTo({ r: hit.r, c: hit.c }, { r: hit.r, c: hit.c });
        }
        capture();
        return;
      }
      if (st.edit.hostSheetId !== sheet.id || !commitEdit('none')) {
        e.preventDefault();
        return;
      }
    }

    if (hit.area === 'corner') {
      setState({ sel: { ranges: [{ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }], active: { r: 0, c: 0 }, anchor: { r: 0, c: 0 } } });
      focusInput();
      return;
    }
    if (hit.area === 'colHeader') {
      if (hit.resize !== undefined) {
        const sc = new Set<number>();
        for (const rg of st.sel.ranges) if (rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1 && hit.resize >= rg.c1 && hit.resize <= rg.c2) for (let c = rg.c1; c <= rg.c2; c++) sc.add(c);
        dragRef.current = { kind: 'colResize', col: hit.resize, startX: x, startW: sheet.colWidth(hit.resize), cols: sc.size ? [...sc] : [hit.resize] };
        (dragRef.current as { orig?: Map<number, number> }).orig = sheet.colWidths;
      } else {
        const add = e.ctrlKey || e.metaKey;
        if (e.shiftKey) {
          const a = st.sel.anchor.c;
          const rg = { r1: 0, r2: MAX_ROWS - 1, c1: Math.min(a, hit.c), c2: Math.max(a, hit.c) };
          setState({ sel: { ranges: [rg], active: st.sel.active, anchor: st.sel.anchor } });
          dragRef.current = { kind: 'colSel', anchor: a, add: false };
        } else {
          const rg = { r1: 0, r2: MAX_ROWS - 1, c1: hit.c, c2: hit.c };
          const vr = vp.rowPanes[vp.rowPanes.length - 1];
          const topRow = vp.rows.indexAt(vr.content / vp.z);
          setState({ sel: { ranges: add ? st.sel.ranges.concat([rg]) : [rg], active: { r: Math.max(vr.first, topRow), c: hit.c }, anchor: { r: 0, c: hit.c } } });
          dragRef.current = { kind: 'colSel', anchor: hit.c, add };
        }
      }
      capture();
      focusInput();
      return;
    }
    if (hit.area === 'rowHeader') {
      if (hit.resize !== undefined) {
        const sr = new Set<number>();
        for (const rg of st.sel.ranges) if (rg.c1 === 0 && rg.c2 >= MAX_COLS - 1 && hit.resize >= rg.r1 && hit.resize <= rg.r2) for (let r = rg.r1; r <= rg.r2; r++) sr.add(r);
        dragRef.current = { kind: 'rowResize', row: hit.resize, startY: y, startH: sheet.rowHeight(hit.resize), rows: sr.size ? [...sr] : [hit.resize] };
        (dragRef.current as { orig?: Map<number, number> }).orig = sheet.rowHeights;
      } else {
        const add = e.ctrlKey || e.metaKey;
        if (e.shiftKey) {
          const a = st.sel.anchor.r;
          const rg = { r1: Math.min(a, hit.r), r2: Math.max(a, hit.r), c1: 0, c2: MAX_COLS - 1 };
          setState({ sel: { ranges: [rg], active: st.sel.active, anchor: st.sel.anchor } });
          dragRef.current = { kind: 'rowSel', anchor: a, add: false };
        } else {
          const rg = { r1: hit.r, r2: hit.r, c1: 0, c2: MAX_COLS - 1 };
          const vc = vp.colPanes[vp.colPanes.length - 1];
          const leftCol = vp.cols.indexAt(vc.content / vp.z);
          setState({ sel: { ranges: add ? st.sel.ranges.concat([rg]) : [rg], active: { r: hit.r, c: Math.max(vc.first, leftCol) }, anchor: { r: hit.r, c: 0 } } });
          dragRef.current = { kind: 'rowSel', anchor: hit.r, add };
        }
      }
      capture();
      focusInput();
      return;
    }
    if (hit.area !== 'cell') return;

    // filter button / list arrow
    const fcol = filterButtonHit(x, y, hit);
    if (fcol !== null) {
      const R = rangeRect(vp, hit.r, fcol, hit.r, fcol)!;
      const root = rootRef.current!.getBoundingClientRect();
      setState({ filterMenu: { col: fcol, x: root.left + R.x + R.w - 16, y: root.top + R.y + R.h } });
      return;
    }
    if (listArrowHit(x, y)) {
      setState({ listMenu: { ...st.sel.active } });
      return;
    }
    if (st.painter) {
      dragRef.current = { kind: 'painter', anchor: { r: hit.r, c: hit.c } };
      setState({ sel: singleSel(hit.r, hit.c) });
      capture();
      return;
    }
    if (fillHandleHit(x, y)) {
      const src = primaryRange(st.sel);
      dragRef.current = { kind: 'fill', src, target: src, ctrl: e.ctrlKey };
      capture();
      return;
    }
    if (selectionBorderHit(x, y, hit)) {
      const src = primaryRange(st.sel);
      dragRef.current = { kind: 'move', src, offset: { r: hit.r - src.r1, c: hit.c - src.c1 }, target: src, copy: e.ctrlKey };
      capture();
      return;
    }
    resetTabAnchor();
    const add = e.ctrlKey || e.metaKey || st.addMode;
    const m = sheet.mergeAt(hit.r, hit.c);
    const cell = m ? { r: m.r1, c: m.c1 } : { r: hit.r, c: hit.c };
    if (e.shiftKey || st.extendMode) {
      const rg = expandForMerges(sheet, normRange({ r1: st.sel.anchor.r, c1: st.sel.anchor.c, r2: hit.r, c2: hit.c }));
      setState({ sel: { ranges: st.sel.ranges.slice(0, -1).concat([rg]), active: st.sel.active, anchor: st.sel.anchor } });
      dragRef.current = { kind: 'select', add: true, anchor: st.sel.anchor };
    } else {
      const rg = m ?? { r1: cell.r, c1: cell.c, r2: cell.r, c2: cell.c };
      const sel: Selection = { ranges: add ? st.sel.ranges.concat([rg]) : [rg], active: cell, anchor: cell };
      setState({ sel, listMenu: null });
      dragRef.current = { kind: 'select', add, anchor: cell };
    }
    capture();
    focusInput();
    // hyperlink: follow on click (after a short hold it just selects, like Excel)
    const link = sheet.getCell(cell.r, cell.c)?.link;
    if (link && !e.shiftKey && !add) {
      const t0 = Date.now();
      const up = () => {
        window.removeEventListener('pointerup', up);
        if (Date.now() - t0 < 350 && dragRef.current === null) openLink(link);
      };
      window.addEventListener('pointerup', up);
    }
  };

  const onPointerMove = (e: React.PointerEvent) => {
    const vp = vpNow();
    if (!vp) return;
    const { x, y } = localXY(e);
    lastPointer.current = { x, y };
    if (dragRef.current) {
      updateDrag(x, y, e);
      if (x < 0 || y < vp.hh || x > vp.width || y > vp.height || x < vp.hw || y < vp.rowPanes[vp.rowPanes.length - 1].start) startAutoScroll();
      else if (x > vp.hw && y > vp.hh && x < vp.width - 2 && y < vp.height - 2) stopAutoScroll();
      return;
    }
    const hit = hitTest(vp, x, y);
    let cur = CELL_CURSOR;
    if (hit.area === 'colHeader') cur = hit.resize !== undefined ? COL_RESIZE_CURSOR : COL_HEADER_CURSOR;
    else if (hit.area === 'rowHeader') cur = hit.resize !== undefined ? ROW_RESIZE_CURSOR : ROW_HEADER_CURSOR;
    else if (hit.area === 'corner') cur = CELL_CURSOR;
    else if (hit.area === 'splitV') cur = COL_RESIZE_CURSOR;
    else if (hit.area === 'splitH') cur = ROW_RESIZE_CURSOR;
    else if (hit.area === 'cell') {
      if (fillHandleHit(x, y)) cur = FILL_CURSOR;
      else if (selectionBorderHit(x, y, hit) && !S().edit) cur = MOVE_CURSOR;
      else if (filterButtonHit(x, y, hit) !== null || listArrowHit(x, y)) cur = 'default';
      else if (S().wb.activeSheet.getCell(hit.r, hit.c)?.link) cur = 'pointer';
      // note hover
      const sheet = S().wb.activeSheet;
      const m = sheet.mergeAt(hit.r, hit.c);
      const nr = m ? m.r1 : hit.r;
      const nc = m ? m.c1 : hit.c;
      if (sheet.getCell(nr, nc)?.note) {
        if (!notePos || notePos.r !== nr || notePos.c !== nc) {
          if (noteTimer.current) clearTimeout(noteTimer.current);
          noteTimer.current = window.setTimeout(() => {
            const R = rangeRect(vp, nr, nc, m ? m.r2 : nr, m ? m.c2 : nc);
            if (R) setNotePos({ r: nr, c: nc, x: R.x + R.w + 8, y: R.y });
          }, 250);
        }
      } else if (notePos) {
        if (noteTimer.current) clearTimeout(noteTimer.current);
        setNotePos(null);
      }
    }
    if (cur !== cursor) setCursor(cur);
  };

  const onPointerUp = (e: React.PointerEvent) => {
    stopAutoScroll();
    const d = dragRef.current;
    dragRef.current = null;
    setResizeTip(null);
    if (!d) return;
    const st = S();
    const sheet = st.wb.activeSheet;
    switch (d.kind) {
      case 'colResize': {
        const w = sheet.colWidths.get(d.col) ?? sheet.colWidth(d.col);
        sheet.colWidths = (d as unknown as { orig: Map<number, number> }).orig;
        sheet.touch();
        setColumnWidth(w, d.cols);
        break;
      }
      case 'rowResize': {
        const h = sheet.rowHeights.get(d.row) ?? sheet.rowHeight(d.row);
        sheet.rowHeights = (d as unknown as { orig: Map<number, number> }).orig;
        sheet.touch();
        setRowHeight(h, d.rows);
        break;
      }
      case 'fill': {
        fillPreview.current = null;
        const t = d.target;
        if (t.r1 !== d.src.r1 || t.r2 !== d.src.r2 || t.c1 !== d.src.c1 || t.c2 !== d.src.c2) {
          autoFill(d.src, t, 'auto', d.ctrl);
          const vp = vpNow();
          if (vp) {
            const R = rangeRect(vp, t.r2, t.c2, t.r2, t.c2);
            if (R) setFillOpts({ x: R.x + R.w + 2, y: R.y + R.h + 2, src: d.src, target: t });
          }
        }
        schedule();
        break;
      }
      case 'painter': {
        applyFormatPainter(primaryRange(S().sel));
        break;
      }
      case 'move': {
        fillPreview.current = null;
        const t = d.target;
        if (t.r1 === d.src.r1 && t.c1 === d.src.c1) break;
        moveOrCopyBlock(d.src, t, d.copy);
        break;
      }
      case 'split': {
        const sp = sheet.split;
        if (sp) {
          sheet.split = null;
          transact('Split', (tx) => tx.setMeta(sheet, 'split', sp.x <= 0 && sp.y <= 0 ? null : sp));
        }
        break;
      }
      default:
        break;
    }
    (e.target as HTMLElement).releasePointerCapture?.(e.pointerId);
    focusInputSoon();
  };

  const focusInputSoon = () => setTimeout(() => focusInput(), 0);

  const onDoubleClick = (e: React.MouseEvent) => {
    const vp = vpNow();
    if (!vp) return;
    const { x, y } = localXY(e);
    const hit = hitTest(vp, x, y);
    const st = S();
    if (hit.area === 'colHeader' && hit.resize !== undefined) {
      const cols = st.sel.ranges.some((rg) => rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1 && hit.resize! >= rg.c1 && hit.resize! <= rg.c2)
        ? st.sel.ranges.flatMap((rg) => Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => rg.c1 + i))
        : [hit.resize];
      autoFitColumns(cols);
      return;
    }
    if (hit.area === 'rowHeader' && hit.resize !== undefined) {
      autoFitRows([hit.resize]);
      return;
    }
    if (hit.area !== 'cell') return;
    if (fillHandleHit(x, y)) {
      // fill down to the extent of the adjacent column's data
      const sheet = st.wb.activeSheet;
      const src = primaryRange(st.sel);
      const probe = src.c1 > 0 ? src.c1 - 1 : src.c2 + 1;
      let r = src.r2;
      const has = (rr: number, cc: number) => {
        const c = sheet.getCell(rr, cc);
        return !!c && (c.v !== undefined || c.f !== undefined);
      };
      if (!has(r + 1, probe) && has(r + 1, src.c2 + 1)) {
        while (has(r + 1, src.c2 + 1)) r++;
      } else while (has(r + 1, probe)) r++;
      if (r > src.r2) autoFill(src, { ...src, r2: r });
      return;
    }
    const sheet = st.wb.activeSheet;
    const cell = sheet.getCell(hit.r, hit.c);
    if (sheet.protection && st.wb.styles.get(sheet.styleIdAt(hit.r, hit.c)).locked !== false) {
      beginEdit('edit');
      return;
    }
    void cell;
    beginEdit('edit');
  };

  const onContextMenu = (e: React.MouseEvent) => {
    e.preventDefault();
    const vp = vpNow();
    if (!vp) return;
    const { x, y } = localXY(e);
    const hit = hitTest(vp, x, y);
    const st = S();
    if (st.edit) {
      if (!commitEdit('none')) return;
    }
    const inSel = (r: number, c: number) => st.sel.ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2);
    if (hit.area === 'colHeader') {
      if (!st.sel.ranges.some((rg) => rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1 && hit.c >= rg.c1 && hit.c <= rg.c2))
        setState({ sel: { ranges: [{ r1: 0, r2: MAX_ROWS - 1, c1: hit.c, c2: hit.c }], active: { r: 0, c: hit.c }, anchor: { r: 0, c: hit.c } } });
      setState({ menu: { x: e.clientX, y: e.clientY, kind: 'col' } });
    } else if (hit.area === 'rowHeader') {
      if (!st.sel.ranges.some((rg) => rg.c1 === 0 && rg.c2 >= MAX_COLS - 1 && hit.r >= rg.r1 && hit.r <= rg.r2))
        setState({ sel: { ranges: [{ r1: hit.r, r2: hit.r, c1: 0, c2: MAX_COLS - 1 }], active: { r: hit.r, c: 0 }, anchor: { r: hit.r, c: 0 } } });
      setState({ menu: { x: e.clientX, y: e.clientY, kind: 'row' } });
    } else if (hit.area === 'cell') {
      if (!inSel(hit.r, hit.c)) {
        const m = st.wb.activeSheet.mergeAt(hit.r, hit.c);
        setState({ sel: m ? { ranges: [m], active: { r: m.r1, c: m.c1 }, anchor: { r: m.r1, c: m.c1 } } : singleSel(hit.r, hit.c) });
      }
      setState({ menu: { x: e.clientX, y: e.clientY, kind: 'cell' } });
    } else if (hit.area === 'corner') setState({ menu: { x: e.clientX, y: e.clientY, kind: 'corner' } });
  };

  // ---------- hidden input (keyboard, clipboard, IME) ----------
  useEffect(() => {
    const el = inputRef.current;
    if (!el) return;
    const onCopy = (e: ClipboardEvent) => onCopyEvent(e, false);
    const onCut = (e: ClipboardEvent) => onCopyEvent(e, true);
    const onPaste = (e: ClipboardEvent) => onPasteEvent(e);
    el.addEventListener('copy', onCopy);
    el.addEventListener('cut', onCut);
    el.addEventListener('paste', onPaste);
    return () => {
      el.removeEventListener('copy', onCopy);
      el.removeEventListener('cut', onCut);
      el.removeEventListener('paste', onPaste);
    };
  }, []);

  // keep the hidden input focused when not editing
  useEffect(() => {
    if (!edit && !dialog) {
      const active = document.activeElement as HTMLElement | null;
      const typing = active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA' || active.tagName === 'SELECT' || active.isContentEditable) && active !== inputRef.current;
      if (!typing) focusInput();
    }
  }, [edit, dialog]);

  const onKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    const ne = e.nativeEvent;
    const ed = S().edit;
    if (ed) {
      // The in-cell editor hasn't taken focus yet (it mounts a frame later): route keys into the edit session.
      if (ne.isComposing) return;
      if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        e.preventDefault();
        const text = ed.text.slice(0, ed.caret) + e.key + ed.text.slice(ed.caret);
        setState({ edit: { ...ed, text, caret: ed.caret + 1, point: undefined } });
        return;
      }
      if (e.key === 'Backspace') {
        e.preventDefault();
        if (ed.caret > 0) setState({ edit: { ...ed, text: ed.text.slice(0, ed.caret - 1) + ed.text.slice(ed.caret), caret: ed.caret - 1 } });
        return;
      }
      const fake = { selectionStart: ed.caret, selectionEnd: ed.caret } as HTMLTextAreaElement;
      if (handleEditKey(ne, fake, null, () => undefined, () => undefined, () => undefined) || e.key === 'Tab') e.preventDefault();
      return;
    }
    if (ne.isComposing) return;
    if (e.key === 'Tab' && e.ctrlKey) return;
    if (handleGridKey(ne)) {
      e.preventDefault();
      return;
    }
    // printable char starts editing
    if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      e.preventDefault();
      startTyping(e.key);
    }
  };

  const onCompositionStart = () => {
    if (!S().edit) beginEdit('enter', '');
  };

  return (
    <div
      ref={rootRef}
      data-grid-root
      className="relative flex-1 overflow-hidden select-none"
      style={{ cursor, background: 'var(--grid-bg)' }}
      onPointerDown={onPointerDown}
      onPointerMove={onPointerMove}
      onPointerUp={onPointerUp}
      onPointerLeave={() => {
        if (noteTimer.current) clearTimeout(noteTimer.current);
      }}
      onMouseDown={(e) => {
        // keep keyboard focus on the hidden grid input (the default action would blur it)
        if (!(e.target as HTMLElement).closest('textarea,input,select,button,[data-allow-focus]')) e.preventDefault();
      }}
      onDoubleClick={onDoubleClick}
      onContextMenu={onContextMenu}
    >
      <canvas ref={canvasRef} style={{ width: size.w, height: size.h, display: 'block' }} data-testid="grid-canvas" />
      <textarea
        ref={inputRef}
        aria-label="Grid input"
        data-testid="grid-input"
        className="absolute opacity-0 pointer-events-none"
        style={{ left: 0, top: 0, width: 1, height: 1, resize: 'none' }}
        onKeyDown={onKeyDown}
        onCompositionStart={onCompositionStart}
        tabIndex={0}
        autoComplete="off"
        spellCheck={false}
      />
      {edit && editorRect && edit.hostSheetId === S().wb.activeSheetId && <CellEditor rect={editorRect} viewportWidth={size.w} viewportHeight={size.h} />}
      <ChartLayer />
      {fillOpts && <AutoFillOptions x={fillOpts.x} y={fillOpts.y} src={fillOpts.src} target={fillOpts.target} onClose={() => setFillOpts(null)} />}
      {notePos && !edit && <NotePopup r={notePos.r} c={notePos.c} x={notePos.x} y={notePos.y} />}
      {resizeTip && (
        <div className="absolute z-30 px-1.5 py-0.5 text-[11px] border shadow-sm pointer-events-none" style={{ left: resizeTip.x + 8, top: resizeTip.y + 4, background: 'var(--tooltip-bg)', borderColor: 'var(--border)', color: 'var(--text)' }}>
          {resizeTip.text}
        </div>
      )}
      {filterMenu && <FilterMenu />}
      {listMenu && <ListDropdown getRect={(r, c) => (currentVp ? rangeRect(currentVp, r, c, r, c) : null)} />}
      <InputMessage />
      <VScrollbar />
    </div>
  );

  function openLink(link: string) {
    if (/^#/.test(link) || /^[^/]*!/.test(link)) {
      import('../../state/actions/find').then((m) => m.goTo(link.replace(/^#/, '')));
      return;
    }
    const url = /^[a-z]+:/i.test(link) ? link : /@/.test(link) ? 'mailto:' + link : 'https://' + link;
    window.open(url, '_blank', 'noopener');
  }
}

function moveOrCopyBlock(src: Range, target: Range, copy: boolean): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  const dr = target.r1 - src.r1;
  const dc = target.c1 - src.c1;
  const occupied = (() => {
    let any = false;
    sheet.forEachInRange(target, (r, c, cell) => {
      if (r >= src.r1 && r <= src.r2 && c >= src.c1 && c <= src.c2) return;
      if (cell.v !== undefined || cell.f !== undefined) any = true;
    });
    return any;
  })();
  const doIt = () => {
    transact(copy ? 'Copy' : 'Move', (tx) => {
      if (copy) {
        const cells: { r: number; c: number; cell: import('../../model/types').Cell }[] = [];
        sheet.forEachInRange(src, (r, c, cell) => cells.push({ r, c, cell }));
        for (let r = target.r1; r <= target.r2; r++) for (let c = target.c1; c <= target.c2; c++) if (sheet.getCell(r, c)) tx.setCell(sheet, r, c, undefined);
        for (const { r, c, cell } of cells) {
          tx.setCell(sheet, r + dr, c + dc, cell.f ? { ...cell, f: shiftFormula(cell.f, dr, dc) } : cell);
        }
      } else {
        // clear destination first, then move with reference retargeting
        for (let r = target.r1; r <= target.r2; r++)
          for (let c = target.c1; c <= target.c2; c++) {
            if (r >= src.r1 && r <= src.r2 && c >= src.c1 && c <= src.c2) continue;
            if (sheet.getCell(r, c)) tx.setCell(sheet, r, c, undefined);
          }
        shiftBlock(tx, sheet, src, dr, dc);
      }
    }, { ranges: [target], active: { r: target.r1, c: target.c1 }, anchor: { r: target.r1, c: target.c1 } });
  };
  if (occupied) {
    import('../../state/store').then((m) =>
      m.openDialog('confirm', { message: "There's already data here. Do you want to replace it?", onOk: doIt }),
    );
  } else doIt();
}

function pageBreaks(sheet: import('../../model/sheet').Sheet): { rows: number[]; cols: number[] } {
  // Automatic page breaks: approximate printable area of the chosen paper
  const ps = sheet.pageSetup;
  const inch = 96;
  const paper: Record<string, [number, number]> = { letter: [8.5, 11], legal: [8.5, 14], a4: [8.27, 11.69], a3: [11.69, 16.54], tabloid: [11, 17] };
  let [pw, ph] = paper[ps.paperSize] ?? paper.letter;
  if (ps.orientation === 'landscape') [pw, ph] = [ph, pw];
  const scale = (ps.fitToPage ? 100 : ps.scale) / 100;
  const availW = ((pw - ps.margins.left - ps.margins.right) * inch) / scale;
  const availH = ((ph - ps.margins.top - ps.margins.bottom) * inch) / scale;
  const used = sheet.usedRange();
  if (!used) return { rows: sheet.rowBreaks.map((b) => b + 1), cols: sheet.colBreaks.map((b) => b + 1) };
  const rows: number[] = [];
  let acc = 0;
  const manualR = new Set(sheet.rowBreaks.map((b) => b + 1));
  for (let r = 0; r <= used.r2; r++) {
    const h = sheet.rowHeight(r);
    if (manualR.has(r) || acc + h > availH) {
      rows.push(r);
      acc = 0;
    }
    acc += h;
  }
  rows.push(used.r2 + 1);
  const cols: number[] = [];
  acc = 0;
  const manualC = new Set(sheet.colBreaks.map((b) => b + 1));
  for (let c = 0; c <= used.c2; c++) {
    const w = sheet.colWidth(c);
    if (manualC.has(c) || acc + w > availW) {
      cols.push(c);
      acc = 0;
    }
    acc += w;
  }
  cols.push(used.c2 + 1);
  return { rows, cols };
}

// ---------- vertical scrollbar ----------
import { onScrollChange, gridApi } from './scrollBus';

function VScrollbar() {
  const [, force] = useState(0);
  const rev = useStore((s) => s.rev);
  void rev;
  useEffect(() => onScrollChange(() => force((n) => n + 1)), []);
  const api = gridApi();
  const ext = api?.extent('y');
  const trackRef = useRef<HTMLDivElement>(null);
  const drag = useRef<{ y: number; pos: number } | null>(null);
  if (!ext) return null;
  const vp = currentVp;
  const top = vp ? vp.hh : 20;
  const trackH = Math.max(40, (vp?.height ?? 300) - top - 34);
  const thumbH = Math.max(18, (trackH * ext.view) / ext.total);
  const thumbY = ((trackH - thumbH) * ext.pos) / Math.max(1, ext.total - ext.view);
  const onDown = (e: React.PointerEvent) => {
    e.stopPropagation();
    drag.current = { y: e.clientY, pos: ext.pos };
    (e.target as HTMLElement).setPointerCapture(e.pointerId);
  };
  const onMove = (e: React.PointerEvent) => {
    if (!drag.current) return;
    e.stopPropagation();
    const dy = e.clientY - drag.current.y;
    const pos = drag.current.pos + (dy * Math.max(1, ext.total - ext.view)) / Math.max(1, trackH - thumbH);
    api!.setScroll('y', pos);
  };
  const onUp = (e: React.PointerEvent) => {
    e.stopPropagation();
    drag.current = null;
  };
  const step = (dir: number) => (e: React.PointerEvent) => {
    e.stopPropagation();
    api!.setScroll('y', ext.pos + dir * 20 * ((vp?.z ?? 1) * 3));
  };
  const onTrack = (e: React.PointerEvent) => {
    e.stopPropagation();
    const rect = trackRef.current!.getBoundingClientRect();
    const y = e.clientY - rect.top;
    api!.setScroll('y', ext.pos + (y < thumbY ? -ext.view : ext.view));
  };
  return (
    <div className="xl-vscroll absolute right-0 flex flex-col items-stretch" style={{ top, bottom: 0, width: 17 }} onPointerDown={(e) => e.stopPropagation()} onDoubleClick={(e) => e.stopPropagation()}>
      <button className="xl-scroll-btn" style={{ height: 17 }} onPointerDown={step(-1)} aria-label="Scroll up">
        <svg width="7" height="4" viewBox="0 0 7 4"><path d="M0 4 L3.5 0 L7 4Z" fill="currentColor" /></svg>
      </button>
      <div ref={trackRef} className="relative flex-1" onPointerDown={onTrack}>
        <div className="xl-scroll-thumb absolute left-[3px] right-[3px]" style={{ top: thumbY, height: thumbH }} onPointerDown={onDown} onPointerMove={onMove} onPointerUp={onUp} />
      </div>
      <button className="xl-scroll-btn" style={{ height: 17 }} onPointerDown={step(1)} aria-label="Scroll down">
        <svg width="7" height="4" viewBox="0 0 7 4"><path d="M0 0 L3.5 4 L7 0Z" fill="currentColor" /></svg>
      </button>
    </div>
  );
}

function InputMessage() {
  const sel = useStore((s) => s.sel);
  const edit = useStore((s) => s.edit);
  useStore((s) => s.rev);
  const sheet = S().wb.activeSheet;
  const dv = validationAt(sheet, sel.active.r, sel.active.c);
  const vp = currentVp;
  if (!dv || !dv.showInput || !(dv.prompt || dv.promptTitle) || !vp || edit) return null;
  const R = rangeRect(vp, sel.active.r, sel.active.c, sel.active.r, sel.active.c);
  if (!R) return null;
  return (
    <div className="absolute z-20 max-w-[220px] px-2 py-1.5 text-[12px] border shadow" style={{ left: R.x + Math.min(R.w, 40), top: R.y + R.h + 6, background: 'var(--note-bg)', borderColor: 'var(--border-strong)', color: '#000' }}>
      {dv.promptTitle && <div className="font-bold mb-0.5">{dv.promptTitle}</div>}
      <div className="whitespace-pre-wrap">{dv.prompt}</div>
    </div>
  );
}

export { colX, rowY };

// Grid viewport geometry: panes (normal / frozen / split), coordinates and hit testing.
import { AxisLayout, sheetLayout } from '../../model/layout';
import type { Sheet } from '../../model/sheet';

export const SPLIT_BAR = 5;

export interface Pane {
  /** Screen position (CSS px) where the pane starts. */
  start: number;
  size: number;
  /** Content offset (zoomed px from index 0) shown at `start`. */
  content: number;
  first: number;
  last: number;
  /** Index in the scroll array this pane is driven by (-1 = fixed/frozen). */
  scrollIdx: number;
}

export interface ScrollPos {
  x: [number, number];
  y: [number, number];
}

export interface Viewport {
  z: number;
  hw: number;
  hh: number;
  width: number;
  height: number;
  colPanes: Pane[];
  rowPanes: Pane[];
  rows: AxisLayout;
  cols: AxisLayout;
  frozen: boolean;
  split: boolean;
}

const scrolls = new Map<string, ScrollPos>();

export function getScroll(sheetId: string): ScrollPos {
  let s = scrolls.get(sheetId);
  if (!s) scrolls.set(sheetId, (s = { x: [0, 0], y: [0, 0] }));
  return s;
}

export function rowHeaderWidth(maxRow: number, z: number): number {
  const digits = String(maxRow + 1).length;
  return Math.round((Math.max(2, digits) * 7 + 12) * z);
}

export function buildViewport(sheet: Sheet, width: number, height: number, scroll: ScrollPos): Viewport {
  const z = sheet.zoom / 100;
  const { rows, cols } = sheetLayout(sheet);
  const hh = sheet.showHeaders ? Math.round(20 * z) : 0;
  // estimate last visible row for header width
  const lastRow = rows.indexAt((scroll.y[1] + scroll.y[0] + height) / z + rows.offset(sheet.freeze.rows));
  const hw = sheet.showHeaders ? rowHeaderWidth(lastRow, z) : 0;
  const colPanes: Pane[] = [];
  const rowPanes: Pane[] = [];
  const fr = sheet.freeze.rows;
  const fc = sheet.freeze.cols;
  const frozen = fr > 0 || fc > 0;
  const split = !frozen && !!sheet.split;
  if (frozen) {
    const FW = cols.offset(fc) * z;
    const FH = rows.offset(fr) * z;
    if (fc > 0) colPanes.push({ start: hw, size: FW, content: 0, first: 0, last: fc - 1, scrollIdx: -1 });
    colPanes.push({ start: hw + FW, size: Math.max(0, width - hw - FW), content: FW + scroll.x[1], first: fc, last: cols.count - 1, scrollIdx: 1 });
    if (fr > 0) rowPanes.push({ start: hh, size: FH, content: 0, first: 0, last: fr - 1, scrollIdx: -1 });
    rowPanes.push({ start: hh + FH, size: Math.max(0, height - hh - FH), content: FH + scroll.y[1], first: fr, last: rows.count - 1, scrollIdx: 1 });
  } else if (split) {
    const sx = Math.min(width - hw - 40, Math.max(0, (cols.offset(sheet.split!.x) * z)));
    const sy = Math.min(height - hh - 40, Math.max(0, rows.offset(sheet.split!.y) * z));
    if (sx > 0) {
      colPanes.push({ start: hw, size: sx, content: scroll.x[0], first: 0, last: cols.count - 1, scrollIdx: 0 });
      colPanes.push({ start: hw + sx + SPLIT_BAR, size: Math.max(0, width - hw - sx - SPLIT_BAR), content: scroll.x[1], first: 0, last: cols.count - 1, scrollIdx: 1 });
    } else colPanes.push({ start: hw, size: width - hw, content: scroll.x[1], first: 0, last: cols.count - 1, scrollIdx: 1 });
    if (sy > 0) {
      rowPanes.push({ start: hh, size: sy, content: scroll.y[0], first: 0, last: rows.count - 1, scrollIdx: 0 });
      rowPanes.push({ start: hh + sy + SPLIT_BAR, size: Math.max(0, height - hh - sy - SPLIT_BAR), content: scroll.y[1], first: 0, last: rows.count - 1, scrollIdx: 1 });
    } else rowPanes.push({ start: hh, size: height - hh, content: scroll.y[1], first: 0, last: rows.count - 1, scrollIdx: 1 });
  } else {
    colPanes.push({ start: hw, size: Math.max(0, width - hw), content: scroll.x[1], first: 0, last: cols.count - 1, scrollIdx: 1 });
    rowPanes.push({ start: hh, size: Math.max(0, height - hh), content: scroll.y[1], first: 0, last: rows.count - 1, scrollIdx: 1 });
  }
  return { z, hw, hh, width, height, colPanes, rowPanes, rows, cols, frozen, split };
}

export function colX(vp: Viewport, p: Pane, c: number): number {
  return p.start + vp.cols.offset(c) * vp.z - p.content;
}

export function rowY(vp: Viewport, p: Pane, r: number): number {
  return p.start + vp.rows.offset(r) * vp.z - p.content;
}

export function paneRange(vp: Viewport, p: Pane, axis: 'row' | 'col'): [number, number] {
  const L = axis === 'row' ? vp.rows : vp.cols;
  const a = Math.max(p.first, L.indexAt(p.content / vp.z));
  const b = Math.min(p.last, L.indexAt((p.content + p.size) / vp.z));
  return [a, b];
}

function paneAt(panes: Pane[], pos: number): Pane | null {
  for (const p of panes) if (pos >= p.start && pos < p.start + p.size) return p;
  return null;
}

export interface Hit {
  area: 'corner' | 'colHeader' | 'rowHeader' | 'cell' | 'splitV' | 'splitH' | 'none';
  r: number;
  c: number;
  colPane: Pane | null;
  rowPane: Pane | null;
  /** Near a header border (for resizing): index whose trailing edge is under the pointer. */
  resize?: number;
}

export function indexFromPos(vp: Viewport, panes: Pane[], pos: number, axis: 'row' | 'col'): { idx: number; pane: Pane } {
  const L = axis === 'row' ? vp.rows : vp.cols;
  let p = paneAt(panes, pos);
  if (!p) p = pos < panes[0].start ? panes[0] : panes[panes.length - 1];
  const content = (pos - p.start + p.content) / vp.z;
  let idx = L.indexAt(Math.max(0, content));
  idx = Math.max(p.first, Math.min(p.last, idx));
  return { idx, pane: p };
}

export function hitTest(vp: Viewport, x: number, y: number): Hit {
  const inColHdr = y < vp.hh;
  const inRowHdr = x < vp.hw;
  if (inColHdr && inRowHdr) return { area: 'corner', r: 0, c: 0, colPane: null, rowPane: null };
  if (vp.split) {
    const sp = vp.colPanes.length > 1 ? vp.colPanes[0].start + vp.colPanes[0].size : -99;
    if (x >= sp && x < sp + SPLIT_BAR) return { area: 'splitV', r: 0, c: 0, colPane: null, rowPane: null };
    const sy = vp.rowPanes.length > 1 ? vp.rowPanes[0].start + vp.rowPanes[0].size : -99;
    if (y >= sy && y < sy + SPLIT_BAR) return { area: 'splitH', r: 0, c: 0, colPane: null, rowPane: null };
  }
  const cc = indexFromPos(vp, vp.colPanes, Math.max(vp.hw, x), 'col');
  const rr = indexFromPos(vp, vp.rowPanes, Math.max(vp.hh, y), 'row');
  if (inColHdr) {
    // resize detection near the right edge of a column (or left edge of the next)
    let resize: number | undefined;
    const x1 = colX(vp, cc.pane, cc.idx);
    const x2 = x1 + vp.cols.size(cc.idx) * vp.z;
    if (Math.abs(x - x2) <= 3) resize = cc.idx;
    else if (Math.abs(x - x1) <= 3 && cc.idx > 0) {
      // hidden columns: resize the previous visible one
      let p = cc.idx - 1;
      while (p > 0 && vp.cols.size(p) === 0) p--;
      resize = p;
    }
    return { area: 'colHeader', r: 0, c: cc.idx, colPane: cc.pane, rowPane: null, resize };
  }
  if (inRowHdr) {
    let resize: number | undefined;
    const y1 = rowY(vp, rr.pane, rr.idx);
    const y2 = y1 + vp.rows.size(rr.idx) * vp.z;
    if (Math.abs(y - y2) <= 2) resize = rr.idx;
    else if (Math.abs(y - y1) <= 2 && rr.idx > 0) {
      let p = rr.idx - 1;
      while (p > 0 && vp.rows.size(p) === 0) p--;
      resize = p;
    }
    return { area: 'rowHeader', r: rr.idx, c: 0, colPane: null, rowPane: rr.pane, resize };
  }
  return { area: 'cell', r: rr.idx, c: cc.idx, colPane: cc.pane, rowPane: rr.pane };
}

/** Screen rectangle of a cell range within the pane that shows its top-left (clipped to nothing if off-screen). */
export function rangeRect(vp: Viewport, r1: number, c1: number, r2: number, c2: number): { x: number; y: number; w: number; h: number } | null {
  const cp = vp.colPanes.find((p) => c1 >= p.first && c1 <= p.last) ?? vp.colPanes[vp.colPanes.length - 1];
  const rp = vp.rowPanes.find((p) => r1 >= p.first && r1 <= p.last) ?? vp.rowPanes[vp.rowPanes.length - 1];
  const x = colX(vp, cp, c1);
  const y = rowY(vp, rp, r1);
  const w = (vp.cols.offset(c2 + 1) - vp.cols.offset(c1)) * vp.z;
  const h = (vp.rows.offset(r2 + 1) - vp.rows.offset(r1)) * vp.z;
  return { x, y, w, h };
}

/** Scroll the main panes so (r, c) is visible. Returns true when the scroll changed. */
export function ensureVisible(vp: Viewport, scroll: ScrollPos, r: number, c: number): boolean {
  let changed = false;
  const cp = vp.colPanes[vp.colPanes.length - 1];
  const rp = vp.rowPanes[vp.rowPanes.length - 1];
  const z = vp.z;
  if (c >= cp.first) {
    const x1 = vp.cols.offset(c) * z;
    const x2 = vp.cols.offset(c + 1) * z;
    const base = vp.frozen ? vp.cols.offset(cp.first) * z : 0;
    const cur = cp.content;
    if (x1 < cur) {
      scroll.x[1] = x1 - base;
      changed = true;
    } else if (x2 > cur + cp.size) {
      scroll.x[1] = Math.min(x1, x2 - cp.size) - base;
      changed = true;
    }
  }
  if (r >= rp.first) {
    const y1 = vp.rows.offset(r) * z;
    const y2 = vp.rows.offset(r + 1) * z;
    const base = vp.frozen ? vp.rows.offset(rp.first) * z : 0;
    const cur = rp.content;
    if (y1 < cur) {
      scroll.y[1] = y1 - base;
      changed = true;
    } else if (y2 > cur + rp.size) {
      scroll.y[1] = Math.min(y1, y2 - rp.size) - base;
      changed = true;
    }
  }
  scroll.x[1] = Math.max(0, scroll.x[1]);
  scroll.y[1] = Math.max(0, scroll.y[1]);
  return changed;
}

import { CellAddr, MAX_COLS, MAX_ROWS, normRange, Range } from './address';
import { sheetLayout } from './layout';
import type { Sheet } from './sheet';

export interface Selection {
  ranges: Range[];
  active: CellAddr;
  anchor: CellAddr;
}

export function singleSel(r: number, c: number): Selection {
  return { ranges: [{ r1: r, c1: c, r2: r, c2: c }], active: { r, c }, anchor: { r, c } };
}

export function rangeSel(rg: Range, active?: CellAddr): Selection {
  const a = active ?? { r: rg.r1, c: rg.c1 };
  return { ranges: [rg], active: a, anchor: a };
}

export function primaryRange(sel: Selection): Range {
  return sel.ranges[sel.ranges.length - 1];
}

/** Grow a range so it fully contains every merged area it touches. */
export function expandForMerges(sheet: Sheet, rg: Range): Range {
  if (!sheet.merges.length) return rg;
  let cur = { ...rg };
  let changed = true;
  let guard = 0;
  while (changed && guard++ < 50) {
    changed = false;
    for (const m of sheet.merges) {
      if (m.r1 <= cur.r2 && cur.r1 <= m.r2 && m.c1 <= cur.c2 && cur.c1 <= m.c2) {
        const n = { r1: Math.min(cur.r1, m.r1), c1: Math.min(cur.c1, m.c1), r2: Math.max(cur.r2, m.r2), c2: Math.max(cur.c2, m.c2) };
        if (n.r1 !== cur.r1 || n.c1 !== cur.c1 || n.r2 !== cur.r2 || n.c2 !== cur.c2) {
          cur = n;
          changed = true;
        }
      }
    }
  }
  return cur;
}

export function hasData(sheet: Sheet, r: number, c: number): boolean {
  const cell = sheet.getCell(r, c);
  return !!cell && (cell.f !== undefined || (cell.v !== undefined && cell.v !== null && cell.v !== ''));
}

/** Move one step, skipping hidden rows/cols and jumping over merged areas. */
export function step(sheet: Sheet, from: CellAddr, dr: number, dc: number): CellAddr {
  const { rows, cols } = sheetLayout(sheet);
  let r = from.r;
  let c = from.c;
  const m = sheet.mergeAt(r, c);
  if (m) {
    if (dr > 0) r = m.r2;
    if (dr < 0) r = m.r1;
    if (dc > 0) c = m.c2;
    if (dc < 0) c = m.c1;
  }
  if (dr) r = rows.nextVisible(r, dr > 0 ? 1 : -1);
  if (dc) c = cols.nextVisible(c, dc > 0 ? 1 : -1);
  const m2 = sheet.mergeAt(r, c);
  if (m2) return { r: m2.r1, c: m2.c1 };
  return { r, c };
}

/** Excel Ctrl+Arrow: jump to the edge of the current data region. */
export function dataEdge(sheet: Sheet, from: CellAddr, dr: number, dc: number): CellAddr {
  const { rows, cols } = sheetLayout(sheet);
  if (dr !== 0) {
    // collect data rows in column
    const c = from.c;
    const list: number[] = [];
    for (const [r, row] of sheet.rows) {
      const cell = row.get(c);
      if (cell && (cell.f !== undefined || (cell.v !== undefined && cell.v !== null && cell.v !== ''))) {
        if (rows.size(r) > 0) list.push(r);
      }
    }
    list.sort((a, b) => a - b);
    const set = new Set(list);
    const r0 = from.r;
    const next = rows.nextVisible(r0, dr > 0 ? 1 : -1);
    if (set.has(r0) && set.has(next) && next !== r0) {
      let r = next;
      while (true) {
        const n = rows.nextVisible(r, dr > 0 ? 1 : -1);
        if (n === r || !set.has(n)) break;
        r = n;
      }
      return { r, c };
    }
    if (dr > 0) {
      const found = list.find((r) => r > r0);
      return { r: found ?? MAX_ROWS - 1, c };
    }
    let found = -1;
    for (const r of list) if (r < r0) found = r;
    return { r: found < 0 ? 0 : found, c };
  }
  const r = from.r;
  const row = sheet.rows.get(r);
  const list: number[] = [];
  if (row)
    for (const [c, cell] of row)
      if ((cell.f !== undefined || (cell.v !== undefined && cell.v !== null && cell.v !== '')) && cols.size(c) > 0) list.push(c);
  list.sort((a, b) => a - b);
  const set = new Set(list);
  const c0 = from.c;
  const next = cols.nextVisible(c0, dc > 0 ? 1 : -1);
  if (set.has(c0) && set.has(next) && next !== c0) {
    let c = next;
    while (true) {
      const n = cols.nextVisible(c, dc > 0 ? 1 : -1);
      if (n === c || !set.has(n)) break;
      c = n;
    }
    return { r, c };
  }
  if (dc > 0) {
    const found = list.find((c) => c > c0);
    return { r, c: found ?? MAX_COLS - 1 };
  }
  let found = -1;
  for (const c of list) if (c < c0) found = c;
  return { r, c: found < 0 ? 0 : found };
}

/** Selection with anchor fixed, extended to `to`. */
export function extendTo(sheet: Sheet, sel: Selection, to: CellAddr): Selection {
  const rg = expandForMerges(sheet, normRange({ r1: sel.anchor.r, c1: sel.anchor.c, r2: to.r, c2: to.c }));
  const ranges = sel.ranges.slice(0, -1).concat([rg]);
  return { ranges, anchor: sel.anchor, active: sel.active };
}

/** The moving corner of the primary range (opposite the anchor). */
export function movingCorner(sel: Selection): CellAddr {
  const rg = primaryRange(sel);
  return {
    r: sel.anchor.r === rg.r1 ? rg.r2 : rg.r1,
    c: sel.anchor.c === rg.c1 ? rg.c2 : rg.c1,
  };
}

/** Enter/Tab inside a multi-cell selection cycles the active cell without collapsing it. */
export function cycleActive(sheet: Sheet, sel: Selection, dr: number, dc: number): Selection {
  const multi = sel.ranges.length > 1 || (() => {
    const rg = primaryRange(sel);
    const m = sheet.mergeAt(rg.r1, rg.c1);
    const single = rg.r1 === rg.r2 && rg.c1 === rg.c2;
    const isMerge = !!m && m.r1 === rg.r1 && m.c1 === rg.c1 && m.r2 === rg.r2 && m.c2 === rg.c2;
    return !single && !isMerge;
  })();
  if (!multi) {
    const a = step(sheet, sel.active, dr, dc);
    return singleSel(a.r, a.c);
  }
  let idx = sel.ranges.findIndex((rg) => sel.active.r >= rg.r1 && sel.active.r <= rg.r2 && sel.active.c >= rg.c1 && sel.active.c <= rg.c2);
  if (idx < 0) idx = sel.ranges.length - 1;
  let rg = sel.ranges[idx];
  let { r, c } = sel.active;
  const { rows, cols } = sheetLayout(sheet);
  const fwd = dr > 0 || dc > 0;
  const byRow = dc !== 0; // Tab moves across columns first
  for (let guard = 0; guard < 100000; guard++) {
    if (byRow) {
      c += fwd ? 1 : -1;
      if (c > rg.c2) {
        c = rg.c1;
        r++;
      } else if (c < rg.c1) {
        c = rg.c2;
        r--;
      }
    } else {
      r += fwd ? 1 : -1;
      if (r > rg.r2) {
        r = rg.r1;
        c++;
      } else if (r < rg.r1) {
        r = rg.r2;
        c--;
      }
    }
    if (r > rg.r2 || c > rg.c2 || r < rg.r1 || c < rg.c1) {
      idx = (idx + (fwd ? 1 : sel.ranges.length - 1)) % sel.ranges.length;
      rg = sel.ranges[idx];
      r = fwd ? rg.r1 : rg.r2;
      c = fwd ? rg.c1 : rg.c2;
    }
    if (rows.size(r) === 0 || cols.size(c) === 0) continue;
    const m = sheet.mergeAt(r, c);
    if (m && (m.r1 !== r || m.c1 !== c)) continue;
    break;
  }
  return { ...sel, active: { r, c } };
}

export function selectionContains(sel: Selection, r: number, c: number): boolean {
  return sel.ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2);
}

export function clampAddr(a: CellAddr): CellAddr {
  return { r: Math.max(0, Math.min(MAX_ROWS - 1, a.r)), c: Math.max(0, Math.min(MAX_COLS - 1, a.c)) };
}

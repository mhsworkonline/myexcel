// A1-style address utilities. All indexes are 0-based internally.

export const MAX_ROWS = 1048576;
export const MAX_COLS = 16384;

export interface CellAddr {
  r: number;
  c: number;
}

/** Inclusive rectangular range. */
export interface Range {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export function colToName(c: number): string {
  let s = '';
  let n = c + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function nameToCol(name: string): number {
  let n = 0;
  const up = name.toUpperCase();
  for (let i = 0; i < up.length; i++) n = n * 26 + (up.charCodeAt(i) - 64);
  return n - 1;
}

export function addrToA1(r: number, c: number, absR = false, absC = false): string {
  return `${absC ? '$' : ''}${colToName(c)}${absR ? '$' : ''}${r + 1}`;
}

export function rangeToA1(rg: Range, abs = false): string {
  if (isFullCols(rg)) return `${abs ? '$' : ''}${colToName(rg.c1)}:${abs ? '$' : ''}${colToName(rg.c2)}`;
  if (isFullRows(rg)) return `${abs ? '$' : ''}${rg.r1 + 1}:${abs ? '$' : ''}${rg.r2 + 1}`;
  const a = addrToA1(rg.r1, rg.c1, abs, abs);
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) return a;
  return `${a}:${addrToA1(rg.r2, rg.c2, abs, abs)}`;
}

const CELL_RE = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/;
const COL_RE = /^\$?([A-Za-z]{1,3})$/;
const ROW_RE = /^\$?(\d{1,7})$/;

export function parseA1(s: string): CellAddr | null {
  const m = CELL_RE.exec(s.trim());
  if (!m) return null;
  const c = nameToCol(m[1]);
  const r = parseInt(m[2], 10) - 1;
  if (r < 0 || r >= MAX_ROWS || c < 0 || c >= MAX_COLS) return null;
  return { r, c };
}

/** Parses "A1", "A1:B5", "A:C", "3:7". Sheet prefixes are not handled here. */
export function parseRange(s: string): Range | null {
  const t = s.trim();
  const parts = t.split(':');
  if (parts.length === 1) {
    const a = parseA1(parts[0]);
    return a ? { r1: a.r, c1: a.c, r2: a.r, c2: a.c } : null;
  }
  if (parts.length !== 2) return null;
  const a = parseA1(parts[0]);
  const b = parseA1(parts[1]);
  if (a && b) return normRange({ r1: a.r, c1: a.c, r2: b.r, c2: b.c });
  const ca = COL_RE.exec(parts[0]);
  const cb = COL_RE.exec(parts[1]);
  if (ca && cb) return normRange({ r1: 0, r2: MAX_ROWS - 1, c1: nameToCol(ca[1]), c2: nameToCol(cb[1]) });
  const ra = ROW_RE.exec(parts[0]);
  const rb = ROW_RE.exec(parts[1]);
  if (ra && rb) return normRange({ r1: parseInt(ra[1], 10) - 1, r2: parseInt(rb[1], 10) - 1, c1: 0, c2: MAX_COLS - 1 });
  return null;
}

export function normRange(rg: Range): Range {
  return {
    r1: Math.min(rg.r1, rg.r2),
    r2: Math.max(rg.r1, rg.r2),
    c1: Math.min(rg.c1, rg.c2),
    c2: Math.max(rg.c1, rg.c2),
  };
}

export function cellRange(r: number, c: number): Range {
  return { r1: r, c1: c, r2: r, c2: c };
}

export function inRange(rg: Range, r: number, c: number): boolean {
  return r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2;
}

export function rangesIntersect(a: Range, b: Range): boolean {
  return a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
}

export function intersect(a: Range, b: Range): Range | null {
  if (!rangesIntersect(a, b)) return null;
  return { r1: Math.max(a.r1, b.r1), c1: Math.max(a.c1, b.c1), r2: Math.min(a.r2, b.r2), c2: Math.min(a.c2, b.c2) };
}

export function unionRange(a: Range, b: Range): Range {
  return { r1: Math.min(a.r1, b.r1), c1: Math.min(a.c1, b.c1), r2: Math.max(a.r2, b.r2), c2: Math.max(a.c2, b.c2) };
}

export function rangeEquals(a: Range, b: Range): boolean {
  return a.r1 === b.r1 && a.r2 === b.r2 && a.c1 === b.c1 && a.c2 === b.c2;
}

export function isFullCols(rg: Range): boolean {
  return rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1;
}

export function isFullRows(rg: Range): boolean {
  return rg.c1 === 0 && rg.c2 >= MAX_COLS - 1;
}

export function rangeSize(rg: Range): { rows: number; cols: number } {
  return { rows: rg.r2 - rg.r1 + 1, cols: rg.c2 - rg.c1 + 1 };
}

export function cellKey(r: number, c: number): string {
  return `${r},${c}`;
}

/** Quote a sheet name for use in formulas when needed. */
export function quoteSheetName(name: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !parseA1(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}

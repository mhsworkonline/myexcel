// Row/column geometry with sparse size overrides. Offsets are in px at 100% zoom.
import { MAX_COLS, MAX_ROWS } from './address';
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT, Sheet } from './sheet';

export class AxisLayout {
  private keys: number[] = [];
  private cum: number[] = []; // cumulative delta including keys[i]
  private sizeOf = new Map<number, number>();

  constructor(public readonly def: number, sizes: Map<number, number>, hidden: Iterable<number>[], public readonly count: number) {
    for (const [k, v] of sizes) if (v !== def) this.sizeOf.set(k, v);
    for (const h of hidden) for (const k of h) this.sizeOf.set(k, 0);
    this.keys = [...this.sizeOf.keys()].sort((a, b) => a - b);
    let acc = 0;
    this.cum = this.keys.map((k) => (acc += this.sizeOf.get(k)! - def));
  }

  size(i: number): number {
    const s = this.sizeOf.get(i);
    return s === undefined ? this.def : s;
  }

  /** Number of override keys strictly less than i. */
  private countBelow(i: number): number {
    let lo = 0;
    let hi = this.keys.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.keys[mid] < i) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  }

  /** Pixel offset of the start of index i. */
  offset(i: number): number {
    const n = this.countBelow(i);
    return i * this.def + (n ? this.cum[n - 1] : 0);
  }

  total(): number {
    return this.offset(this.count);
  }

  /** Index containing pixel px (clamped). Zero-size indexes are skipped. */
  indexAt(px: number): number {
    if (px <= 0) return this.firstVisible(0);
    let lo = 0;
    let hi = this.count - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.offset(mid) <= px) lo = mid;
      else hi = mid - 1;
    }
    // lo is the last index whose start <= px; if it has zero size move forward
    while (lo < this.count - 1 && this.size(lo) === 0) lo++;
    return lo;
  }

  firstVisible(from: number): number {
    let i = from;
    while (i < this.count - 1 && this.size(i) === 0) i++;
    return i;
  }

  nextVisible(i: number, dir: 1 | -1): number {
    let j = i + dir;
    while (j >= 0 && j < this.count && this.size(j) === 0) j += dir;
    if (j < 0 || j >= this.count) return i;
    return j;
  }
}

const cache = new WeakMap<Sheet, { rowsKey: unknown[]; colsKey: unknown[]; rows: AxisLayout; cols: AxisLayout }>();

export function sheetLayout(sheet: Sheet): { rows: AxisLayout; cols: AxisLayout } {
  const rowsKey = [sheet.rowHeights, sheet.hiddenRows, sheet.filteredRows];
  const colsKey = [sheet.colWidths, sheet.hiddenCols];
  let c = cache.get(sheet);
  if (!c || c.rowsKey.some((k, i) => k !== rowsKey[i]) || c.colsKey.some((k, i) => k !== colsKey[i])) {
    const rows =
      c && !c.rowsKey.some((k, i) => k !== rowsKey[i])
        ? c.rows
        : new AxisLayout(DEFAULT_ROW_HEIGHT, sheet.rowHeights, [sheet.hiddenRows, sheet.filteredRows], MAX_ROWS);
    const cols =
      c && !c.colsKey.some((k, i) => k !== colsKey[i]) ? c.cols : new AxisLayout(DEFAULT_COL_WIDTH, sheet.colWidths, [sheet.hiddenCols], MAX_COLS);
    c = { rowsKey, colsKey, rows, cols };
    cache.set(sheet, c);
  }
  return { rows: c.rows, cols: c.cols };
}

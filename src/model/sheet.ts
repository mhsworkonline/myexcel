import { MAX_COLS, MAX_ROWS, Range } from './address';
import {
  AutoFilter,
  Cell,
  CFRule,
  ChartSpec,
  DataValidation,
  defaultPageSetup,
  PageSetup,
  PivotSpec,
  Scenario,
  SheetProtection,
  SparklineGroup,
  TableDef,
} from './types';

export const DEFAULT_COL_WIDTH = 64; // px at 100% (8.43 chars)
export const DEFAULT_ROW_HEIGHT = 20; // px at 100% (15pt)

export type SheetVisibility = 'visible' | 'hidden' | 'veryHidden';

let idSeq = 0;
export function newId(prefix = 'id'): string {
  idSeq++;
  return `${prefix}${Date.now().toString(36)}${idSeq.toString(36)}${Math.random().toString(36).slice(2, 6)}`;
}

/** Meta keys that can be snapshotted by the command layer. */
export interface SheetMeta {
  colWidths: Map<number, number>;
  rowHeights: Map<number, number>;
  /** Size of every column/row without its own entry (Select All + resize, Standard Width, files). */
  defaultColWidth: number;
  defaultRowHeight: number;
  hiddenRows: Set<number>;
  hiddenCols: Set<number>;
  colStyles: Map<number, number>;
  rowStyles: Map<number, number>;
  merges: Range[];
  freeze: { rows: number; cols: number };
  split: { x: number; y: number } | null;
  tabColor?: string;
  visibility: SheetVisibility;
  showGridlines: boolean;
  showHeaders: boolean;
  zoom: number;
  autoFilter: AutoFilter | null;
  conditionalFormats: CFRule[];
  validations: DataValidation[];
  tables: TableDef[];
  protection: SheetProtection | null;
  charts: ChartSpec[];
  pivots: PivotSpec[];
  sparklines: SparklineGroup[];
  pageSetup: PageSetup;
  rowBreaks: number[];
  colBreaks: number[];
  scenarios: Scenario[];
  /** Rows hidden by AutoFilter (distinct from manually hidden rows). */
  filteredRows: Set<number>;
  rtl: boolean;
}

export type MetaKey = keyof SheetMeta;

export class Sheet implements SheetMeta {
  id: string;
  name: string;
  rows = new Map<number, Map<number, Cell>>();
  colWidths = new Map<number, number>();
  rowHeights = new Map<number, number>();
  defaultColWidth = DEFAULT_COL_WIDTH;
  defaultRowHeight = DEFAULT_ROW_HEIGHT;
  hiddenRows = new Set<number>();
  hiddenCols = new Set<number>();
  colStyles = new Map<number, number>();
  rowStyles = new Map<number, number>();
  merges: Range[] = [];
  freeze = { rows: 0, cols: 0 };
  split: { x: number; y: number } | null = null;
  tabColor?: string;
  visibility: SheetVisibility = 'visible';
  showGridlines = true;
  showHeaders = true;
  zoom = 100;
  autoFilter: AutoFilter | null = null;
  conditionalFormats: CFRule[] = [];
  validations: DataValidation[] = [];
  tables: TableDef[] = [];
  protection: SheetProtection | null = null;
  charts: ChartSpec[] = [];
  pivots: PivotSpec[] = [];
  sparklines: SparklineGroup[] = [];
  pageSetup: PageSetup = defaultPageSetup();
  rowBreaks: number[] = [];
  colBreaks: number[] = [];
  scenarios: Scenario[] = [];
  filteredRows = new Set<number>();
  rtl = false;

  /** Bumped on every mutation; used for render caching. */
  rev = 0;
  private usedCache: Range | null | undefined;

  constructor(name: string, id?: string) {
    this.name = name;
    this.id = id ?? newId('sh');
  }

  getCell(r: number, c: number): Cell | undefined {
    return this.rows.get(r)?.get(c);
  }

  /** Raw write; does not notify. Pass undefined to delete. */
  setCellRaw(r: number, c: number, cell: Cell | undefined): void {
    let row = this.rows.get(r);
    if (!cell || isEmptyCell(cell)) {
      if (row) {
        row.delete(c);
        if (row.size === 0) this.rows.delete(r);
      }
    } else {
      if (!row) {
        row = new Map();
        this.rows.set(r, row);
      }
      row.set(c, cell);
    }
    this.touch();
  }

  touch(): void {
    this.rev++;
    this.usedCache = undefined;
  }

  /** Bounding box of non-empty cells (values, formulas, styles, notes). */
  usedRange(): Range | null {
    if (this.usedCache !== undefined) return this.usedCache;
    let r1 = Infinity;
    let r2 = -1;
    let c1 = Infinity;
    let c2 = -1;
    for (const [r, row] of this.rows) {
      if (row.size === 0) continue;
      if (r < r1) r1 = r;
      if (r > r2) r2 = r;
      for (const c of row.keys()) {
        if (c < c1) c1 = c;
        if (c > c2) c2 = c;
      }
    }
    for (const m of this.merges) {
      r1 = Math.min(r1, m.r1);
      c1 = Math.min(c1, m.c1);
      r2 = Math.max(r2, m.r2);
      c2 = Math.max(c2, m.c2);
    }
    this.usedCache = r2 < 0 ? null : { r1, c1, r2, c2 };
    return this.usedCache;
  }

  /** Data (value/formula) extents, ignoring pure-style cells. */
  dataRange(): Range | null {
    let r1 = Infinity;
    let r2 = -1;
    let c1 = Infinity;
    let c2 = -1;
    for (const [r, row] of this.rows) {
      for (const [c, cell] of row) {
        if (cell.f === undefined && (cell.v === undefined || cell.v === null || cell.v === '')) continue;
        if (r < r1) r1 = r;
        if (r > r2) r2 = r;
        if (c < c1) c1 = c;
        if (c > c2) c2 = c;
      }
    }
    return r2 < 0 ? null : { r1, c1, r2, c2 };
  }

  /** Iterate existing cells inside a range in row-major order. */
  forEachInRange(rg: Range, cb: (r: number, c: number, cell: Cell) => void): void {
    const rowCount = rg.r2 - rg.r1 + 1;
    if (rowCount > this.rows.size * 2) {
      const keys = [...this.rows.keys()].filter((r) => r >= rg.r1 && r <= rg.r2).sort((a, b) => a - b);
      for (const r of keys) this.iterRow(r, rg, cb);
    } else {
      for (let r = rg.r1; r <= rg.r2; r++) this.iterRow(r, rg, cb);
    }
  }

  private iterRow(r: number, rg: Range, cb: (r: number, c: number, cell: Cell) => void): void {
    const row = this.rows.get(r);
    if (!row) return;
    const colCount = rg.c2 - rg.c1 + 1;
    if (colCount > row.size * 2) {
      const cols = [...row.keys()].filter((c) => c >= rg.c1 && c <= rg.c2).sort((a, b) => a - b);
      for (const c of cols) cb(r, c, row.get(c)!);
    } else {
      for (let c = rg.c1; c <= rg.c2; c++) {
        const cell = row.get(c);
        if (cell) cb(r, c, cell);
      }
    }
  }

  sortedRowKeys(): number[] {
    return [...this.rows.keys()].sort((a, b) => a - b);
  }

  isRowHidden(r: number): boolean {
    return this.hiddenRows.has(r) || this.filteredRows.has(r);
  }

  isColHidden(c: number): boolean {
    return this.hiddenCols.has(c);
  }

  colWidth(c: number): number {
    if (this.hiddenCols.has(c)) return 0;
    return this.colWidths.get(c) ?? this.defaultColWidth;
  }

  rowHeight(r: number): number {
    if (this.isRowHidden(r)) return 0;
    return this.rowHeights.get(r) ?? this.defaultRowHeight;
  }

  mergeAt(r: number, c: number): Range | undefined {
    for (const m of this.merges) if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) return m;
    return undefined;
  }

  /** Style id resolution: cell > row > column. */
  styleIdAt(r: number, c: number): number {
    const cell = this.getCell(r, c);
    if (cell?.s !== undefined) return cell.s;
    const rs = this.rowStyles.get(r);
    if (rs !== undefined) return rs;
    return this.colStyles.get(c) ?? 0;
  }

  // ----- structural shifting of cell storage (model only; formula text handled by workbook) -----

  shiftRows(at: number, delta: number): Map<number, Map<number, Cell>> {
    // Returns the removed rows when delta < 0.
    const removed = new Map<number, Map<number, Cell>>();
    const next = new Map<number, Map<number, Cell>>();
    for (const [r, row] of this.rows) {
      if (r < at) next.set(r, row);
      else if (delta < 0 && r < at - delta) removed.set(r, row);
      else if (r + delta < MAX_ROWS) next.set(r + delta, row);
    }
    this.rows = next;
    this.rowHeights = shiftMap(this.rowHeights, at, delta, MAX_ROWS);
    this.rowStyles = shiftMap(this.rowStyles, at, delta, MAX_ROWS);
    this.hiddenRows = shiftSet(this.hiddenRows, at, delta, MAX_ROWS);
    this.filteredRows = shiftSet(this.filteredRows, at, delta, MAX_ROWS);
    this.touch();
    return removed;
  }

  shiftCols(at: number, delta: number): Map<number, Map<number, Cell>> {
    // Returns removed cells keyed by row → (col → cell) when delta < 0.
    const removed = new Map<number, Map<number, Cell>>();
    for (const [r, row] of this.rows) {
      const nrow = new Map<number, Cell>();
      for (const [c, cell] of row) {
        if (c < at) nrow.set(c, cell);
        else if (delta < 0 && c < at - delta) {
          let rm = removed.get(r);
          if (!rm) removed.set(r, (rm = new Map()));
          rm.set(c, cell);
        } else if (c + delta < MAX_COLS) nrow.set(c + delta, cell);
      }
      if (nrow.size) this.rows.set(r, nrow);
      else this.rows.delete(r);
    }
    this.colWidths = shiftMap(this.colWidths, at, delta, MAX_COLS);
    this.colStyles = shiftMap(this.colStyles, at, delta, MAX_COLS);
    this.hiddenCols = shiftSet(this.hiddenCols, at, delta, MAX_COLS);
    this.touch();
    return removed;
  }
}

export function isEmptyCell(c: Cell): boolean {
  return (
    (c.v === undefined || c.v === null || c.v === '') &&
    c.f === undefined &&
    (c.s === undefined || c.s === 0) &&
    !c.note &&
    !c.link
  );
}

function shiftMap<T>(m: Map<number, T>, at: number, delta: number, max: number): Map<number, T> {
  const out = new Map<number, T>();
  for (const [k, v] of m) {
    if (k < at) out.set(k, v);
    else if (delta < 0 && k < at - delta) continue;
    else if (k + delta < max) out.set(k + delta, v);
  }
  return out;
}

function shiftSet(s: Set<number>, at: number, delta: number, max: number): Set<number> {
  const out = new Set<number>();
  for (const k of s) {
    if (k < at) out.add(k);
    else if (delta < 0 && k < at - delta) continue;
    else if (k + delta < max) out.add(k + delta);
  }
  return out;
}

export const META_KEYS: MetaKey[] = [
  'colWidths',
  'rowHeights',
  'defaultColWidth',
  'defaultRowHeight',
  'hiddenRows',
  'hiddenCols',
  'colStyles',
  'rowStyles',
  'merges',
  'freeze',
  'split',
  'tabColor',
  'visibility',
  'showGridlines',
  'showHeaders',
  'zoom',
  'autoFilter',
  'conditionalFormats',
  'validations',
  'tables',
  'protection',
  'charts',
  'pivots',
  'sparklines',
  'pageSetup',
  'rowBreaks',
  'colBreaks',
  'scenarios',
  'filteredRows',
  'rtl',
];

export function cloneValue<T>(v: T): T {
  if (v instanceof Map) return new Map(v) as unknown as T;
  if (v instanceof Set) return new Set(v) as unknown as T;
  if (v === undefined || v === null || typeof v !== 'object') return v;
  return structuredClone(v);
}

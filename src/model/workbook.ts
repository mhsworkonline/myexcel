import type { Range } from './address';
import { renameSheetInFormula } from './formula';
import { cloneValue, META_KEYS, MetaKey, newId, Sheet } from './sheet';
import { adjustFormula, adjustRange, adjustRefString, StructOp } from './structure';
import { StyleTable, CellStyle } from './styles';
import type { Cell, DefinedName } from './types';

export interface CellChange {
  r: number;
  c: number;
}

export interface WorkbookListener {
  cellsChanged(sheet: Sheet, changes: CellChange[]): void;
  structureChanged(sheet: Sheet, op: StructOp): void;
  sheetAdded(sheet: Sheet, index: number): void;
  sheetRemoved(sheet: Sheet): void;
  sheetRenamed(sheet: Sheet, oldName: string): void;
  sheetsReordered(): void;
  namesChanged(): void;
}

export interface WorkbookProps {
  title?: string;
  author?: string;
  created?: number;
  modified?: number;
  /** Normal-style font of this workbook when it differs from Calibri (from app settings). */
  defaultFont?: string;
}

export class Workbook {
  sheets: Sheet[] = [];
  styles = new StyleTable();
  names: DefinedName[] = [];
  props: WorkbookProps = {};
  listeners: WorkbookListener[] = [];
  activeSheetId = '';
  /** Increments on every change; used for dirty tracking. */
  rev = 0;

  static createDefault(opts: { defaultFont?: string; zoom?: number } = {}): Workbook {
    const wb = new Workbook();
    const s = new Sheet('Sheet1');
    if (opts.zoom) s.zoom = opts.zoom;
    if (opts.defaultFont && opts.defaultFont !== 'Calibri') wb.props.defaultFont = opts.defaultFont;
    wb.sheets.push(s);
    wb.activeSheetId = s.id;
    return wb;
  }

  get activeSheet(): Sheet {
    return this.sheets.find((s) => s.id === this.activeSheetId) ?? this.sheets[0];
  }

  sheetById(id: string): Sheet | undefined {
    return this.sheets.find((s) => s.id === id);
  }

  sheetByName(name: string): Sheet | undefined {
    const n = name.toLowerCase();
    return this.sheets.find((s) => s.name.toLowerCase() === n);
  }

  style(id: number | undefined): CellStyle {
    return this.styles.get(id);
  }

  bump(): void {
    this.rev++;
  }

  // ----- cell mutations (notify engine) -----

  setCells(sheet: Sheet, changes: { r: number; c: number; cell: Cell | undefined }[]): void {
    if (!changes.length) return;
    for (const ch of changes) sheet.setCellRaw(ch.r, ch.c, ch.cell);
    this.bump();
    const list = changes.map((c) => ({ r: c.r, c: c.c }));
    for (const l of this.listeners) l.cellsChanged(sheet, list);
  }

  setMeta<K extends MetaKey>(sheet: Sheet, key: K, value: Sheet[K]): void {
    (sheet as unknown as Record<string, unknown>)[key] = value;
    sheet.touch();
    this.bump();
  }

  // ----- sheets -----

  uniqueSheetName(base = 'Sheet'): string {
    let i = this.sheets.length + 1;
    const exists = (n: string) => !!this.sheetByName(n);
    if (base !== 'Sheet' && !exists(base)) return base;
    let name = `${base}${i}`;
    while (exists(name)) name = `${base}${++i}`;
    return name;
  }

  insertSheet(sheet: Sheet, index: number): void {
    this.sheets.splice(index, 0, sheet);
    this.bump();
    for (const l of this.listeners) l.sheetAdded(sheet, index);
  }

  removeSheet(sheet: Sheet): number {
    const idx = this.sheets.indexOf(sheet);
    if (idx < 0) return -1;
    this.sheets.splice(idx, 1);
    if (this.activeSheetId === sheet.id) {
      const vis = this.sheets.filter((s) => s.visibility === 'visible');
      this.activeSheetId = (this.sheets[Math.min(idx, this.sheets.length - 1)] ?? vis[0])?.id ?? '';
      const act = this.sheetById(this.activeSheetId);
      if (act && act.visibility !== 'visible' && vis[0]) this.activeSheetId = vis[0].id;
    }
    this.bump();
    for (const l of this.listeners) l.sheetRemoved(sheet);
    return idx;
  }

  renameSheet(sheet: Sheet, newName: string): void {
    const old = sheet.name;
    if (old === newName) return;
    sheet.name = newName;
    // update formula text referencing the sheet
    for (const s of this.sheets) {
      for (const [r, row] of s.rows) {
        for (const [c, cell] of row) {
          if (cell.f) {
            const nf = renameSheetInFormula(cell.f, old, newName);
            if (nf !== cell.f) row.set(c, { ...cell, f: nf });
            void r;
          }
        }
      }
      for (const ch of s.charts) {
        for (const se of ch.series) {
          se.values = renameSheetInFormula('=' + se.values, old, newName).slice(1);
          if (se.categories) se.categories = renameSheetInFormula('=' + se.categories, old, newName).slice(1);
          if (se.xValues) se.xValues = renameSheetInFormula('=' + se.xValues, old, newName).slice(1);
        }
      }
    }
    for (const n of this.names) n.ref = renameSheetInFormula('=' + n.ref, old, newName).slice(1);
    sheet.touch();
    this.bump();
    for (const l of this.listeners) l.sheetRenamed(sheet, old);
  }

  moveSheet(sheet: Sheet, toIndex: number): void {
    const from = this.sheets.indexOf(sheet);
    if (from < 0) return;
    this.sheets.splice(from, 1);
    this.sheets.splice(Math.max(0, Math.min(toIndex, this.sheets.length)), 0, sheet);
    this.bump();
    for (const l of this.listeners) l.sheetsReordered();
  }

  setNames(names: DefinedName[]): void {
    this.names = names;
    this.bump();
    for (const l of this.listeners) l.namesChanged();
  }

  // ----- structural ops -----

  /**
   * Insert (delta>0) or delete (delta<0) rows/cols at `at` on `sheet`.
   * Adjusts formulas across the workbook, names and sheet-level ranges.
   * Returns removed cells for deletes.
   */
  applyStructOp(sheet: Sheet, op: StructOp): Map<number, Map<number, Cell>> {
    const removed = op.axis === 'row' ? sheet.shiftRows(op.at, op.delta) : sheet.shiftCols(op.at, op.delta);
    // Formula text across all sheets
    for (const s of this.sheets) {
      for (const row of s.rows.values()) {
        for (const [c, cell] of row) {
          if (!cell.f) continue;
          const nf = adjustFormula(cell.f, s.name, sheet.name, op);
          if (nf !== cell.f) row.set(c, { ...cell, f: nf });
        }
      }
      // chart & sparkline references can point at any sheet
      for (const ch of s.charts) {
        for (const se of ch.series) {
          se.values = adjustRefString(se.values, s.name, sheet.name, op);
          if (se.categories) se.categories = adjustRefString(se.categories, s.name, sheet.name, op);
          if (se.xValues) se.xValues = adjustRefString(se.xValues, s.name, sheet.name, op);
        }
      }
      for (const g of s.sparklines) for (const it of g.items) it.ref = adjustRefString(it.ref, s.name, sheet.name, op);
      s.touch();
    }
    for (const n of this.names) {
      const host = n.scope ? this.sheetById(n.scope)?.name ?? sheet.name : sheet.name;
      n.ref = adjustRefString(n.ref, host, sheet.name, op);
    }
    // Sheet-level ranges
    const adj = (rg: Range) => adjustRange(rg, op);
    sheet.merges = sheet.merges.map(adj).filter((r): r is Range => !!r && !(r.r1 === r.r2 && r.c1 === r.c2));
    for (const cf of sheet.conditionalFormats) cf.ranges = cf.ranges.map(adj).filter((r): r is Range => !!r);
    sheet.conditionalFormats = sheet.conditionalFormats.filter((c) => c.ranges.length);
    for (const dv of sheet.validations) dv.ranges = dv.ranges.map(adj).filter((r): r is Range => !!r);
    sheet.validations = sheet.validations.filter((c) => c.ranges.length);
    sheet.tables = sheet.tables.filter((t) => {
      const nr = adj(t.range);
      if (!nr) return false;
      if (op.axis === 'col') {
        // keep columns list in sync with width
        const w = nr.c2 - nr.c1 + 1;
        if (w > t.columns.length) {
          const insertAt = Math.max(0, op.at - t.range.c1);
          const add = Array.from({ length: w - t.columns.length }, (_, i) => ({ name: uniqueColName(t.columns.map((c) => c.name), `Column${t.columns.length + i + 1}`) }));
          t.columns.splice(insertAt, 0, ...add);
        } else if (w < t.columns.length) {
          const delAt = Math.max(0, op.at - t.range.c1);
          t.columns.splice(delAt, t.columns.length - w);
        }
      }
      t.range = nr;
      return true;
    });
    if (sheet.autoFilter) {
      const nr = adj(sheet.autoFilter.range);
      if (!nr) sheet.autoFilter = null;
      else {
        const filters: typeof sheet.autoFilter.filters = {};
        for (const [k, v] of Object.entries(sheet.autoFilter.filters)) {
          const col = +k;
          if (op.axis === 'col') {
            const nc = adjustRange({ r1: 0, r2: 0, c1: col, c2: col }, op);
            if (nc) filters[nc.c1] = v;
          } else filters[col] = v;
        }
        sheet.autoFilter = { ...sheet.autoFilter, range: nr, filters };
      }
    }
    for (const pv of this.sheets.flatMap((s) => s.pivots)) {
      if (pv.sourceSheetId === sheet.id) {
        const nr = adj(pv.sourceRange);
        if (nr) pv.sourceRange = nr;
      }
    }
    if (sheet.pageSetup.printArea) {
      const nr = adj(sheet.pageSetup.printArea);
      sheet.pageSetup = { ...sheet.pageSetup, printArea: nr ?? undefined };
    }
    const shiftList = (list: number[]) =>
      list
        .map((b) => (op.delta > 0 ? (b >= op.at ? b + op.delta : b) : b < op.at ? b : b < op.at - op.delta ? -1 : b + op.delta))
        .filter((b) => b >= 0);
    if (op.axis === 'row') sheet.rowBreaks = shiftList(sheet.rowBreaks);
    else sheet.colBreaks = shiftList(sheet.colBreaks);
    // Sparkline locations & scenarios on this sheet
    for (const g of sheet.sparklines) {
      g.items = g.items.filter((it) => {
        const nr = adjustRange({ r1: it.r, c1: it.c, r2: it.r, c2: it.c }, op);
        if (!nr) return false;
        it.r = nr.r1;
        it.c = nr.c1;
        return true;
      });
    }
    sheet.touch();
    this.bump();
    for (const l of this.listeners) l.structureChanged(sheet, op);
    return removed;
  }

  // ----- serialization -----

  toJSON(): SerializedWorkbook {
    return {
      version: 1,
      styles: this.styles.toJSON(),
      names: this.names,
      props: this.props,
      activeSheetId: this.activeSheetId,
      sheets: this.sheets.map(serializeSheet),
    };
  }

  static fromJSON(data: SerializedWorkbook): Workbook {
    const wb = new Workbook();
    wb.styles = StyleTable.fromJSON(data.styles);
    wb.names = data.names ?? [];
    wb.props = data.props ?? {};
    wb.sheets = data.sheets.map(deserializeSheet);
    wb.activeSheetId = data.activeSheetId && wb.sheetById(data.activeSheetId) ? data.activeSheetId : wb.sheets[0]?.id ?? '';
    return wb;
  }
}

function uniqueColName(existing: string[], base: string): string {
  let n = base;
  let i = 2;
  while (existing.includes(n)) n = `${base}${i++}`;
  return n;
}

export interface SerializedSheet {
  id: string;
  name: string;
  cells: [number, number, Cell][];
  meta: Record<string, unknown>;
}

export interface SerializedWorkbook {
  version: 1;
  styles: CellStyle[];
  names: DefinedName[];
  props: WorkbookProps;
  activeSheetId: string;
  sheets: SerializedSheet[];
}

function serializeSheet(s: Sheet): SerializedSheet {
  const cells: [number, number, Cell][] = [];
  for (const [r, row] of s.rows) for (const [c, cell] of row) cells.push([r, c, cell]);
  const meta: Record<string, unknown> = {};
  for (const k of META_KEYS) {
    const v = (s as unknown as Record<string, unknown>)[k];
    if (v instanceof Map) meta[k] = { __map: [...v.entries()] };
    else if (v instanceof Set) meta[k] = { __set: [...v] };
    else meta[k] = v;
  }
  return { id: s.id, name: s.name, cells, meta };
}

function deserializeSheet(d: SerializedSheet): Sheet {
  const s = new Sheet(d.name, d.id);
  for (const [r, c, cell] of d.cells) {
    let row = s.rows.get(r);
    if (!row) s.rows.set(r, (row = new Map()));
    row.set(c, cell);
  }
  for (const k of META_KEYS) {
    const v = d.meta[k] as { __map?: [number, number][]; __set?: number[] } | undefined;
    if (v === undefined) continue;
    let val: unknown = v;
    if (v && typeof v === 'object' && '__map' in v) val = new Map(v.__map);
    else if (v && typeof v === 'object' && '__set' in v) val = new Set(v.__set);
    (s as unknown as Record<string, unknown>)[k] = val;
  }
  s.touch();
  return s;
}

export function cloneSheet(src: Sheet, name: string): Sheet {
  const data = serializeSheet(src);
  const s = deserializeSheet(structuredClone({ ...data, id: newId('sh'), name }));
  // new ids for owned objects
  for (const cf of s.conditionalFormats) cf.id = newId('cf');
  for (const dv of s.validations) dv.id = newId('dv');
  for (const ch of s.charts) ch.id = newId('ch');
  s.tables = []; // table names must be unique; Excel renames — we drop the structure but keep cells
  s.pivots = [];
  return s;
}

export function metaSnapshot(s: Sheet, keys: MetaKey[] = META_KEYS): Partial<Record<MetaKey, unknown>> {
  const o: Partial<Record<MetaKey, unknown>> = {};
  for (const k of keys) o[k] = cloneValue((s as unknown as Record<string, unknown>)[k]);
  return o;
}

export function restoreMeta(s: Sheet, snap: Partial<Record<MetaKey, unknown>>): void {
  for (const [k, v] of Object.entries(snap)) (s as unknown as Record<string, unknown>)[k] = cloneValue(v);
  s.touch();
}

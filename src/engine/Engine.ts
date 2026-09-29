// HyperFormula wrapper. The only module that talks to HyperFormula.
import { DetailedCellError, HyperFormula, RawCellContent } from 'hyperformula';
import type { Sheet } from '../model/sheet';
import type { StructOp } from '../model/structure';
import type { Cell, CellValue } from '../model/types';
import type { CellChange, Workbook, WorkbookListener } from '../model/workbook';

export interface ErrorVal {
  error: string;
  circular?: boolean;
  message?: string;
}

export type ComputedValue = CellValue | ErrorVal;

export function isErrorVal(v: unknown): v is ErrorVal {
  return typeof v === 'object' && v !== null && 'error' in v;
}

const HF_CONFIG = {
  licenseKey: 'gpl-v3',
  maxRows: 1048576,
  maxColumns: 16384,
  functionArgSeparator: ',',
  decimalSeparator: '.' as const,
  thousandSeparator: '' as const,
  evaluateNullToZero: true,
  useArrayArithmetic: true,
  smartRounding: true,
  caseSensitive: false,
  useWildcards: true,
  useRegularExpressions: false,
  matchWholeCell: true,
  dateFormats: ['MM/DD/YYYY', 'MM/DD/YY', 'YYYY-MM-DD'],
  timeFormats: ['hh:mm', 'hh:mm:ss.sss'],
  currencySymbol: ['$'],
  localeLang: 'en-US',
};

export function cellToRaw(cell: Cell | undefined): RawCellContent {
  if (!cell) return null;
  if (cell.f) return cell.f;
  const v = cell.v;
  if (v === undefined || v === null) return null;
  if (typeof v === 'string') {
    if (cell.e) return v; // error literal like #N/A
    return v === '' ? null : "'" + v;
  }
  return v;
}

export class Engine implements WorkbookListener {
  private hf: HyperFormula | null = null;
  private ids = new Map<string, number>();
  /** Cells reported with a circular reference error (for the status bar). */
  circular: { sheet: string; r: number; c: number } | null = null;
  /** Bumps whenever computed values may have changed. */
  rev = 0;
  onChange?: () => void;

  constructor(public wb: Workbook) {
    wb.listeners.push(this);
  }

  dispose(): void {
    this.wb.listeners = this.wb.listeners.filter((l) => l !== this);
    this.hf?.destroy();
    this.hf = null;
  }

  get attached(): boolean {
    return this.hf !== null;
  }

  /** Whether the workbook needs a calculation engine at all. */
  workbookNeedsEngine(): boolean {
    if (this.wb.names.length) return true;
    for (const s of this.wb.sheets) for (const row of s.rows.values()) for (const cell of row.values()) if (cell.f) return true;
    return false;
  }

  attachIfNeeded(): void {
    if (!this.hf && this.workbookNeedsEngine()) this.attach();
  }

  attach(): void {
    if (this.hf) this.hf.destroy();
    const sheets: Record<string, RawCellContent[][]> = {};
    for (const s of this.wb.sheets) sheets[s.name] = [];
    const hf = HyperFormula.buildFromSheets(sheets, HF_CONFIG);
    this.hf = hf;
    this.ids.clear();
    for (const s of this.wb.sheets) this.ids.set(s.id, hf.getSheetId(s.name)!);
    hf.suspendEvaluation();
    try {
      this.syncNames();
      for (const s of this.wb.sheets) {
        const sid = this.ids.get(s.id)!;
        for (const [r, row] of s.rows) {
          for (const [c, cell] of row) {
            const raw = cellToRaw(cell);
            if (raw === null) continue;
            try {
              hf.setCellContents({ sheet: sid, row: r, col: c }, raw);
            } catch {
              hf.setCellContents({ sheet: sid, row: r, col: c }, '#NAME?');
            }
          }
        }
      }
    } finally {
      hf.resumeEvaluation();
    }
    this.bump();
  }

  private bump(): void {
    this.rev++;
    this.circular = null;
    this.onChange?.();
  }

  private sid(sheet: Sheet): number | undefined {
    return this.ids.get(sheet.id);
  }

  // ----- WorkbookListener -----

  cellsChanged(sheet: Sheet, changes: CellChange[]): void {
    if (!this.hf) {
      if (changes.some((ch) => sheet.getCell(ch.r, ch.c)?.f)) this.attach();
      return;
    }
    const hf = this.hf;
    const sid = this.sid(sheet);
    if (sid === undefined) return;
    const doSet = () => {
      for (const ch of changes) {
        const raw = cellToRaw(sheet.getCell(ch.r, ch.c));
        try {
          hf.setCellContents({ sheet: sid, row: ch.r, col: ch.c }, raw);
        } catch {
          hf.setCellContents({ sheet: sid, row: ch.r, col: ch.c }, '#NAME?');
        }
      }
    };
    if (changes.length > 1) hf.batch(doSet);
    else doSet();
    this.bump();
  }

  structureChanged(sheet: Sheet, op: StructOp): void {
    if (!this.hf) return;
    const sid = this.sid(sheet);
    if (sid === undefined) return;
    const n = Math.abs(op.delta);
    if (op.axis === 'row') {
      if (op.delta > 0) this.hf.addRows(sid, [op.at, n]);
      else this.hf.removeRows(sid, [op.at, n]);
    } else {
      if (op.delta > 0) this.hf.addColumns(sid, [op.at, n]);
      else this.hf.removeColumns(sid, [op.at, n]);
    }
    this.bump();
  }

  sheetAdded(sheet: Sheet): void {
    if (!this.hf) {
      this.attachIfNeeded();
      return;
    }
    const hf = this.hf;
    const name = hf.addSheet(sheet.name);
    const sid = hf.getSheetId(name)!;
    this.ids.set(sheet.id, sid);
    hf.batch(() => {
      for (const [r, row] of sheet.rows) for (const [c, cell] of row) {
        const raw = cellToRaw(cell);
        if (raw !== null) {
          try {
            hf.setCellContents({ sheet: sid, row: r, col: c }, raw);
          } catch {
            /* ignore invalid */
          }
        }
      }
    });
    this.bump();
  }

  sheetRemoved(sheet: Sheet): void {
    if (!this.hf) return;
    const sid = this.sid(sheet);
    if (sid !== undefined) this.hf.removeSheet(sid);
    this.ids.delete(sheet.id);
    this.bump();
  }

  sheetRenamed(sheet: Sheet): void {
    if (!this.hf) return;
    const sid = this.sid(sheet);
    if (sid !== undefined) this.hf.renameSheet(sid, sheet.name);
    this.bump();
  }

  sheetsReordered(): void {
    /* HyperFormula addresses sheets by name; order does not matter. */
  }

  namesChanged(): void {
    if (!this.hf) {
      this.attachIfNeeded();
      return;
    }
    this.syncNames();
    this.bump();
  }

  private syncNames(): void {
    const hf = this.hf!;
    for (const ne of hf.getAllNamedExpressionsSerialized()) {
      try {
        hf.removeNamedExpression(ne.name, ne.scope);
      } catch {
        /* ignore */
      }
    }
    for (const n of this.wb.names) {
      const scope = n.scope ? this.ids.get(n.scope) : undefined;
      try {
        hf.addNamedExpression(n.name, '=' + n.ref, scope);
      } catch {
        /* invalid name – ignored */
      }
    }
  }

  // ----- queries -----

  getValue(sheet: Sheet, r: number, c: number): ComputedValue {
    const cell = sheet.getCell(r, c);
    if (!this.hf) {
      if (!cell) return null;
      if (cell.e && typeof cell.v === 'string') return { error: cell.v };
      return cell.v ?? null;
    }
    const sid = this.sid(sheet);
    if (sid === undefined) return null;
    const v = this.hf.getCellValue({ sheet: sid, row: r, col: c });
    return this.convert(v, sheet, r, c);
  }

  private convert(v: unknown, sheet?: Sheet, r?: number, c?: number): ComputedValue {
    if (v instanceof DetailedCellError) {
      if (v.type === 'CYCLE') {
        if (sheet && r !== undefined && c !== undefined) this.circular = { sheet: sheet.name, r, c };
        return { error: '#CYCLE!', circular: true };
      }
      let code = v.value;
      if (code === '#ERROR!' || code === '#LIC!') code = v.message?.includes('Parsing') ? '#NAME?' : '#VALUE!';
      return { error: code, message: v.message };
    }
    if (v === undefined) return null;
    return v as CellValue;
  }

  private circCache: { rev: number; sheetId: string; hit: { sheet: string; r: number; c: number } | null } | null = null;

  /** First formula cell on `sheet` caught in a circular reference (for the status bar). */
  findCircular(sheet: Sheet): { sheet: string; r: number; c: number } | null {
    if (!this.hf) return null;
    if (this.circCache && this.circCache.rev === this.rev && this.circCache.sheetId === sheet.id) return this.circCache.hit;
    const sid = this.sid(sheet);
    let hit: { sheet: string; r: number; c: number } | null = null;
    let n = 0;
    if (sid !== undefined) {
      outer: for (const [r, row] of sheet.rows) {
        for (const [c, cell] of row) {
          if (!cell.f) continue;
          if (++n > 50000) break outer;
          const v = this.hf.getCellValue({ sheet: sid, row: r, col: c });
          if (v instanceof DetailedCellError && v.type === 'CYCLE') {
            hit = { sheet: sheet.name, r, c };
            break outer;
          }
        }
      }
    }
    this.circCache = { rev: this.rev, sheetId: sheet.id, hit };
    return hit;
  }

  /** Detailed type used to auto-apply date/percent/currency formats to formula results. */
  valueType(sheet: Sheet, r: number, c: number): string | undefined {
    if (!this.hf) return undefined;
    const sid = this.sid(sheet);
    if (sid === undefined) return undefined;
    try {
      return this.hf.getCellValueDetailedType({ sheet: sid, row: r, col: c });
    } catch {
      return undefined;
    }
  }

  /** Evaluates an arbitrary formula in the context of a sheet (for CF, validation, goal seek...). */
  evaluate(formula: string, sheet: Sheet): ComputedValue {
    if (!this.hf) this.attach();
    const sid = this.sid(sheet);
    if (sid === undefined) return null;
    try {
      const f = formula.startsWith('=') ? formula : '=' + formula;
      const v = this.hf!.calculateFormula(f, sid);
      if (Array.isArray(v)) return this.convert(v[0]?.[0]);
      return this.convert(v);
    } catch (e) {
      return { error: '#NAME?', message: String(e) };
    }
  }

  /** Evaluates to a 2D array (for list validation sources, charts on names...). */
  evaluateArray(formula: string, sheet: Sheet): ComputedValue[][] {
    if (!this.hf) this.attach();
    const sid = this.sid(sheet);
    if (sid === undefined) return [];
    try {
      const f = formula.startsWith('=') ? formula : '=' + formula;
      const v = this.hf!.calculateFormula(f, sid);
      if (Array.isArray(v)) return v.map((row) => row.map((x) => this.convert(x)));
      return [[this.convert(v)]];
    } catch {
      return [];
    }
  }

  validateFormula(formula: string): boolean {
    if (!this.hf) this.attach();
    try {
      return this.hf!.validateFormula(formula);
    } catch {
      return false;
    }
  }

  isFunctionSupported(name: string): boolean {
    return Engine.functionNames().includes(name.toUpperCase());
  }

  private static fnCache: string[] | null = null;
  static functionNames(): string[] {
    if (!this.fnCache) this.fnCache = HyperFormula.getRegisteredFunctionNames('enGB').slice().sort();
    return this.fnCache;
  }

  /** Recalculate everything (F9). */
  recalc(): void {
    if (this.hf) this.hf.rebuildAndRecalculate();
    this.bump();
  }

  /** Names HyperFormula considers invalid are rejected up-front. */
  isValidName(name: string): boolean {
    if (!this.hf) this.attach();
    try {
      return this.hf!.isItPossibleToAddNamedExpression(name, '=1');
    } catch {
      return false;
    }
  }
}

// Command layer: every mutation is a Command so undo/redo covers everything.
import { cloneValue, META_KEYS, MetaKey, Sheet } from './sheet';
import type { StructOp } from './structure';
import type { Cell, DefinedName } from './types';
import { metaSnapshot, restoreMeta, Workbook } from './workbook';

export interface Command {
  label: string;
  apply(wb: Workbook): void;
  revert(wb: Workbook): void;
}

export interface CellDelta {
  r: number;
  c: number;
  before: Cell | undefined;
  after: Cell | undefined;
}

function need(wb: Workbook, id: string): Sheet {
  const s = wb.sheetById(id);
  if (!s) throw new Error(`Sheet ${id} not found`);
  return s;
}

export class CellsCommand implements Command {
  constructor(public label: string, public sheetId: string, public deltas: CellDelta[]) {}
  apply(wb: Workbook): void {
    wb.setCells(need(wb, this.sheetId), this.deltas.map((d) => ({ r: d.r, c: d.c, cell: d.after })));
  }
  revert(wb: Workbook): void {
    wb.setCells(need(wb, this.sheetId), this.deltas.map((d) => ({ r: d.r, c: d.c, cell: d.before })));
  }
}

export class MetaCommand<K extends MetaKey> implements Command {
  constructor(public label: string, public sheetId: string, public key: K, public before: Sheet[K], public after: Sheet[K]) {}
  apply(wb: Workbook): void {
    wb.setMeta(need(wb, this.sheetId), this.key, cloneValue(this.after));
  }
  revert(wb: Workbook): void {
    wb.setMeta(need(wb, this.sheetId), this.key, cloneValue(this.before));
  }
}

/** Closure-based command for workbook-level operations (sheets, names...). */
export class FnCommand implements Command {
  constructor(public label: string, private doFn: (wb: Workbook) => void, private undoFn: (wb: Workbook) => void) {}
  apply(wb: Workbook): void {
    this.doFn(wb);
  }
  revert(wb: Workbook): void {
    this.undoFn(wb);
  }
}

export class CompositeCommand implements Command {
  constructor(public label: string, public cmds: Command[]) {}
  apply(wb: Workbook): void {
    for (const c of this.cmds) c.apply(wb);
  }
  revert(wb: Workbook): void {
    for (let i = this.cmds.length - 1; i >= 0; i--) this.cmds[i].revert(wb);
  }
}

const STRUCT_META: MetaKey[] = META_KEYS.filter((k) => k !== 'zoom' && k !== 'showGridlines' && k !== 'showHeaders');

export class StructCommand implements Command {
  private formulas: [string, number, number, string][] | null = null;
  private names: DefinedName[] | null = null;
  private metas: [string, Partial<Record<MetaKey, unknown>>][] | null = null;
  private removed: Map<number, Map<number, Cell>> | null = null;

  constructor(public label: string, public sheetId: string, public op: StructOp) {}

  apply(wb: Workbook): void {
    const sheet = need(wb, this.sheetId);
    if (!this.metas) {
      this.metas = wb.sheets.map((s) => [s.id, metaSnapshot(s, STRUCT_META)]);
      this.names = structuredClone(wb.names);
      if (this.op.delta < 0) {
        const f: [string, number, number, string][] = [];
        for (const s of wb.sheets) for (const [r, row] of s.rows) for (const [c, cell] of row) if (cell.f) f.push([s.id, r, c, cell.f]);
        this.formulas = f;
      }
    }
    const removed = wb.applyStructOp(sheet, this.op);
    if (this.op.delta < 0) this.removed = removed;
  }

  revert(wb: Workbook): void {
    const sheet = need(wb, this.sheetId);
    wb.applyStructOp(sheet, { ...this.op, delta: -this.op.delta });
    const changed = new Map<string, { r: number; c: number; cell: Cell | undefined }[]>();
    const push = (sid: string, r: number, c: number, cell: Cell | undefined) => {
      let l = changed.get(sid);
      if (!l) changed.set(sid, (l = []));
      l.push({ r, c, cell });
    };
    if (this.op.delta < 0 && this.removed) {
      for (const [r, row] of this.removed) for (const [c, cell] of row) push(sheet.id, r, c, cell);
    }
    if (this.formulas) {
      for (const [sid, r, c, f] of this.formulas) {
        const s = wb.sheetById(sid);
        if (!s) continue;
        const cur = s.getCell(r, c);
        if (cur && cur.f !== f) push(sid, r, c, { ...cur, f });
      }
    }
    if (this.metas) for (const [sid, snap] of this.metas) {
      const s = wb.sheetById(sid);
      if (s) restoreMeta(s, snap);
    }
    if (this.names) wb.setNames(structuredClone(this.names));
    for (const [sid, list] of changed) {
      const s = wb.sheetById(sid);
      if (s) wb.setCells(s, list);
    }
  }
}

export interface HistoryEntry {
  cmd: Command;
  /** Opaque UI state (selection) captured before/after, restored on undo/redo. */
  uiBefore?: unknown;
  uiAfter?: unknown;
}

export class History {
  undoStack: HistoryEntry[] = [];
  redoStack: HistoryEntry[] = [];
  limit = 100;

  push(e: HistoryEntry): void {
    this.undoStack.push(e);
    if (this.undoStack.length > this.limit) this.undoStack.shift();
    this.redoStack = [];
  }

  undo(wb: Workbook): HistoryEntry | undefined {
    const e = this.undoStack.pop();
    if (!e) return;
    e.cmd.revert(wb);
    this.redoStack.push(e);
    return e;
  }

  redo(wb: Workbook): HistoryEntry | undefined {
    const e = this.redoStack.pop();
    if (!e) return;
    e.cmd.apply(wb);
    this.undoStack.push(e);
    return e;
  }

  clear(): void {
    this.undoStack = [];
    this.redoStack = [];
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }
  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }
}

/**
 * Transaction: applies mutations immediately while recording commands.
 * Cell writes are batched per sheet and flushed to the engine in one go.
 */
export class Tx {
  cmds: Command[] = [];
  private pending = new Map<string, Map<string, CellDelta>>();
  /** Every cell written in this transaction as it was before (sheet id → row → col → cell). */
  readonly before = new Map<string, Map<number, Map<number, Cell | undefined>>>();
  /** Runs once before a transaction is finalised (the state layer re-fits row heights here). */
  static beforeFinish: ((tx: Tx) => void) | null = null;
  private finishing = false;

  constructor(public wb: Workbook, public label: string) {}

  cell(sheet: Sheet, r: number, c: number): Cell | undefined {
    return sheet.getCell(r, c);
  }

  setCell(sheet: Sheet, r: number, c: number, cell: Cell | undefined): void {
    let m = this.pending.get(sheet.id);
    if (!m) this.pending.set(sheet.id, (m = new Map()));
    const key = `${r},${c}`;
    const prev = m.get(key);
    const before = prev ? prev.before : sheet.getCell(r, c);
    m.set(key, { r, c, before, after: cell });
    let bs = this.before.get(sheet.id);
    if (!bs) this.before.set(sheet.id, (bs = new Map()));
    let br = bs.get(r);
    if (!br) bs.set(r, (br = new Map()));
    if (!br.has(c)) br.set(c, before);
    sheet.setCellRaw(r, c, cell);
  }

  /** Patch fields of a cell (undefined values delete the field). */
  patchCell(sheet: Sheet, r: number, c: number, patch: Partial<Cell>): void {
    const cur = sheet.getCell(r, c) ?? {};
    const next: Cell = { ...cur };
    for (const [k, v] of Object.entries(patch)) {
      if (v === undefined) delete (next as Record<string, unknown>)[k];
      else (next as Record<string, unknown>)[k] = v;
    }
    this.setCell(sheet, r, c, next);
  }

  flush(): void {
    for (const [sid, m] of this.pending) {
      if (!m.size) continue;
      const deltas = [...m.values()].filter((d) => d.before !== d.after);
      const sheet = this.wb.sheetById(sid);
      if (!sheet || !deltas.length) continue;
      const cmd = new CellsCommand(this.label, sid, deltas);
      // cells are already written raw; notify listeners
      this.wb.setCells(sheet, deltas.map((d) => ({ r: d.r, c: d.c, cell: d.after })));
      this.cmds.push(cmd);
    }
    this.pending.clear();
  }

  setMeta<K extends MetaKey>(sheet: Sheet, key: K, value: Sheet[K]): void {
    this.flush();
    const before = cloneValue(sheet[key]);
    const cmd = new MetaCommand(this.label, sheet.id, key, before, cloneValue(value));
    cmd.apply(this.wb);
    this.cmds.push(cmd);
  }

  run(cmd: Command): void {
    this.flush();
    cmd.apply(this.wb);
    this.cmds.push(cmd);
  }

  /** Finalises the transaction; returns the composite command or null if nothing changed. */
  finish(): Command | null {
    if (Tx.beforeFinish && !this.finishing) {
      this.finishing = true;
      Tx.beforeFinish(this);
    }
    this.flush();
    if (!this.cmds.length) return null;
    return this.cmds.length === 1 ? Object.assign(this.cmds[0], { label: this.label }) : new CompositeCommand(this.label, this.cmds);
  }
}

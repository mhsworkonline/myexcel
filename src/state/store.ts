import { create } from 'zustand';
import { Engine } from '../engine/Engine';
import { CellAddr, Range } from '../model/address';
import { History, Tx } from '../model/commands';
import { Selection, singleSel } from '../model/selection';
import type { Sheet } from '../model/sheet';
import { Workbook } from '../model/workbook';
import { setDocDefaultFont } from '../model/styles';

export type RibbonTab = 'File' | 'Home' | 'Insert' | 'Page Layout' | 'Formulas' | 'Data' | 'Review' | 'View';

export interface EditState {
  sheetId: string;
  r: number;
  c: number;
  text: string;
  caret: number;
  /** enter: arrows commit + move; edit: arrows move caret; point: arrows pick references */
  mode: 'enter' | 'edit';
  fromBar: boolean;
  /** Reference being inserted by pointing (text span it occupies). */
  point?: { start: number; end: number; anchor: CellAddr; corner: CellAddr; sheetId: string };
  /** Sheet id the formula is being edited on (for cross-sheet pointing). */
  hostSheetId: string;
}

export interface ClipState {
  sheetId: string;
  range: Range;
  cut: boolean;
}

export interface PainterState {
  sheetId: string;
  range: Range;
  sticky: boolean;
}

export interface MenuState {
  x: number;
  y: number;
  kind: 'cell' | 'row' | 'col' | 'tab' | 'chart' | 'corner';
  sheetId?: string;
  chartId?: string;
}

export interface DialogState {
  type: string;
  props?: Record<string, unknown>;
}

export interface FileState {
  name: string;
  handle?: unknown;
  format: 'xlsx' | 'csv' | 'ods' | 'xls' | 'json';
  dirty: boolean;
  lastSaved?: number;
  saving: boolean;
  autosaved?: number;
}

export interface AppState {
  wb: Workbook;
  engine: Engine;
  history: History;
  rev: number;
  file: FileState;
  sel: Selection;
  sheetSel: Record<string, Selection>;
  edit: EditState | null;
  clip: ClipState | null;
  painter: PainterState | null;
  ribbonTab: RibbonTab;
  ribbonCollapsed: boolean;
  backstage: boolean;
  /** Page the File backstage opens on (e.g. "open" for Open Recent). */
  backstagePage: string | null;
  dialog: DialogState | null;
  menu: MenuState | null;
  theme: 'light' | 'dark';
  showFormulas: boolean;
  status: string | null;
  viewMode: 'normal' | 'pageBreak' | 'pageLayout';
  scrollTo: { r: number; c: number; seq: number } | null;
  formulaBarExpanded: boolean;
  selectedChartId: string | null;
  /** Filter dropdown open for column (absolute col index). */
  filterMenu: { col: number; x: number; y: number } | null;
  /** Validation list dropdown open. */
  listMenu: { r: number; c: number } | null;
  extendMode: boolean;
  addMode: boolean;
  pivotPanel: string | null;
}

function initialWorkbook(): { wb: Workbook; engine: Engine } {
  const wb = Workbook.createDefault();
  const engine = new Engine(wb);
  return { wb, engine };
}

const init = initialWorkbook();

export const useStore = create<AppState>(() => ({
  wb: init.wb,
  engine: init.engine,
  history: new History(),
  rev: 0,
  file: { name: 'Book1', format: 'xlsx', dirty: false, saving: false },
  sel: singleSel(0, 0),
  sheetSel: {},
  edit: null,
  clip: null,
  painter: null,
  ribbonTab: 'Home',
  ribbonCollapsed: false,
  backstage: false,
  backstagePage: null,
  dialog: null,
  menu: null,
  theme: 'light',
  showFormulas: false,
  status: null,
  viewMode: 'normal',
  scrollTo: null,
  formulaBarExpanded: false,
  selectedChartId: null,
  filterMenu: null,
  listMenu: null,
  extendMode: false,
  addMode: false,
  pivotPanel: null,
}));

export const S = () => useStore.getState();
export const setState = useStore.setState;

export function bump(): void {
  useStore.setState((s) => ({ rev: s.rev + 1 }));
}

export function activeSheet(): Sheet {
  return S().wb.activeSheet;
}

let scrollSeq = 0;
export function requestScroll(r: number, c: number): void {
  setState({ scrollTo: { r, c, seq: ++scrollSeq } });
}

export function setStatus(msg: string | null): void {
  setState({ status: msg });
}

export function markDirty(): void {
  const f = S().file;
  if (!f.dirty) setState({ file: { ...f, dirty: true } });
}

interface UiSnap {
  sheetId: string;
  sel: Selection;
}

/**
 * Run a mutation as one undoable step. `fn` receives the transaction and the active sheet.
 * Returns false when nothing changed.
 */
export function transact(label: string, fn: (tx: Tx, sheet: Sheet) => void | false, selAfter?: Selection): boolean {
  const st = S();
  const sheet = st.wb.activeSheet;
  const tx = new Tx(st.wb, label);
  const before: UiSnap = { sheetId: sheet.id, sel: st.sel };
  const res = fn(tx, sheet);
  const cmd = tx.finish();
  if (res === false && !cmd) return false;
  if (cmd) {
    const after: UiSnap = { sheetId: S().wb.activeSheetId, sel: selAfter ?? S().sel };
    st.history.push({ cmd, uiBefore: before, uiAfter: after });
    markDirty();
  }
  if (selAfter) setState({ sel: selAfter });
  bump();
  return !!cmd;
}

export function undo(): void {
  const st = S();
  if (st.edit) return;
  const e = st.history.undo(st.wb);
  if (!e) return;
  restoreUi(e.uiBefore as UiSnap | undefined);
  markDirty();
  setStatus(null);
  bump();
}

export function redo(): void {
  const st = S();
  if (st.edit) return;
  const e = st.history.redo(st.wb);
  if (!e) return;
  restoreUi(e.uiAfter as UiSnap | undefined);
  markDirty();
  bump();
}

function restoreUi(ui: UiSnap | undefined): void {
  if (!ui) return;
  const st = S();
  if (st.wb.sheetById(ui.sheetId)) {
    if (st.wb.activeSheetId !== ui.sheetId) st.wb.activeSheetId = ui.sheetId;
    setState({ sel: ui.sel, clip: null });
    requestScroll(ui.sel.active.r, ui.sel.active.c);
  }
}

/** Replace the whole workbook (open file / new). */
export function loadWorkbook(wb: Workbook, file: Partial<FileState>): void {
  const st = S();
  setDocDefaultFont(wb.props.defaultFont);
  st.engine.dispose();
  const engine = new Engine(wb);
  engine.attachIfNeeded();
  engine.onChange = () => bump();
  st.history.clear();
  setState({
    wb,
    engine,
    sel: singleSel(0, 0),
    sheetSel: {},
    edit: null,
    clip: null,
    painter: null,
    dialog: null,
    menu: null,
    selectedChartId: null,
    file: { name: 'Book1', format: 'xlsx', dirty: false, saving: false, ...file },
    rev: st.rev + 1,
  });
}

export function openDialog(type: string, props?: Record<string, unknown>): void {
  setState({ dialog: { type, props }, menu: null });
}

export function closeDialog(): void {
  setState({ dialog: null });
}

export function alertBox(message: string, title = 'MyExcel', icon: 'info' | 'warning' | 'error' = 'warning'): void {
  openDialog('alert', { message, title, icon });
}

init.engine.onChange = () => bump();

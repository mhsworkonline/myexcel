import { isErrorVal } from '../../engine/Engine';
import { readCsv, writeCsv } from '../../io/csv';
import { fileAdapter, FILTERS, FileTypeFilter } from '../../io/FileAdapter';
import { Workbook } from '../../model/workbook';
import { settings } from '../settings';
import { alertBox, FileState, loadWorkbook, openDialog, S, setState, setStatus } from '../store';
import { getComputed } from '../values';
import { displayTextAt } from './sortFilter';

export type SaveFormat = 'xlsx' | 'csv' | 'tsv' | 'ods' | 'pdf';

// ---- Windows (desktop): one workbook per window, like Excel ----

export interface WindowOps {
  /** Open a new window with this file, or with a blank workbook. */
  open(path?: string): Promise<void>;
  /** Close this window (the unsaved-changes prompt runs first). */
  close(): Promise<void>;
}
let windows: WindowOps | null = null;
/** Set by the desktop bridge; the web build keeps one workbook per tab. */
export function setWindowOps(ops: WindowOps): void {
  windows = ops;
}

/** An untouched new workbook, which opening a file replaces instead of opening another window. */
export function isPristine(): boolean {
  const st = S();
  return !st.file.dirty && st.file.handle === undefined && !st.file.lastSaved && st.history.undoStack.length === 0 && st.wb.sheets.every((sh) => !sh.usedRange());
}

/** Web: replacing the open workbook asks first when it has unsaved changes. */
function afterDiscardCheck(proceed: () => void): void {
  if (!S().file.dirty) return proceed();
  openDialog('confirm', { message: `Want to save your changes to '${S().file.name}'?

Unsaved changes will be lost.`, okLabel: "Don't Save", onOk: proceed });
}

/** Open a file by path: in this window if it's an untouched new workbook, otherwise in a new window. Returns false if the caller should open it here. */
export function openPathInNewWindow(path: string): boolean {
  if (!windows || isPristine()) return false;
  void windows.open(path);
  return true;
}

/** File > Close / Ctrl+W: closes the window on desktop; on the web it starts a new blank workbook. */
export function closeWorkbook(): void {
  if (windows) void windows.close();
  else newWorkbook();
}

export function newWorkbook(confirmed = false): void {
  // Excel opens every new workbook in its own window
  if (windows && !confirmed) {
    void windows.open();
    return;
  }
  if (S().file.dirty && !confirmed) {
    openDialog('confirm', { message: `Want to save your changes to '${S().file.name}'?\n\nUnsaved changes will be lost.`, okLabel: "Don't Save", onOk: () => newWorkbook(true) });
    return;
  }
  const n = ((window as unknown as { __bookSeq?: number }).__bookSeq = ((window as unknown as { __bookSeq?: number }).__bookSeq ?? 1) + 1);
  const st = settings();
  loadWorkbook(Workbook.createDefault({ defaultFont: st.defaultFont, zoom: st.defaultZoom }), { name: `Book${n}`, format: 'xlsx' });
  clearAutosave();
}

function ext(name: string): string {
  return (/\.([^.\\/]+)$/.exec(name)?.[1] ?? '').toLowerCase();
}

export async function parseFile(name: string, data: ArrayBuffer): Promise<{ wb: Workbook; format: FileState['format'] }> {
  const e = ext(name);
  if (e === 'csv' || e === 'tsv' || e === 'txt') {
    const text = new TextDecoder('utf-8').decode(data);
    return { wb: readCsv(text, name), format: 'csv' };
  }
  if (e === 'ods') {
    const { readOds } = await import('../../io/ods');
    return { wb: await readOds(data), format: 'ods' };
  }
  if (e === 'xls') {
    const { readXls } = await import('../../io/xls');
    return { wb: await readXls(data), format: 'xls' };
  }
  const { readXlsxFull } = await import('../../io/xlsxExtras');
  return { wb: await readXlsxFull(data), format: 'xlsx' };
}

export const OPEN_FILTERS: FileTypeFilter[] = [
  { description: 'All Spreadsheets', extensions: ['xlsx', 'xlsm', 'xls', 'csv', 'tsv', 'txt', 'ods'], mime: FILTERS.xlsx.mime },
  FILTERS.xlsx,
  FILTERS.xls,
  FILTERS.csv,
  FILTERS.ods,
];

export async function openFile(): Promise<void> {
  const f = await fileAdapter().open(OPEN_FILTERS);
  if (!f) return;
  if (typeof f.handle === 'string' && openPathInNewWindow(f.handle)) return;
  afterDiscardCheck(() => void openFromData(f.name, f.data, f.handle));
}

/** A file dropped on the window (no path available): replaces the workbook after the unsaved-changes check. */
export function openDropped(name: string, data: ArrayBuffer): void {
  afterDiscardCheck(() => void openFromData(name, data));
}

export async function openFromData(name: string, data: ArrayBuffer, handle?: unknown): Promise<void> {
  setStatus('Opening…');
  try {
    const { wb, format } = await parseFile(name, data);
    const base = name.replace(/\.[^.]+$/, '');
    // .xls is read-only: keep no handle so Save goes through Save As (.xlsx), like Excel's compatibility mode
    const keep = format !== 'xls' && (typeof handle === 'string' || format === 'xlsx' || format === 'csv');
    loadWorkbook(wb, { name: base, handle: keep ? handle : undefined, format, dirty: false, lastSaved: Date.now() });
    setState({ backstage: false });
    fileAdapter().addRecent({ name, path: typeof handle === 'string' ? handle : undefined, opened: Date.now() });
    setStatus(null);
  } catch (e) {
    setStatus(null);
    alertBox(`We found a problem with some content in '${name}'.\n\n${(e as Error).message}`, 'MyExcel', 'error');
  }
}

export function textAtForExport(sheetIdx: number) {
  const sheet = S().wb.sheets[sheetIdx];
  return (r: number, c: number) => displayTextAt(sheet, r, c);
}

export async function serialize(format: SaveFormat): Promise<Uint8Array | string> {
  const st = S();
  const wb = st.wb;
  if (format === 'csv' || format === 'tsv') {
    const idx = wb.sheets.indexOf(wb.activeSheet);
    return '﻿' + writeCsv(wb.activeSheet, textAtForExport(idx), format === 'tsv' ? '\t' : ',');
  }
  if (format === 'ods') {
    const { writeOds } = await import('../../io/ods');
    return writeOds(wb, (s, r, c) => getComputed(s, r, c));
  }
  const { writeXlsxFull } = await import('../../io/xlsxExtras');
  const buf = await writeXlsxFull(wb, (s, r, c) => {
    const v = getComputed(s, r, c);
    return isErrorVal(v) ? { error: v.error } : v;
  });
  return new Uint8Array(buf);
}

/** Format to write when saving in place: from the file's own extension (desktop paths) or the open format. */
function inPlaceFormat(f: FileState): SaveFormat | null {
  if (typeof f.handle === 'string') {
    const e = ext(f.handle);
    if (e === 'xlsx' || e === 'xlsm') return 'xlsx';
    if (e === 'csv') return 'csv';
    if (e === 'tsv' || e === 'txt') return 'tsv';
    if (e === 'ods') return 'ods';
    return null;
  }
  if (f.format === 'xlsx' || f.format === 'csv') return f.format;
  return null;
}

const FILTER_FOR: Record<SaveFormat, FileTypeFilter> = { xlsx: FILTERS.xlsx, csv: FILTERS.csv, tsv: FILTERS.tsv, ods: FILTERS.ods, pdf: FILTERS.pdf };

/** Ctrl+S: write straight back to the opened file when possible, otherwise Save As. */
export async function saveFile(): Promise<boolean> {
  const st = S();
  if (st.edit) {
    const { commitEdit } = await import('./edit');
    if (!commitEdit('none')) return false;
  }
  const f = S().file;
  const fmt = f.handle ? inPlaceFormat(f) : null;
  if (!fmt) return saveAs(f.format === 'csv' ? 'csv' : f.format === 'ods' ? 'ods' : 'xlsx');
  setState({ file: { ...f, saving: true } });
  try {
    const data = await serialize(fmt);
    const res = await fileAdapter().save(data, f.handle, `${f.name}.${fmt}`);
    if (!res) {
      setState({ file: { ...S().file, saving: false } });
      return saveAs(fmt);
    }
    setState({ file: { ...S().file, dirty: false, saving: false, lastSaved: Date.now() } });
    if ((fmt === 'csv' || fmt === 'tsv') && S().wb.sheets.length > 1) setStatus('Only the active sheet was saved to CSV.');
    await clearAutosave();
    return true;
  } catch (e) {
    setState({ file: { ...S().file, saving: false } });
    alertBox(`The file couldn't be saved.\n\n${(e as Error).message}`, 'MyExcel', 'error');
    return false;
  }
}

export async function saveAs(format: SaveFormat = 'xlsx'): Promise<boolean> {
  const f = S().file;
  setState({ file: { ...f, saving: true } });
  try {
    if (format === 'pdf') {
      const { exportPdf } = await import('../print');
      await exportPdf();
      setState({ file: { ...S().file, saving: false } });
      return true;
    }
    const data = await serialize(format);
    const res = await fileAdapter().saveAs(data, `${f.name}.${format}`, [FILTER_FOR[format]]);
    if (!res) {
      setState({ file: { ...S().file, saving: false } });
      return false;
    }
    const name = res.name.replace(/\.[^.]+$/, '');
    const fmt = (format === 'tsv' ? 'csv' : format) as FileState['format'];
    setState({ file: { ...S().file, name, handle: res.handle, format: fmt, dirty: false, saving: false, lastSaved: Date.now() }, backstage: false });
    fileAdapter().addRecent({ name: res.name, path: typeof res.handle === 'string' ? res.handle : undefined, opened: Date.now() });
    await clearAutosave();
    return true;
  } catch (e) {
    setState({ file: { ...S().file, saving: false } });
    alertBox(`The file couldn't be saved.\n\n${(e as Error).message}`, 'MyExcel', 'error');
    return false;
  }
}

// ---------- crash recovery ----------

let timer: ReturnType<typeof setInterval> | null = null;
let lastRev = -1;

/** Snapshot the workbook every 30 s while there are unsaved changes (the adapter decides where). */
export function startAutosave(intervalMs = 30000): void {
  if (timer) return;
  timer = setInterval(async () => {
    const st = S();
    if (!st.file.dirty || st.wb.rev === lastRev) return;
    lastRev = st.wb.rev;
    try {
      await fileAdapter().writeRecovery({
        fileName: st.file.name,
        path: typeof st.file.handle === 'string' ? st.file.handle : undefined,
        savedAt: Date.now(),
        data: st.wb.toJSON(),
      });
      setState({ file: { ...S().file, autosaved: Date.now() } });
    } catch {
      /* quota, disk or serialisation failure – ignore */
    }
  }, intervalMs);
}

export async function checkRecovery(): Promise<void> {
  const rec = await fileAdapter().readRecovery();
  if (!rec) return;
  openDialog('recovery', { fileName: rec.fileName, savedAt: rec.savedAt });
}

export async function recoverAutosave(): Promise<void> {
  const rec = await fileAdapter().readRecovery();
  if (!rec) return;
  const wb = Workbook.fromJSON(rec.data as Parameters<typeof Workbook.fromJSON>[0]);
  loadWorkbook(wb, { name: `${rec.fileName} (Recovered)`, handle: rec.path, format: 'xlsx', dirty: true });
  // the recovered copy now belongs to this window: drop the old snapshot, the next autosave writes this window's own
  await fileAdapter().clearRecovery();
  lastRev = -1;
}

export async function clearAutosave(): Promise<void> {
  lastRev = -1;
  try {
    await fileAdapter().clearRecovery();
  } catch {
    /* ignore */
  }
}

// ---- Test build: Help > Reload keeps the open workbook ----
const RELOAD_KEY = 'myexcel.reloadRestore';

/** Snapshot the open workbook (saved or not) so it survives a UI reload. */
export async function snapshotForReload(): Promise<void> {
  const st = S();
  await fileAdapter().writeRecovery({
    fileName: st.file.name,
    path: typeof st.file.handle === 'string' ? st.file.handle : undefined,
    savedAt: Date.now(),
    data: st.wb.toJSON(),
  });
  sessionStorage.setItem(RELOAD_KEY, JSON.stringify({ dirty: st.file.dirty, format: st.file.format }));
}

/** After a reload started by snapshotForReload: reopen the workbook as it was. Returns true if it did. */
export async function restoreAfterReload(): Promise<boolean> {
  let meta: { dirty: boolean; format: FileState['format'] } | null = null;
  try {
    meta = JSON.parse(sessionStorage.getItem(RELOAD_KEY) || 'null');
    sessionStorage.removeItem(RELOAD_KEY);
  } catch {
    return false;
  }
  if (!meta) return false;
  const rec = await fileAdapter().readRecovery();
  if (!rec) return false;
  const wb = Workbook.fromJSON(rec.data as Parameters<typeof Workbook.fromJSON>[0]);
  loadWorkbook(wb, { name: rec.fileName, handle: rec.path, format: meta.format, dirty: meta.dirty });
  // A clean workbook needs no crash snapshot; a dirty one keeps it until saved, as usual.
  if (!meta.dirty) await clearAutosave();
  lastRev = -1;
  return true;
}

// File I/O abstraction so a desktop shell (Tauri) can replace the browser file picker.

export interface OpenedFile {
  name: string;
  data: ArrayBuffer;
  /** Opaque handle used for "Save" without prompting (FileSystemFileHandle or native path). */
  handle?: unknown;
}

export interface SavedFile {
  name: string;
  handle?: unknown;
}

export interface FileTypeFilter {
  description: string;
  extensions: string[]; // without dots
  mime: string;
}

export interface RecentFile {
  name: string;
  path?: string;
  opened: number;
}

/** User preferences persisted by the adapter (localStorage on the web, settings.json on desktop). */
export interface AppSettings {
  theme: 'light' | 'dark';
  /** Zoom (%) for new workbooks and sheets. */
  defaultZoom: number;
  /** Font for new workbooks. */
  defaultFont: string;
  recentFiles: RecentFile[];
}

export const DEFAULT_SETTINGS: AppSettings = { theme: 'light', defaultZoom: 100, defaultFont: 'Calibri', recentFiles: [] };

/** Crash-recovery snapshot of the open workbook. */
export interface RecoveryRecord {
  fileName: string;
  /** Path/handle of the file the snapshot belongs to, when it has one on disk. */
  path?: string;
  savedAt: number;
  data: unknown; // SerializedWorkbook
}

export interface FileAdapter {
  readonly kind: 'browser' | 'tauri';
  open(filters: FileTypeFilter[]): Promise<OpenedFile | null>;
  /** Save to an existing handle; returns null if the handle can't be written (caller falls back to saveAs). */
  save(data: Uint8Array | string, handle: unknown, name: string): Promise<SavedFile | null>;
  saveAs(data: Uint8Array | string, suggestedName: string, filters: FileTypeFilter[]): Promise<SavedFile | null>;
  recentFiles(): Promise<RecentFile[]>;
  openRecent?(entry: RecentFile): Promise<OpenedFile | null>;
  addRecent(entry: RecentFile): Promise<void>;
  loadSettings(): Promise<AppSettings>;
  saveSettings(s: AppSettings): Promise<void>;
  writeRecovery(rec: RecoveryRecord): Promise<void>;
  readRecovery(): Promise<RecoveryRecord | null>;
  clearRecovery(): Promise<void>;
}

export const FILTERS: Record<string, FileTypeFilter> = {
  xlsx: { description: 'Excel Workbook', extensions: ['xlsx'], mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
  xlsm: { description: 'Excel Macro-Enabled Workbook', extensions: ['xlsm'], mime: 'application/vnd.ms-excel.sheet.macroEnabled.12' },
  xls: { description: 'Excel 97-2003 Workbook', extensions: ['xls'], mime: 'application/vnd.ms-excel' },
  csv: { description: 'CSV (Comma delimited)', extensions: ['csv'], mime: 'text/csv' },
  tsv: { description: 'Text (Tab delimited)', extensions: ['tsv', 'txt'], mime: 'text/tab-separated-values' },
  ods: { description: 'OpenDocument Spreadsheet', extensions: ['ods'], mime: 'application/vnd.oasis.opendocument.spreadsheet' },
  pdf: { description: 'PDF', extensions: ['pdf'], mime: 'application/pdf' },
};

let current: FileAdapter | null = null;

export function setFileAdapter(a: FileAdapter): void {
  current = a;
}

export function fileAdapter(): FileAdapter {
  if (!current) throw new Error('FileAdapter not initialised');
  return current;
}

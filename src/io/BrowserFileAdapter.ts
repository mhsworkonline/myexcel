import { clearAutosave, readAutosave, writeAutosave } from './autosave';
import { AppSettings, DEFAULT_SETTINGS, FileAdapter, FileTypeFilter, OpenedFile, RecentFile, RecoveryRecord, SavedFile } from './FileAdapter';

interface FsHandle {
  name: string;
  getFile(): Promise<File>;
  createWritable(): Promise<{ write(d: unknown): Promise<void>; close(): Promise<void> }>;
  queryPermission?(o: { mode: string }): Promise<string>;
  requestPermission?(o: { mode: string }): Promise<string>;
}

type PickerWindow = Window & {
  showOpenFilePicker?: (o: unknown) => Promise<FsHandle[]>;
  showSaveFilePicker?: (o: unknown) => Promise<FsHandle>;
};

const SETTINGS_KEY = 'myexcel.settings';

function toAccept(filters: FileTypeFilter[]) {
  return filters.map((f) => ({ description: f.description, accept: { [f.mime]: f.extensions.map((e) => '.' + e) } }));
}

/** Browser implementation: File System Access API when available, <input type=file> / download otherwise. */
export class BrowserFileAdapter implements FileAdapter {
  readonly kind = 'browser' as const;

  async open(filters: FileTypeFilter[]): Promise<OpenedFile | null> {
    const w = window as PickerWindow;
    if (w.showOpenFilePicker) {
      try {
        const [h] = await w.showOpenFilePicker({ types: toAccept(filters), excludeAcceptAllOption: false, multiple: false });
        const f = await h.getFile();
        return { name: f.name, data: await f.arrayBuffer(), handle: h };
      } catch (e) {
        if ((e as Error).name === 'AbortError') return null;
      }
    }
    return new Promise((resolve) => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = filters.flatMap((f) => f.extensions.map((e) => '.' + e)).join(',');
      input.onchange = async () => {
        const f = input.files?.[0];
        if (!f) return resolve(null);
        resolve({ name: f.name, data: await f.arrayBuffer() });
      };
      input.addEventListener('cancel', () => resolve(null));
      input.click();
    });
  }

  async save(data: Uint8Array | string, handle: unknown, name: string): Promise<SavedFile | null> {
    const h = handle as FsHandle | undefined;
    if (!h || typeof h.createWritable !== 'function') return null;
    try {
      if (h.queryPermission && (await h.queryPermission({ mode: 'readwrite' })) !== 'granted') {
        if (!h.requestPermission || (await h.requestPermission({ mode: 'readwrite' })) !== 'granted') return null;
      }
      const wr = await h.createWritable();
      await wr.write(data);
      await wr.close();
      return { name: h.name ?? name, handle: h };
    } catch {
      return null;
    }
  }

  async saveAs(data: Uint8Array | string, suggestedName: string, filters: FileTypeFilter[]): Promise<SavedFile | null> {
    const w = window as PickerWindow;
    if (w.showSaveFilePicker) {
      try {
        const h = await w.showSaveFilePicker({ suggestedName, types: toAccept(filters) });
        const wr = await h.createWritable();
        await wr.write(data);
        await wr.close();
        return { name: h.name, handle: h };
      } catch (e) {
        if ((e as Error).name === 'AbortError') return null;
      }
    }
    const blob = new Blob([data as BlobPart], { type: filters[0]?.mime ?? 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = suggestedName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
    return { name: suggestedName };
  }

  async recentFiles(): Promise<RecentFile[]> {
    return (await this.loadSettings()).recentFiles;
  }

  async addRecent(entry: RecentFile): Promise<void> {
    const s = await this.loadSettings();
    s.recentFiles = [entry, ...s.recentFiles.filter((e) => e.name !== entry.name)].slice(0, 20);
    await this.saveSettings(s);
  }

  async loadSettings(): Promise<AppSettings> {
    try {
      const raw = JSON.parse(localStorage.getItem(SETTINGS_KEY) ?? '{}');
      // migrate the pre-settings keys
      const legacyTheme = localStorage.getItem('myexcel.theme');
      const legacyRecent = localStorage.getItem('myexcel.recent');
      return {
        ...DEFAULT_SETTINGS,
        ...(legacyTheme ? { theme: legacyTheme } : {}),
        ...(legacyRecent ? { recentFiles: JSON.parse(legacyRecent) } : {}),
        ...raw,
      };
    } catch {
      return { ...DEFAULT_SETTINGS };
    }
  }

  async saveSettings(s: AppSettings): Promise<void> {
    try {
      localStorage.setItem(SETTINGS_KEY, JSON.stringify(s));
    } catch {
      /* storage unavailable */
    }
  }

  async writeRecovery(rec: RecoveryRecord): Promise<void> {
    await writeAutosave(rec as Parameters<typeof writeAutosave>[0]);
  }

  async readRecovery(): Promise<RecoveryRecord | null> {
    return readAutosave();
  }

  async clearRecovery(): Promise<void> {
    await clearAutosave();
  }
}

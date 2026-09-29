import type { FileAdapter, FileTypeFilter, OpenedFile, RecentFile, SavedFile } from './FileAdapter';

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

const RECENT_KEY = 'myexcel.recent';

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
    try {
      return JSON.parse(localStorage.getItem(RECENT_KEY) ?? '[]');
    } catch {
      return [];
    }
  }

  async addRecent(entry: RecentFile): Promise<void> {
    try {
      const list = (await this.recentFiles()).filter((e) => e.name !== entry.name);
      list.unshift(entry);
      localStorage.setItem(RECENT_KEY, JSON.stringify(list.slice(0, 20)));
    } catch {
      /* storage unavailable */
    }
  }
}

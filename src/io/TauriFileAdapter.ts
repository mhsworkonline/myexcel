// Desktop file adapter: native open/save dialogs and direct filesystem access via Tauri v2.
import type { FileAdapter, FileTypeFilter, OpenedFile, RecentFile, SavedFile } from './FileAdapter';

const RECENT_KEY = 'myexcel.recent.desktop';

function baseName(p: string): string {
  return p.split(/[\\/]/).pop() ?? p;
}

function toFilters(filters: FileTypeFilter[]) {
  return filters.map((f) => ({ name: f.description, extensions: f.extensions }));
}

export class TauriFileAdapter implements FileAdapter {
  readonly kind = 'tauri' as const;

  async readPath(path: string): Promise<OpenedFile> {
    const { readFile } = await import('@tauri-apps/plugin-fs');
    const bytes = await readFile(path);
    return { name: baseName(path), data: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer, handle: path };
  }

  async open(filters: FileTypeFilter[]): Promise<OpenedFile | null> {
    const { open } = await import('@tauri-apps/plugin-dialog');
    const path = await open({ multiple: false, directory: false, filters: toFilters(filters) });
    if (!path || Array.isArray(path)) return null;
    return this.readPath(path);
  }

  private async write(path: string, data: Uint8Array | string): Promise<void> {
    const fs = await import('@tauri-apps/plugin-fs');
    if (typeof data === 'string') await fs.writeTextFile(path, data);
    else await fs.writeFile(path, data);
  }

  async save(data: Uint8Array | string, handle: unknown, name: string): Promise<SavedFile | null> {
    if (typeof handle !== 'string') return null;
    try {
      await this.write(handle, data);
      return { name: baseName(handle) || name, handle };
    } catch {
      return null;
    }
  }

  async saveAs(data: Uint8Array | string, suggestedName: string, filters: FileTypeFilter[]): Promise<SavedFile | null> {
    const { save } = await import('@tauri-apps/plugin-dialog');
    const path = await save({ defaultPath: suggestedName, filters: toFilters(filters) });
    if (!path) return null;
    await this.write(path, data);
    return { name: baseName(path), handle: path };
  }

  async recentFiles(): Promise<RecentFile[]> {
    try {
      return JSON.parse(localStorage.getItem(RECENT_KEY) ?? '[]');
    } catch {
      return [];
    }
  }

  async openRecent(entry: RecentFile): Promise<OpenedFile | null> {
    if (!entry.path) return null;
    try {
      return await this.readPath(entry.path);
    } catch {
      return null;
    }
  }

  async addRecent(entry: RecentFile): Promise<void> {
    try {
      const list = (await this.recentFiles()).filter((e) => (e.path ?? e.name) !== (entry.path ?? entry.name));
      list.unshift(entry);
      localStorage.setItem(RECENT_KEY, JSON.stringify(list.slice(0, 20)));
    } catch {
      /* ignore */
    }
  }
}

/** Wire native menus and the startup file (file association) to the app. */
export async function installDesktopBridge(handlers: Record<string, () => void>, openPath: (p: string) => Promise<void>): Promise<void> {
  const { listen } = await import('@tauri-apps/api/event');
  const { invoke } = await import('@tauri-apps/api/core');
  await listen<string>('menu', (e) => handlers[e.payload]?.());
  const startup = await invoke<string | null>('startup_file');
  if (startup) await openPath(startup);
}

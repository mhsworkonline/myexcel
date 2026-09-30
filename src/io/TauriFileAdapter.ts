// Desktop file adapter (Tauri v2): native dialogs, direct filesystem access, and local
// app-data storage for settings, recent files and crash-recovery snapshots.
import type { AppSettings, FileAdapter, FileTypeFilter, OpenedFile, RecentFile, RecoveryRecord, SavedFile } from './FileAdapter';
import { DEFAULT_SETTINGS } from './FileAdapter';

function baseName(p: string): string {
  return p.split(/[\\/]/).pop() ?? p;
}

function toFilters(filters: FileTypeFilter[]) {
  return filters.map((f) => ({ name: f.description, extensions: f.extensions }));
}

async function fs() {
  return import('@tauri-apps/plugin-fs');
}

async function paths() {
  return import('@tauri-apps/api/path');
}

/** Write via a temp file + rename so a crash mid-write never leaves a truncated file. */
async function atomicWriteText(path: string, text: string): Promise<void> {
  const f = await fs();
  const tmp = path + '.tmp';
  await f.writeTextFile(tmp, text);
  try {
    if (await f.exists(path)) await f.remove(path);
  } catch {
    /* ignore */
  }
  await f.rename(tmp, path);
}

export class TauriFileAdapter implements FileAdapter {
  readonly kind = 'tauri' as const;
  private settingsCache: AppSettings | null = null;

  private async configFile(): Promise<string> {
    const p = await paths();
    const dir = await p.appConfigDir();
    const f = await fs();
    if (!(await f.exists(dir))) await f.mkdir(dir, { recursive: true });
    return p.join(dir, 'settings.json');
  }

  /** This window's snapshot (one per window, so several open windows don't overwrite each other). */
  private async recoveryFile(): Promise<string> {
    const p = await paths();
    const dir = await p.join(await p.appDataDir(), 'recovery');
    const f = await fs();
    if (!(await f.exists(dir))) await f.mkdir(dir, { recursive: true });
    const { invoke } = await import('@tauri-apps/api/core');
    return p.join(dir, `snapshot-${await invoke<number>('instance_id')}.json`);
  }

  /** A crashed window's snapshot offered by readRecovery; removed by clearRecovery. */
  private claimed: string | null = null;

  async readPath(path: string): Promise<OpenedFile> {
    const f = await fs();
    const bytes = await f.readFile(path);
    return { name: baseName(path), data: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer, handle: path };
  }

  async open(filters: FileTypeFilter[]): Promise<OpenedFile | null> {
    const { open } = await import('@tauri-apps/plugin-dialog');
    const path = await open({ multiple: false, directory: false, filters: toFilters(filters) });
    if (!path || Array.isArray(path)) return null;
    return this.readPath(path);
  }

  private async write(path: string, data: Uint8Array | string): Promise<void> {
    const f = await fs();
    if (typeof data === 'string') await f.writeTextFile(path, data);
    else await f.writeFile(path, data);
  }

  /** Ctrl+S: write straight back to the file the workbook was opened from / last saved to. */
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
    return (await this.loadSettings()).recentFiles;
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
    const s = await this.loadSettings();
    const key = (e: RecentFile) => (e.path ?? e.name).toLowerCase();
    s.recentFiles = [entry, ...s.recentFiles.filter((e) => key(e) !== key(entry))].slice(0, 20);
    await this.saveSettings(s);
  }

  async loadSettings(): Promise<AppSettings> {
    if (this.settingsCache) return { ...this.settingsCache, recentFiles: [...this.settingsCache.recentFiles] };
    let s: AppSettings = { ...DEFAULT_SETTINGS };
    try {
      const f = await fs();
      const file = await this.configFile();
      if (await f.exists(file)) s = { ...DEFAULT_SETTINGS, ...JSON.parse(await f.readTextFile(file)) };
    } catch {
      /* corrupt or unreadable settings: fall back to defaults */
    }
    this.settingsCache = s;
    return { ...s, recentFiles: [...s.recentFiles] };
  }

  async saveSettings(s: AppSettings): Promise<void> {
    this.settingsCache = { ...s, recentFiles: [...s.recentFiles] };
    await atomicWriteText(await this.configFile(), JSON.stringify(s, null, 2));
  }

  async writeRecovery(rec: RecoveryRecord): Promise<void> {
    await atomicWriteText(await this.recoveryFile(), JSON.stringify(rec));
  }

  /** This window's own snapshot (after Help > Reload), else the newest one left by a window that crashed. */
  async readRecovery(): Promise<RecoveryRecord | null> {
    try {
      const f = await fs();
      const own = await this.recoveryFile();
      if (await f.exists(own)) return JSON.parse(await f.readTextFile(own)) as RecoveryRecord;
      const { invoke } = await import('@tauri-apps/api/core');
      for (const file of await invoke<string[]>('recovery_candidates')) {
        try {
          const rec = JSON.parse(await f.readTextFile(file)) as RecoveryRecord;
          this.claimed = file;
          return rec;
        } catch {
          await f.remove(file).catch(() => undefined); // unreadable leftover
        }
      }
      return null;
    } catch {
      return null;
    }
  }

  async clearRecovery(): Promise<void> {
    const f = await fs();
    for (const file of [await this.recoveryFile().catch(() => null), this.claimed]) {
      if (!file) continue;
      try {
        if (await f.exists(file)) await f.remove(file);
      } catch {
        /* ignore */
      }
    }
    this.claimed = null;
  }
}

export interface DesktopBridgeOptions {
  handlers: Record<string, () => void>;
  openPath: (p: string) => Promise<void>;
  /** Called on window close when there are unsaved changes; resolves true when it is OK to close. */
  confirmClose: () => Promise<boolean>;
  /** Called when the window really closes (clean up recovery etc.). */
  beforeExit: () => Promise<void>;
  subscribeTitle: (set: (title: string) => void) => void;
}

/** Wire native menus, the launch file (file association), close prompt and window title. */
export async function installDesktopBridge(o: DesktopBridgeOptions): Promise<void> {
  const { listen } = await import('@tauri-apps/api/event');
  const { invoke } = await import('@tauri-apps/api/core');
  const { getCurrentWindow } = await import('@tauri-apps/api/window');
  const win = getCurrentWindow();
  await listen<string>('menu', (e) => o.handlers[e.payload]?.());
  let closing = false;
  await win.onCloseRequested(async (event) => {
    event.preventDefault();
    if (closing) return;
    closing = true;
    try {
      if (await o.confirmClose()) {
        await o.beforeExit();
        await win.destroy();
      }
    } finally {
      closing = false;
    }
  });
  let lastTitle = '';
  o.subscribeTitle((t) => {
    if (t !== lastTitle) {
      lastTitle = t;
      win.setTitle(t).catch(() => undefined);
    }
  });
  const startup = await invoke<string | null>('startup_file');
  if (startup) await o.openPath(startup);
}

export async function nativeClipboardRead(): Promise<{ html: string; text: string }> {
  const { invoke } = await import('@tauri-apps/api/core');
  const r = await invoke<{ html?: string | null; text?: string | null }>('clipboard_read');
  return { html: r.html ?? '', text: r.text ?? '' };
}

export async function nativeClipboardWrite(html: string, text: string): Promise<void> {
  const { invoke } = await import('@tauri-apps/api/core');
  await invoke('clipboard_write', { html, text });
}

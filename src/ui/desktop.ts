// Desktop (Tauri) integration: native menu commands, launch file, close prompt,
// window title and native clipboard. Only loaded when running inside Tauri.
import { setNativeClipboard, copyToSystem, pasteFromSystem } from '../state/actions/clipboard';
import * as file from '../state/actions/file';
import { freezePanes } from '../state/actions/structure';
import { clearAutosave } from '../state/actions/file';
import { bump, openDialog, redo, S, setState, transact, undo } from '../state/store';
import { setThemeSetting } from '../state/settings';
import type { TauriFileAdapter } from '../io/TauriFileAdapter';

export function isTauri(): boolean {
  const w = window as unknown as { __TAURI_INTERNALS__?: unknown };
  return !!w.__TAURI_INTERNALS__;
}

function zoomBy(d: number): void {
  const sh = S().wb.activeSheet;
  sh.zoom = d === 0 ? 100 : Math.max(10, Math.min(400, sh.zoom + d));
  sh.touch();
  bump();
}

/** Menu ids (src-tauri/src/lib.rs) → the same actions the ribbon and shortcuts use. */
export const MENU_COMMANDS: Record<string, () => void> = {
  new: () => file.newWorkbook(),
  close: () => file.closeWorkbook(),
  open: () => file.openFile(),
  recent: () => setState({ backstage: true, backstagePage: 'open' }),
  save: () => void file.saveFile(),
  saveAs: () => openDialog('saveAs'),
  exportPdf: () => void file.saveAs('pdf'),
  pageSetup: () => openDialog('pageSetup'),
  print: () => openDialog('print'),
  options: () => setState({ backstage: true, backstagePage: 'options' }),
  exit: () => import('@tauri-apps/api/window').then((m) => m.getCurrentWindow().close()),
  undo,
  redo,
  cut: () => void copyToSystem(true),
  copy: () => void copyToSystem(false),
  paste: () => void pasteFromSystem(),
  pasteSpecial: () => openDialog('pasteSpecial'),
  find: () => openDialog('findReplace', { tab: 'find' }),
  replace: () => openDialog('findReplace', { tab: 'replace' }),
  goto: () => openDialog('goto'),
  selectAll: () => import('../state/actions/find').then((m) => m.selectAllCells()),
  toggleTheme: () => void setThemeSetting(S().theme === 'dark' ? 'light' : 'dark'),
  toggleFormulaBar: () => {
    if (document.body.dataset.hideFormulaBar) delete document.body.dataset.hideFormulaBar;
    else document.body.dataset.hideFormulaBar = '1';
    bump();
  },
  toggleGridlines: () => {
    const sh = S().wb.activeSheet;
    transact('Gridlines', (tx) => tx.setMeta(sh, 'showGridlines', !sh.showGridlines));
  },
  zoomIn: () => zoomBy(10),
  zoomOut: () => zoomBy(-10),
  zoomReset: () => zoomBy(0),
  freeze: () => {
    const sh = S().wb.activeSheet;
    freezePanes(sh.freeze.rows || sh.freeze.cols ? 'unfreeze' : 'panes');
  },
  shortcuts: () => openDialog('shortcuts'),
  // Test build only (the menu item exists only there): pick up a freshly deployed UI.
  reload: async () => {
    await file.snapshotForReload();
    location.reload();
  },
  about: () => setState({ backstage: true, backstagePage: 'about' }),
};

/** "test" in the side-by-side test build, whose title also carries the UI build number. */
let channel: 'release' | 'test' = 'release';

function appTitle(): string {
  if (channel !== 'test') return 'MyExcel';
  const n = process.env.NEXT_PUBLIC_TEST_BUILD;
  return `MyExcel (Test${n ? ' build ' + n : ''})`;
}

function windowTitle(): string {
  const f = S().file;
  const status = f.saving ? 'Saving…' : f.dirty ? 'Not saved' : f.lastSaved ? 'Saved' : '';
  return `${f.name}${status ? ' • ' + status : ''} – ${appTitle()}`;
}

/** Native "save changes?" prompt; resolves true when the window may close. */
async function confirmClose(): Promise<boolean> {
  const st = S();
  if (!st.file.dirty) return true;
  const { message } = await import('@tauri-apps/plugin-dialog');
  const res = await message(`Want to save your changes to '${st.file.name}'?`, {
    title: 'MyExcel',
    kind: 'warning',
    buttons: { yes: 'Save', no: "Don't Save", cancel: 'Cancel' },
  });
  if (res === 'Save' || res === 'Yes') return file.saveFile();
  if (res === "Don't Save" || res === 'No') return true;
  return false;
}

export async function installDesktop(adapter: TauriFileAdapter): Promise<void> {
  const m = await import('../io/TauriFileAdapter');
  const { invoke } = await import('@tauri-apps/api/core');
  channel = await invoke<'release' | 'test'>('build_channel').catch(() => 'release' as const);
  file.setWindowOps({
    open: (path) => invoke('open_window', { path: path ?? null }),
    close: () => import('@tauri-apps/api/window').then((w) => w.getCurrentWindow().close()),
  });
  setNativeClipboard({ read: m.nativeClipboardRead, write: m.nativeClipboardWrite });
  const { useStore } = await import('../state/store');
  await m.installDesktopBridge({
    handlers: MENU_COMMANDS,
    openPath: async (p) => {
      try {
        const f = await adapter.readPath(p);
        await file.openFromData(f.name, f.data, f.handle);
      } catch (e) {
        openDialog('alert', { message: `MyExcel cannot open '${p}'.\n\n${(e as Error).message}`, icon: 'error' });
      }
    },
    confirmClose,
    // A deliberate close (saved or "Don't Save") needs no recovery snapshot.
    beforeExit: () => clearAutosave(),
    subscribeTitle: (set) => {
      set(windowTitle());
      useStore.subscribe((s, prev) => {
        if (s.file !== prev.file) set(windowTitle());
      });
    },
  });
}

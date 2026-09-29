'use client';
import { useEffect } from 'react';
import { BrowserFileAdapter } from '../io/BrowserFileAdapter';
import { setFileAdapter } from '../io/FileAdapter';
import { checkRecovery, openFromData, startAutosave } from '../state/actions/file';
import { S, setState, useStore } from '../state/store';
import { Backstage } from './Backstage';
import { ContextMenus } from './ContextMenus';
import { DialogHost } from './dialogs/DialogHost';
import { FormulaBar } from './FormulaBar';
import { Grid } from './grid/Grid';
import { Ribbon } from './ribbon/Ribbon';
import { SheetTabs } from './SheetTabs';
import { StatusBar } from './StatusBar';
import { TitleBar } from './TitleBar';
import { installPhase3 } from './phase3';

let booted = false;

function boot(): void {
  if (booted) return;
  booted = true;
  const w = window as unknown as { __TAURI__?: unknown; __TAURI_INTERNALS__?: unknown };
  if (w.__TAURI__ || w.__TAURI_INTERNALS__) {
    import('../io/TauriFileAdapter').then((m) => setFileAdapter(new m.TauriFileAdapter())).catch(() => setFileAdapter(new BrowserFileAdapter()));
  } else setFileAdapter(new BrowserFileAdapter());
  let theme: 'light' | 'dark' = 'light';
  try {
    const t = localStorage.getItem('myexcel.theme');
    if (t === 'dark' || t === 'light') theme = t;
  } catch {
    /* ignore */
  }
  document.documentElement.dataset.theme = theme;
  setState({ theme });
  installPhase3();
  startAutosave();
  checkRecovery();
  // expose a tiny test hook for automation
  (window as unknown as { __myexcel?: unknown }).__myexcel = { S, openFromData };
}

export function App() {
  useEffect(() => {
    boot();
    const beforeUnload = (e: BeforeUnloadEvent) => {
      if (S().file.dirty) {
        e.preventDefault();
        e.returnValue = '';
      }
    };
    window.addEventListener('beforeunload', beforeUnload);
    // file drop to open
    const onDrop = async (e: DragEvent) => {
      const f = e.dataTransfer?.files?.[0];
      if (!f) return;
      e.preventDefault();
      await openFromData(f.name, await f.arrayBuffer());
    };
    const onDragOver = (e: DragEvent) => {
      if (e.dataTransfer?.types.includes('Files')) e.preventDefault();
    };
    window.addEventListener('drop', onDrop);
    window.addEventListener('dragover', onDragOver);
    // global shortcuts that must work even when focus is in chrome
    const onKey = (e: KeyboardEvent) => {
      const st = S();
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's' && !st.dialog) {
        e.preventDefault();
        import('../state/actions/file').then((m) => m.saveFile());
      }
      if (e.altKey && e.key.toLowerCase() === 'f' && !st.edit) {
        e.preventDefault();
        setState({ backstage: true });
      }
      if (e.key === 'Escape' && st.backstage) setState({ backstage: false });
    };
    window.addEventListener('keydown', onKey);
    return () => {
      window.removeEventListener('beforeunload', beforeUnload);
      window.removeEventListener('drop', onDrop);
      window.removeEventListener('dragover', onDragOver);
      window.removeEventListener('keydown', onKey);
    };
  }, []);
  const theme = useStore((s) => s.theme);
  return (
    <div className="flex flex-col h-screen w-screen overflow-hidden" data-theme={theme}>
      <TitleBar />
      <Ribbon />
      <FormulaBar />
      <div className="flex-1 flex min-h-0" style={{ borderTop: '1px solid var(--border)' }}>
        <Grid />
      </div>
      <SheetTabs />
      <StatusBar />
      <ContextMenus />
      <DialogHost />
      <Backstage />
    </div>
  );
}

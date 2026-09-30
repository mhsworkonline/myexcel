// Application settings (theme, default zoom, default font, recent files), persisted by the FileAdapter:
// settings.json in the app config directory on desktop, localStorage on the web.
import { AppSettings, DEFAULT_SETTINGS, fileAdapter } from '../io/FileAdapter';
import { bump, S, setState } from './store';

let current: AppSettings = { ...DEFAULT_SETTINGS };

export function settings(): AppSettings {
  return current;
}

export async function loadSettings(): Promise<AppSettings> {
  current = await fileAdapter().loadSettings();
  current.defaultZoom = Math.max(10, Math.min(400, Number(current.defaultZoom) || 100));
  return current;
}

export async function updateSettings(patch: Partial<AppSettings>): Promise<void> {
  // recent files are owned by the adapter; re-read so we don't clobber them
  const latest = await fileAdapter().loadSettings();
  current = { ...latest, ...patch };
  await fileAdapter().saveSettings(current);
}

export function applyTheme(t: 'light' | 'dark'): void {
  document.documentElement.dataset.theme = t;
  setState({ theme: t });
  bump();
}

export async function setThemeSetting(t: 'light' | 'dark'): Promise<void> {
  applyTheme(t);
  await updateSettings({ theme: t });
}

/** Apply defaults to the untouched start-up workbook (it was created before settings loaded). */
export function applyDefaultsToPristineWorkbook(): void {
  const st = S();
  if (st.file.dirty || st.history.canUndo || st.file.lastSaved) return;
  for (const s of st.wb.sheets) {
    s.zoom = current.defaultZoom;
    s.touch();
  }
  if (current.defaultFont && current.defaultFont !== 'Calibri') st.wb.props.defaultFont = current.defaultFont;
  import('../model/styles').then((m) => {
    m.setDocDefaultFont(st.wb.props.defaultFont);
    bump();
  });
}

'use client';
import { ChevronDown, Moon, Redo2, Save, Search, Sun, Undo2 } from 'lucide-react';
import { saveFile } from '../state/actions/file';
import { openDialog, redo, S, undo, useStore } from '../state/store';
import { AppIcon } from './icons';
import { Menu } from './Menu';
import { useState } from 'react';
import { setTheme } from './ribbon/OtherTabs';

function timeAgo(t?: number): string {
  if (!t) return '';
  const s = Math.round((Date.now() - t) / 1000);
  if (s < 60) return 'just now';
  const m = Math.round(s / 60);
  return `${m} minute${m === 1 ? '' : 's'} ago`;
}

export function TitleBar() {
  const file = useStore((s) => s.file);
  const theme = useStore((s) => s.theme);
  useStore((s) => s.rev);
  const h = S().history;
  const [menu, setMenu] = useState<null | { x: number; y: number; kind: 'undo' | 'redo' | 'title' }>(null);
  const status = file.saving ? 'Saving…' : file.dirty ? (file.autosaved ? 'Autosaved locally' : 'Not saved') : file.lastSaved ? 'Saved' : '';
  const fmt = file.format === 'xlsx' ? '' : ` [${file.format.toUpperCase()}]`;
  return (
    <div className="flex items-center h-[36px] px-2 gap-1 select-none" style={{ background: 'var(--title-bg)' }} data-testid="titlebar">
      <div className="flex items-center gap-1.5 pr-2">
        <AppIcon />
      </div>
      {/* Quick Access Toolbar */}
      <div className="flex items-center gap-[1px]" role="toolbar" aria-label="Quick Access Toolbar">
        <span className="text-[12px] mr-1 opacity-80 hidden md:inline">AutoSave</span>
        <button role="switch" aria-checked={true} className="relative w-[34px] h-[17px] rounded-full mr-2 hidden md:inline-block" style={{ background: 'var(--accent)' }} title="AutoSave to local recovery is on">
          <span className="absolute right-[3px] top-[3px] w-[11px] h-[11px] rounded-full bg-white" />
        </button>
        <button className="xl-btn-sm !h-[26px] !px-1" title="Save (Ctrl+S)" onClick={() => saveFile()} data-testid="qat-save">
          <Save size={16} />
        </button>
        <div className="flex items-center">
          <button className="xl-btn-sm !h-[26px] !px-1" title={h.canUndo ? `Undo ${h.undoStack[h.undoStack.length - 1].cmd.label} (Ctrl+Z)` : "Can't Undo"} disabled={!h.canUndo} onClick={undo} data-testid="qat-undo">
            <Undo2 size={16} className={h.canUndo ? '' : 'opacity-40'} />
          </button>
          <button className="xl-btn-sm !h-[26px] !px-0" disabled={!h.canUndo} onClick={(e) => setMenu({ x: e.currentTarget.getBoundingClientRect().left, y: 34, kind: 'undo' })} aria-label="Undo list">
            <ChevronDown size={11} className={h.canUndo ? '' : 'opacity-40'} />
          </button>
        </div>
        <div className="flex items-center">
          <button className="xl-btn-sm !h-[26px] !px-1" title={h.canRedo ? `Redo ${h.redoStack[h.redoStack.length - 1].cmd.label} (Ctrl+Y)` : "Can't Redo"} disabled={!h.canRedo} onClick={redo} data-testid="qat-redo">
            <Redo2 size={16} className={h.canRedo ? '' : 'opacity-40'} />
          </button>
          <button className="xl-btn-sm !h-[26px] !px-0" disabled={!h.canRedo} onClick={(e) => setMenu({ x: e.currentTarget.getBoundingClientRect().left, y: 34, kind: 'redo' })} aria-label="Redo list">
            <ChevronDown size={11} className={h.canRedo ? '' : 'opacity-40'} />
          </button>
        </div>
      </div>
      <div className="flex-1 flex justify-center min-w-0">
        <button className="flex items-center gap-1.5 px-2 h-[26px] rounded hover:bg-[var(--hover)] min-w-0" onClick={(e) => setMenu({ x: e.currentTarget.getBoundingClientRect().left, y: 34, kind: 'title' })} data-testid="doc-title">
          <span className="font-semibold truncate text-[12.5px]">{file.name}{fmt}</span>
          {status && <span className="opacity-70 text-[12px] whitespace-nowrap">• {status}</span>}
          <ChevronDown size={12} className="opacity-70" />
        </button>
      </div>
      <div className="hidden lg:flex items-center h-[26px] w-[300px] rounded-[4px] px-2 gap-2 mr-4 cursor-text" style={{ background: 'var(--panel)', border: '1px solid var(--border)' }} onClick={() => openDialog('findReplace', { tab: 'find' })}>
        <Search size={14} className="opacity-60" />
        <span className="opacity-60 text-[12px]">Search (Alt+Q)</span>
      </div>
      <button className="xl-btn-sm !h-[26px] !px-1.5" title={theme === 'dark' ? 'Switch to light theme' : 'Switch to dark theme'} onClick={() => setTheme(theme === 'dark' ? 'light' : 'dark')} data-testid="theme-toggle">
        {theme === 'dark' ? <Sun size={16} /> : <Moon size={16} />}
      </button>
      <div className="w-[26px] h-[26px] rounded-full flex items-center justify-center text-[11px] font-semibold text-white ml-1" style={{ background: '#8764B8' }} title="Local user">
        ME
      </div>
      {menu && menu.kind !== 'title' && (
        <Menu
          x={menu.x}
          y={menu.y}
          onClose={() => setMenu(null)}
          items={(menu.kind === 'undo' ? [...h.undoStack].reverse() : [...h.redoStack].reverse()).slice(0, 20).map((e, i) => ({
            label: e.cmd.label,
            onClick: () => {
              for (let k = 0; k <= i; k++) (menu.kind === 'undo' ? undo : redo)();
            },
          }))}
        />
      )}
      {menu?.kind === 'title' && (
        <Menu
          x={menu.x}
          y={menu.y}
          onClose={() => setMenu(null)}
          minWidth={260}
          items={[
            { label: `File name: ${file.name}`, disabled: true },
            { label: `Status: ${status || 'New'}${file.lastSaved ? ' ' + timeAgo(file.lastSaved) : ''}`, disabled: true },
            { separator: true },
            { label: 'Save', onClick: () => saveFile(), shortcut: 'Ctrl+S' },
            { label: 'Save As...', onClick: () => openDialog('saveAs'), shortcut: 'F12' },
          ]}
        />
      )}
    </div>
  );
}

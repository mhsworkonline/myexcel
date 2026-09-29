'use client';
import { ChevronLeft, ChevronRight, Plus } from 'lucide-react';
import { useEffect, useRef, useState } from 'react';
import { unprotectSheet } from '../state/actions/data';
import { activateSheet, addSheet, copySheet, deleteSheet, hideSheet, moveSheet, renameSheet, setTabColor } from '../state/actions/structure';
import { openDialog, S, setState, useStore } from '../state/store';
import { gridApi, onScrollChange } from './grid/scrollBus';
import { Menu } from './Menu';
import { ColorPalette } from './ribbon/parts';

export function SheetTabs() {
  useStore((s) => s.rev);
  const wb = useStore((s) => s.wb);
  const menu = useStore((s) => s.menu);
  const [renaming, setRenaming] = useState<string | null>(null);
  const [drag, setDrag] = useState<{ id: string; over: number; copy: boolean } | null>(null);
  const [tabsW, setTabsW] = useState(() => {
    try {
      return +(localStorage.getItem('myexcel.tabsW') ?? 0) || 0;
    } catch {
      return 0;
    }
  });
  const stripRef = useRef<HTMLDivElement>(null);
  const rootRef = useRef<HTMLDivElement>(null);
  const visible = wb.sheets.filter((s) => s.visibility === 'visible');
  const active = wb.activeSheetId;

  useEffect(() => {
    const onRename = (e: Event) => setRenaming((e as CustomEvent).detail as string);
    window.addEventListener('xl-rename-sheet', onRename);
    return () => window.removeEventListener('xl-rename-sheet', onRename);
  }, []);

  useEffect(() => {
    // keep active tab in view
    const el = stripRef.current?.querySelector(`[data-sheet-id="${active}"]`) as HTMLElement | null;
    el?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }, [active]);

  const scrollTabs = (dir: number, far: boolean) => {
    const s = stripRef.current;
    if (!s) return;
    if (far) s.scrollLeft = dir < 0 ? 0 : s.scrollWidth;
    else s.scrollBy({ left: dir * 120 });
  };

  const width = tabsW || Math.round((rootRef.current?.getBoundingClientRect().width ?? 1200) * 0.55);

  return (
    <div ref={rootRef} className="flex items-stretch h-[30px] select-none" style={{ background: 'var(--chrome-bg)', borderTop: '1px solid var(--border)' }} data-testid="sheet-tabs">
      <div className="flex items-stretch min-w-0" style={{ width }}>
        <div className="flex items-center px-1 gap-0.5">
          <button className="w-5 h-5 flex items-center justify-center rounded hover:bg-[var(--hover)]" title="Scroll to the previous sheet (Ctrl+click for first)" onClick={(e) => scrollTabs(-1, e.ctrlKey)}>
            <ChevronLeft size={14} />
          </button>
          <button className="w-5 h-5 flex items-center justify-center rounded hover:bg-[var(--hover)]" title="Scroll to the next sheet (Ctrl+click for last)" onClick={(e) => scrollTabs(1, e.ctrlKey)}>
            <ChevronRight size={14} />
          </button>
        </div>
        <div ref={stripRef} className="flex items-stretch overflow-hidden min-w-0" onWheel={(e) => stripRef.current && (stripRef.current.scrollLeft += e.deltaY)}>
          {visible.map((s, i) => {
            const isActive = s.id === active;
            return (
              <div
                key={s.id}
                data-sheet-id={s.id}
                data-testid={`sheet-tab-${s.name}`}
                className={'xl-sheet-tab ' + (isActive ? 'xl-sheet-tab-active' : '')}
                style={{
                  borderLeft: drag && drag.over === i ? '2px solid var(--text)' : undefined,
                  ...(s.tabColor && !isActive ? { boxShadow: `inset 0 -4px 0 ${s.tabColor}` } : {}),
                  ...(s.tabColor && isActive ? { background: `linear-gradient(to bottom, var(--tab-active-bg) 70%, ${s.tabColor}44)` } : {}),
                }}
                onPointerDown={(e) => {
                  if (e.button !== 0) return;
                  if (renaming) return;
                  activateSheet(s.id);
                  const startX = e.clientX;
                  let moved = false;
                  const onMove = (ev: PointerEvent) => {
                    if (Math.abs(ev.clientX - startX) < 6 && !moved) return;
                    moved = true;
                    const tabs = [...(stripRef.current?.querySelectorAll('[data-sheet-id]') ?? [])] as HTMLElement[];
                    let over = tabs.length;
                    for (let k = 0; k < tabs.length; k++) {
                      const r = tabs[k].getBoundingClientRect();
                      if (ev.clientX < r.left + r.width / 2) {
                        over = k;
                        break;
                      }
                    }
                    setDrag({ id: s.id, over, copy: ev.ctrlKey });
                  };
                  const onUp = (ev: PointerEvent) => {
                    window.removeEventListener('pointermove', onMove);
                    window.removeEventListener('pointerup', onUp);
                    setDrag((d) => {
                      if (d && moved) {
                        const target = d.over < visible.length ? wb.sheets.indexOf(visible[d.over]) : wb.sheets.length;
                        if (ev.ctrlKey) copySheet(d.id, target);
                        else moveSheet(d.id, target);
                      }
                      return null;
                    });
                  };
                  window.addEventListener('pointermove', onMove);
                  window.addEventListener('pointerup', onUp);
                }}
                onDoubleClick={() => setRenaming(s.id)}
                onContextMenu={(e) => {
                  e.preventDefault();
                  activateSheet(s.id);
                  setState({ menu: { x: e.clientX, y: e.clientY, kind: 'tab', sheetId: s.id } });
                }}
              >
                {renaming === s.id ? (
                  <input
                    autoFocus
                    defaultValue={s.name}
                    className="outline-none border px-1 w-[110px] text-[12px]"
                    style={{ background: 'var(--input-bg)', borderColor: 'var(--accent)' }}
                    onFocus={(e) => e.target.select()}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') {
                        if (renameSheet(s.id, (e.target as HTMLInputElement).value)) setRenaming(null);
                      } else if (e.key === 'Escape') setRenaming(null);
                      e.stopPropagation();
                    }}
                    onBlur={(e) => {
                      renameSheet(s.id, e.target.value);
                      setRenaming(null);
                    }}
                  />
                ) : (
                  <span>{s.name}</span>
                )}
              </div>
            );
          })}
        </div>
        <div className="flex items-center px-1.5">
          <button className="w-[22px] h-[22px] flex items-center justify-center rounded-full hover:bg-[var(--hover)]" style={{ border: '1px solid var(--border-strong)' }} title="New sheet (Shift+F11)" onClick={() => addSheet(wb.sheets.length)} data-testid="add-sheet">
            <Plus size={14} />
          </button>
        </div>
      </div>
      <div
        className="w-[6px] cursor-col-resize flex items-center justify-center opacity-60 hover:opacity-100"
        title="Drag to resize"
        onPointerDown={(e) => {
          const start = e.clientX;
          const w0 = width;
          const move = (ev: PointerEvent) => setTabsW(Math.max(160, Math.min((rootRef.current?.getBoundingClientRect().width ?? 1000) - 120, w0 + ev.clientX - start)));
          const up = () => {
            window.removeEventListener('pointermove', move);
            window.removeEventListener('pointerup', up);
            try {
              localStorage.setItem('myexcel.tabsW', String(Math.round(width)));
            } catch {
              /* ignore */
            }
          };
          window.addEventListener('pointermove', move);
          window.addEventListener('pointerup', up);
        }}
      >
        ⋮
      </div>
      <HScrollbar />
      {menu?.kind === 'tab' && menu.sheetId && <TabMenu x={menu.x} y={menu.y} sheetId={menu.sheetId} />}
    </div>
  );
}

function TabMenu({ x, y, sheetId }: { x: number; y: number; sheetId: string }) {
  const wb = S().wb;
  const sheet = wb.sheetById(sheetId)!;
  const close = () => setState({ menu: null });
  return (
    <Menu
      x={x}
      y={y - 280}
      onClose={close}
      items={[
        { label: 'Insert...', onClick: () => addSheet() },
        { label: 'Delete', onClick: () => deleteSheet(sheetId) },
        { label: 'Rename', onClick: () => window.dispatchEvent(new CustomEvent('xl-rename-sheet', { detail: sheetId })) },
        { label: 'Move or Copy...', onClick: () => openDialog('moveCopySheet', { sheetId }) },
        { label: 'View Code', disabled: true },
        sheet.protection ? { label: 'Unprotect Sheet...', onClick: () => (sheet.protection?.passwordHash ? openDialog('unprotectSheet') : unprotectSheet('')) } : { label: 'Protect Sheet...', onClick: () => openDialog('protectSheet') },
        { label: 'Tab Color', submenu: [{ render: (c) => <ColorPalette close={c} noneLabel="No Color" onPick={(col) => setTabColor(sheetId, col ?? undefined)} /> }] },
        { separator: true },
        { label: 'Hide', onClick: () => hideSheet(sheetId) },
        { label: 'Unhide...', disabled: !wb.sheets.some((s) => s.visibility === 'hidden'), onClick: () => openDialog('unhideSheet') },
        { separator: true },
        { label: 'Select All Sheets', disabled: true },
      ]}
    />
  );
}

function HScrollbar() {
  const [, force] = useState(0);
  useStore((s) => s.rev);
  useEffect(() => onScrollChange(() => force((n) => n + 1)), []);
  const trackRef = useRef<HTMLDivElement>(null);
  const drag = useRef<{ x: number; pos: number } | null>(null);
  const api = gridApi();
  const ext = api?.extent('x');
  const trackW = trackRef.current?.getBoundingClientRect().width ?? 300;
  const thumbW = ext ? Math.max(24, (trackW * ext.view) / ext.total) : 40;
  const thumbX = ext ? ((trackW - thumbW) * ext.pos) / Math.max(1, ext.total - ext.view) : 0;
  return (
    <div className="xl-hscroll flex-1 flex items-center min-w-[80px] mr-[17px]" onPointerDown={(e) => e.stopPropagation()}>
      <button className="xl-scroll-btn w-[17px] h-[17px]" aria-label="Scroll left" onPointerDown={() => ext && api!.setScroll('x', ext.pos - 64)}>
        <svg width="4" height="7" viewBox="0 0 4 7"><path d="M4 0 L0 3.5 L4 7Z" fill="currentColor" /></svg>
      </button>
      <div
        ref={trackRef}
        className="relative flex-1 h-[17px]"
        onPointerDown={(e) => {
          if (!ext) return;
          const r = trackRef.current!.getBoundingClientRect();
          const x = e.clientX - r.left;
          api!.setScroll('x', ext.pos + (x < thumbX ? -ext.view : ext.view));
        }}
      >
        <div
          className="xl-scroll-thumb absolute top-[4px] bottom-[4px]"
          style={{ left: thumbX, width: thumbW }}
          onPointerDown={(e) => {
            e.stopPropagation();
            if (!ext) return;
            drag.current = { x: e.clientX, pos: ext.pos };
            (e.target as HTMLElement).setPointerCapture(e.pointerId);
          }}
          onPointerMove={(e) => {
            if (!drag.current || !ext) return;
            const dx = e.clientX - drag.current.x;
            api!.setScroll('x', drag.current.pos + (dx * Math.max(1, ext.total - ext.view)) / Math.max(1, trackW - thumbW));
          }}
          onPointerUp={() => (drag.current = null)}
        />
      </div>
      <button className="xl-scroll-btn w-[17px] h-[17px]" aria-label="Scroll right" onPointerDown={() => ext && api!.setScroll('x', ext.pos + 64)}>
        <svg width="4" height="7" viewBox="0 0 4 7"><path d="M0 0 L4 3.5 L0 7Z" fill="currentColor" /></svg>
      </button>
    </div>
  );
}

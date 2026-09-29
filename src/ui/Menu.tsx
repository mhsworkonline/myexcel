'use client';
import { ChevronRight, Check } from 'lucide-react';
import { ReactNode, useEffect, useLayoutEffect, useRef, useState } from 'react';

export interface MenuItem {
  label?: string;
  icon?: ReactNode;
  shortcut?: string;
  onClick?: () => void;
  submenu?: MenuItem[];
  disabled?: boolean;
  separator?: boolean;
  checked?: boolean;
  /** Custom content (e.g. colour grids). */
  render?: (close: () => void) => ReactNode;
  testId?: string;
  bold?: boolean;
}

interface Props {
  items: MenuItem[];
  x: number;
  y: number;
  onClose: () => void;
  minWidth?: number;
  level?: number;
}

export function Menu({ items, x, y, onClose, minWidth = 200, level = 0 }: Props) {
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState({ x, y });
  const [open, setOpen] = useState<number | null>(null);
  const [subPos, setSubPos] = useState<{ x: number; y: number } | null>(null);
  const timer = useRef<number | null>(null);

  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    let nx = x;
    let ny = y;
    if (nx + r.width > window.innerWidth - 4) nx = Math.max(4, level > 0 ? x - r.width - (minWidth || 0) : window.innerWidth - r.width - 4);
    if (ny + r.height > window.innerHeight - 4) ny = Math.max(4, window.innerHeight - r.height - 4);
    setPos({ x: nx, y: ny });
  }, [x, y, level, minWidth]);

  useEffect(() => {
    if (level > 0) return;
    const onDown = (e: PointerEvent) => {
      const t = e.target as HTMLElement;
      if (!t.closest('[data-menu]')) onClose();
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    window.addEventListener('pointerdown', onDown, true);
    window.addEventListener('keydown', onKey, true);
    return () => {
      window.removeEventListener('pointerdown', onDown, true);
      window.removeEventListener('keydown', onKey, true);
    };
  }, [onClose, level]);

  return (
    <div
      ref={ref}
      data-menu
      className="xl-menu fixed z-[60] py-1"
      style={{ left: pos.x, top: pos.y, minWidth }}
      onContextMenu={(e) => e.preventDefault()}
      role="menu"
    >
      {items.map((it, i) => {
        if (it.separator) return <div key={i} className="xl-menu-sep" />;
        if (it.render) return <div key={i}>{it.render(onClose)}</div>;
        return (
          <div
            key={i}
            role="menuitem"
            data-testid={it.testId}
            aria-disabled={it.disabled}
            className={'xl-menu-item ' + (it.disabled ? 'opacity-45 pointer-events-none' : '') + (open === i ? ' xl-menu-item-open' : '')}
            onMouseEnter={(e) => {
              if (timer.current) clearTimeout(timer.current);
              const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
              if (it.submenu) {
                timer.current = window.setTimeout(() => {
                  setOpen(i);
                  setSubPos({ x: rect.right - 2, y: rect.top - 4 });
                }, 120);
              } else timer.current = window.setTimeout(() => setOpen(null), 150);
            }}
            onClick={(e) => {
              e.stopPropagation();
              if (it.submenu) {
                const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
                setOpen(i);
                setSubPos({ x: rect.right - 2, y: rect.top - 4 });
                return;
              }
              onClose();
              it.onClick?.();
            }}
          >
            <span className="xl-menu-icon">{it.checked ? <Check size={14} /> : it.icon}</span>
            <span className={'flex-1 whitespace-nowrap ' + (it.bold ? 'font-semibold' : '')}>{it.label}</span>
            {it.shortcut && <span className="ml-6 text-[11px] opacity-60">{it.shortcut}</span>}
            {it.submenu && <ChevronRight size={14} className="ml-4 opacity-70" />}
          </div>
        );
      })}
      {open !== null && items[open]?.submenu && subPos && (
        <Menu items={items[open].submenu!} x={subPos.x} y={subPos.y} onClose={onClose} level={level + 1} minWidth={180} />
      )}
    </div>
  );
}

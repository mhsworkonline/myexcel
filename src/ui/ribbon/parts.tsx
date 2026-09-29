'use client';
import { ChevronDown } from 'lucide-react';
import { ReactNode, useRef, useState } from 'react';
import { Menu, MenuItem } from '../Menu';

export function Group({ label, children, launcher, title }: { label: string; children: ReactNode; launcher?: () => void; title?: string }) {
  return (
    <div className="xl-group" data-group={label}>
      <div className="flex-1 flex items-start gap-[2px]">{children}</div>
      <div className="xl-group-label">{label}</div>
      {launcher && (
        <button className="xl-launcher" onClick={launcher} title={title ?? `${label} Settings`} aria-label={`${label} dialog launcher`}>
          <svg width="8" height="8" viewBox="0 0 8 8">
            <path d="M0.5 0.5h3M0.5 0.5v3M3 3l4.5 4.5M7.5 4v3.5H4" stroke="currentColor" fill="none" strokeWidth="1" />
          </svg>
        </button>
      )}
    </div>
  );
}

export function Col({ children, className = '' }: { children: ReactNode; className?: string }) {
  return <div className={'flex flex-col gap-[1px] ' + className}>{children}</div>;
}

export function Row({ children, className = '' }: { children: ReactNode; className?: string }) {
  return <div className={'flex items-center gap-[1px] ' + className}>{children}</div>;
}

/** Opens a dropdown (menu items or custom panel) anchored under an element. */
function useDropdown() {
  const [open, setOpen] = useState<{ x: number; y: number } | null>(null);
  const toggle = (el: HTMLElement) => {
    if (open) return setOpen(null);
    const r = el.getBoundingClientRect();
    setOpen({ x: r.left, y: r.bottom + 1 });
  };
  return { open, setOpen, toggle };
}

export interface BtnProps {
  icon?: ReactNode;
  label?: string;
  title?: string;
  onClick?: () => void;
  checked?: boolean;
  disabled?: boolean;
  menu?: MenuItem[] | (() => MenuItem[]);
  panel?: (close: () => void) => ReactNode;
  testId?: string;
  wide?: boolean;
}

export function LargeButton({ icon, label, title, onClick, checked, disabled, menu, panel, testId }: BtnProps) {
  const dd = useDropdown();
  const ref = useRef<HTMLButtonElement>(null);
  const hasDrop = !!menu || !!panel;
  const lines = (label ?? '').split('\n');
  return (
    <>
      <button
        ref={ref}
        data-testid={testId}
        title={title ?? label?.replace('\n', ' ')}
        disabled={disabled}
        className={'xl-btn-lg ' + (checked ? 'xl-btn-checked ' : '') + (disabled ? 'opacity-40' : '')}
        onClick={() => {
          if (hasDrop && !onClick) dd.toggle(ref.current!);
          else onClick?.();
        }}
      >
        <span className="h-[34px] flex items-center justify-center">{icon}</span>
        <span className="leading-[14px]">
          {lines[0]}
          {lines[1] !== undefined && <br />}
          {lines[1]}
          {hasDrop && (
            <ChevronDown
              size={10}
              className="inline ml-[1px] -mt-[1px]"
              onClick={(e) => {
                if (onClick) {
                  e.stopPropagation();
                  dd.toggle(ref.current!);
                }
              }}
            />
          )}
        </span>
      </button>
      <DropdownHost dd={dd} menu={menu} panel={panel} />
    </>
  );
}

export function SmallButton({ icon, label, title, onClick, checked, disabled, menu, panel, testId }: BtnProps) {
  const dd = useDropdown();
  const ref = useRef<HTMLDivElement>(null);
  const hasDrop = !!menu || !!panel;
  if (hasDrop && onClick) {
    // split button: main action + arrow
    return (
      <>
        <div ref={ref} className={'xl-split h-[22px] ' + (checked ? 'xl-btn-checked' : '')}>
          <button data-testid={testId} title={title ?? label} disabled={disabled} onClick={onClick} className={disabled ? 'opacity-40' : ''}>
            {icon}
            {label && <span className="text-[12px]">{label}</span>}
          </button>
          <button className="!px-[1px]" onClick={() => dd.toggle(ref.current!)} aria-label={(title ?? label ?? '') + ' options'} data-testid={testId ? testId + '-drop' : undefined}>
            <ChevronDown size={10} />
          </button>
        </div>
        <DropdownHost dd={dd} menu={menu} panel={panel} />
      </>
    );
  }
  return (
    <>
      <div ref={ref} className="inline-flex">
        <button
          data-testid={testId}
          title={title ?? label}
          disabled={disabled}
          className={'xl-btn-sm ' + (checked ? 'xl-btn-checked ' : '') + (disabled ? 'opacity-40' : '')}
          onClick={() => (hasDrop ? dd.toggle(ref.current!) : onClick?.())}
        >
          {icon}
          {label && <span>{label}</span>}
          {hasDrop && <ChevronDown size={10} />}
        </button>
      </div>
      <DropdownHost dd={dd} menu={menu} panel={panel} />
    </>
  );
}

function DropdownHost({ dd, menu, panel }: { dd: ReturnType<typeof useDropdown>; menu?: MenuItem[] | (() => MenuItem[]); panel?: (close: () => void) => ReactNode }) {
  if (!dd.open) return null;
  const close = () => dd.setOpen(null);
  if (menu) return <Menu items={typeof menu === 'function' ? menu() : menu} x={dd.open.x} y={dd.open.y} onClose={close} />;
  if (panel) return <Menu items={[{ render: (c) => panel(c) }]} x={dd.open.x} y={dd.open.y} onClose={close} />;
  return null;
}

export function Combo({
  value,
  options,
  onCommit,
  width,
  title,
  testId,
  renderOption,
}: {
  value: string;
  options: string[];
  onCommit: (v: string) => void;
  width: number;
  title?: string;
  testId?: string;
  renderOption?: (o: string) => ReactNode;
}) {
  const [text, setText] = useState<string | null>(null);
  const dd = useDropdown();
  const ref = useRef<HTMLDivElement>(null);
  return (
    <>
      <div ref={ref} className="xl-combo" style={{ width }} title={title}>
        <input
          data-testid={testId}
          value={text ?? value}
          onFocus={(e) => {
            setText(value);
            e.target.select();
          }}
          onChange={(e) => setText(e.target.value)}
          onBlur={() => setText(null)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              onCommit((e.target as HTMLInputElement).value);
              setText(null);
              (e.target as HTMLInputElement).blur();
            } else if (e.key === 'Escape') {
              setText(null);
              (e.target as HTMLInputElement).blur();
            }
          }}
        />
        <button className="h-full px-[3px] hover:bg-[var(--hover)]" onClick={() => dd.toggle(ref.current!)} aria-label={(title ?? '') + ' list'}>
          <ChevronDown size={11} />
        </button>
      </div>
      {dd.open && (
        <div className="fixed inset-0 z-[59]" onPointerDown={() => dd.setOpen(null)}>
          <div className="xl-menu absolute max-h-[360px] overflow-auto py-1" style={{ left: dd.open.x, top: dd.open.y, minWidth: width }} onPointerDown={(e) => e.stopPropagation()}>
            {options.map((o) => (
              <div
                key={o}
                className={'xl-menu-item !h-auto min-h-[24px] ' + (o === value ? 'xl-item-active' : '')}
                onClick={() => {
                  onCommit(o);
                  dd.setOpen(null);
                }}
              >
                {renderOption ? renderOption(o) : o}
              </div>
            ))}
          </div>
        </div>
      )}
    </>
  );
}

// ---------- colour palette ----------

export const THEME_BASE = ['#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];
export const STANDARD_COLORS = ['#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0'];
const THEME_NAMES = ['White', 'Black', 'Gray', 'Blue-Gray', 'Blue', 'Orange', 'Gray', 'Gold', 'Blue', 'Green'];

function shade(hex: string, t: number): string {
  const n = parseInt(hex.slice(1), 16);
  const ch = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => Math.round(t >= 0 ? v + (255 - v) * t : v * (1 + t)));
  return '#' + ch.map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
}

function variants(base: string, i: number): { c: string; label: string }[] {
  if (i === 0) return [-0.05, -0.15, -0.25, -0.35, -0.5].map((t) => ({ c: shade(base, t), label: `Darker ${Math.round(-t * 100)}%` }));
  if (i === 1) return [0.5, 0.35, 0.25, 0.15, 0.05].map((t) => ({ c: shade(base, t), label: `Lighter ${Math.round(t * 100)}%` }));
  return [
    { c: shade(base, 0.8), label: 'Lighter 80%' },
    { c: shade(base, 0.6), label: 'Lighter 60%' },
    { c: shade(base, 0.4), label: 'Lighter 40%' },
    { c: shade(base, -0.25), label: 'Darker 25%' },
    { c: shade(base, -0.5), label: 'Darker 50%' },
  ];
}

export function ColorPalette({ onPick, close, noneLabel, autoLabel }: { onPick: (c: string | null) => void; close: () => void; noneLabel?: string; autoLabel?: string }) {
  const pick = (c: string | null) => {
    onPick(c);
    close();
  };
  return (
    <div className="px-2 py-1.5 w-[190px]" onPointerDown={(e) => e.stopPropagation()}>
      {autoLabel && (
        <button className="xl-menu-item !mx-0 w-full" onClick={() => pick(null)}>
          <span className="inline-block w-3.5 h-3.5 border" style={{ background: '#000', borderColor: '#777' }} />
          {autoLabel}
        </button>
      )}
      <div className="xl-menu-header !bg-transparent !px-0 text-[11px]">Theme Colors</div>
      <div className="flex gap-[3px]">
        {THEME_BASE.map((b, i) => (
          <div key={b + i} className="flex flex-col gap-[0px]">
            <button className="xl-color-cell mb-[3px]" style={{ background: b }} title={THEME_NAMES[i]} onClick={() => pick(b)} />
            {variants(b, i).map((v) => (
              <button key={v.c} className="xl-color-cell" style={{ background: v.c }} title={`${THEME_NAMES[i]}, ${v.label}`} onClick={() => pick(v.c)} />
            ))}
          </div>
        ))}
      </div>
      <div className="xl-menu-header !bg-transparent !px-0 text-[11px] mt-1">Standard Colors</div>
      <div className="flex gap-[3px]">
        {STANDARD_COLORS.map((c) => (
          <button key={c} className="xl-color-cell" style={{ background: c }} title={c} onClick={() => pick(c)} />
        ))}
      </div>
      <div className="xl-menu-sep" />
      {noneLabel && (
        <button className="xl-menu-item !mx-0 w-full" onClick={() => pick(null)}>
          <span className="inline-block w-3.5 h-3.5 border" style={{ background: 'repeating-linear-gradient(45deg,#fff 0 3px,#ddd 3px 6px)', borderColor: '#999' }} />
          {noneLabel}
        </button>
      )}
      <label className="xl-menu-item !mx-0 w-full cursor-default">
        <span className="inline-block w-3.5 h-3.5 rounded-full" style={{ background: 'conic-gradient(red, yellow, lime, cyan, blue, magenta, red)' }} />
        More Colors...
        <input type="color" className="w-0 h-0 opacity-0" onChange={(e) => pick(e.target.value.toUpperCase())} />
      </label>
    </div>
  );
}

/** Icon with a colour bar underneath (Fill Color / Font Color). */
export function ColorIcon({ icon, color }: { icon: ReactNode; color: string }) {
  return (
    <span className="relative inline-flex flex-col items-center leading-none">
      {icon}
      <span className="block w-[16px] h-[4px] -mt-[1px]" style={{ background: color }} />
    </span>
  );
}

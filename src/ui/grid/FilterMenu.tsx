'use client';
import { ArrowDownAZ, ArrowUpAZ, ArrowDownWideNarrow, ArrowUpNarrowWide, FilterX, Search } from 'lucide-react';
import { useEffect, useMemo, useRef, useState } from 'react';
import type { ColumnFilter } from '../../model/types';
import { columnValues, displayTextAt, setColumnFilter, sortByColor, sortByColumn } from '../../state/actions/sortFilter';
import { openDialog, S, setState, useStore } from '../../state/store';
import { Menu, MenuItem } from '../Menu';
import { colToName } from '../../model/address';

export function FilterMenu() {
  const fm = useStore((s) => s.filterMenu)!;
  const sheet = S().wb.activeSheet;
  const af = sheet.autoFilter;
  const col = fm.col;
  const data = useMemo(() => columnValues(sheet, col), [sheet, col]);
  const current = af?.filters[col];
  const initial = useMemo(() => {
    const set = new Set<string>();
    if (current?.type === 'values') current.values.forEach((v) => set.add(v));
    else data.values.forEach((v) => set.add(v.text));
    return set;
  }, [current, data]);
  const [checked, setChecked] = useState<Set<string>>(initial);
  const [blanks, setBlanks] = useState(current?.type === 'values' ? current.blanks : true);
  const [q, setQ] = useState('');
  const [sub, setSub] = useState<{ items: MenuItem[]; x: number; y: number } | null>(null);
  const ref = useRef<HTMLDivElement>(null);
  const header = af ? displayTextAt(sheet, af.range.r1, col) || `Column ${colToName(col)}` : '';
  const close = () => setState({ filterMenu: null });

  useEffect(() => {
    const onDown = (e: PointerEvent) => {
      const t = e.target as HTMLElement;
      if (!t.closest('[data-filter-menu]') && !t.closest('[data-menu]')) close();
    };
    window.addEventListener('pointerdown', onDown, true);
    return () => window.removeEventListener('pointerdown', onDown, true);
  }, []);

  if (!af) return null;
  const visible = data.values.filter((v) => !q || v.text.toLowerCase().includes(q.toLowerCase()));
  const allChecked = visible.every((v) => checked.has(v.text)) && (!data.hasBlanks || blanks);
  const someChecked = visible.some((v) => checked.has(v.text)) || blanks;
  const numeric = data.numeric;

  const apply = () => {
    let f: ColumnFilter | null;
    if (q) {
      f = { type: 'values', values: visible.filter((v) => checked.has(v.text)).map((v) => v.text), blanks: false };
    } else if (allChecked) f = null;
    else f = { type: 'values', values: [...checked].filter((t) => data.values.some((v) => v.text === t)), blanks: data.hasBlanks ? blanks : false };
    setColumnFilter(col, f);
    close();
  };

  const custom = (op: string, op2?: string) => () => {
    close();
    openDialog('customFilter', { col, op, op2 });
  };
  const textItems: MenuItem[] = [
    { label: 'Equals...', onClick: custom('eq') },
    { label: 'Does Not Equal...', onClick: custom('ne') },
    { separator: true },
    { label: 'Begins With...', onClick: custom('beginsWith') },
    { label: 'Ends With...', onClick: custom('endsWith') },
    { separator: true },
    { label: 'Contains...', onClick: custom('contains') },
    { label: 'Does Not Contain...', onClick: custom('notContains') },
    { separator: true },
    { label: 'Custom Filter...', onClick: custom('eq') },
  ];
  const numberItems: MenuItem[] = [
    { label: 'Equals...', onClick: custom('eq') },
    { label: 'Does Not Equal...', onClick: custom('ne') },
    { separator: true },
    { label: 'Greater Than...', onClick: custom('gt') },
    { label: 'Greater Than Or Equal To...', onClick: custom('gte') },
    { label: 'Less Than...', onClick: custom('lt') },
    { label: 'Less Than Or Equal To...', onClick: custom('lte') },
    { label: 'Between...', onClick: custom('gte', 'lte') },
    { separator: true },
    { label: 'Top 10...', onClick: () => { close(); openDialog('top10Filter', { col }); } },
    { label: 'Above Average', onClick: () => { setColumnFilter(col, { type: 'dynamic', kind: 'aboveAverage' }); close(); } },
    { label: 'Below Average', onClick: () => { setColumnFilter(col, { type: 'dynamic', kind: 'belowAverage' }); close(); } },
    { separator: true },
    { label: 'Custom Filter...', onClick: custom('eq') },
  ];
  // colours present in the column
  const fills = new Set<string>();
  const fonts = new Set<string>();
  for (let r = af.range.r1 + 1; r <= af.range.r2; r++) {
    const s = S().wb.styles.get(sheet.styleIdAt(r, col));
    if (s.fillColor) fills.add(s.fillColor);
    if (s.fontColor) fonts.add(s.fontColor);
  }
  const swatch = (c: string) => <span className="inline-block w-10 h-3 border" style={{ background: c, borderColor: '#999' }} />;
  const colorItems = (font: boolean, sort: boolean): MenuItem[] => {
    const list = [...(font ? fonts : fills)];
    if (!list.length) return [{ label: 'No colors', disabled: true }];
    return list.map((c) => ({
      label: '',
      icon: swatch(c),
      onClick: () => {
        if (sort) sortByColor(col, c, font);
        else setColumnFilter(col, { type: 'color', color: c, font });
        close();
      },
    }));
  };

  const openSub = (items: MenuItem[]) => (e: React.MouseEvent) => {
    const r = (e.currentTarget as HTMLElement).getBoundingClientRect();
    setSub({ items, x: r.right - 2, y: r.top - 4 });
  };

  const x = Math.min(fm.x, window.innerWidth - 270);
  const y = Math.min(fm.y + 2, window.innerHeight - 440);
  return (
    <>
      <div ref={ref} data-filter-menu className="xl-menu fixed z-[55] w-[260px] py-1 text-[12px]" style={{ left: x, top: y }} onPointerDown={(e) => e.stopPropagation()}>
        <MenuRow icon={numeric ? <ArrowUpNarrowWide size={15} /> : <ArrowDownAZ size={15} />} label={numeric ? 'Sort Smallest to Largest' : 'Sort A to Z'} onClick={() => { sortByColumn(col, false); close(); }} onEnter={() => setSub(null)} />
        <MenuRow icon={numeric ? <ArrowDownWideNarrow size={15} /> : <ArrowUpAZ size={15} />} label={numeric ? 'Sort Largest to Smallest' : 'Sort Z to A'} onClick={() => { sortByColumn(col, true); close(); }} onEnter={() => setSub(null)} />
        <MenuRow label="Sort by Color" arrow onEnter={openSub([{ label: 'Sort by Cell Color', disabled: true }, ...colorItems(false, true), { separator: true }, { label: 'Sort by Font Color', disabled: true }, ...colorItems(true, true)])} />
        <div className="xl-menu-sep" />
        <MenuRow icon={<FilterX size={15} />} label={`Clear Filter From "${header}"`} disabled={!current} onClick={() => { setColumnFilter(col, null); close(); }} onEnter={() => setSub(null)} />
        <MenuRow label="Filter by Color" arrow onEnter={openSub([{ label: 'Filter by Cell Color', disabled: true }, ...colorItems(false, false), { separator: true }, { label: 'Filter by Font Color', disabled: true }, ...colorItems(true, false)])} />
        <MenuRow label={numeric ? 'Number Filters' : 'Text Filters'} arrow checked={current?.type === 'custom' || current?.type === 'top10' || current?.type === 'dynamic'} onEnter={openSub(numeric ? numberItems : textItems)} />
        <div className="px-2 pt-1" onMouseEnter={() => setSub(null)}>
          <div className="flex items-center border px-1.5 h-[24px]" style={{ borderColor: 'var(--border-strong)', background: 'var(--input-bg)' }}>
            <input className="flex-1 bg-transparent outline-none text-[12px]" placeholder="Search" value={q} onChange={(e) => setQ(e.target.value)} data-testid="filter-search" />
            <Search size={13} className="opacity-60" />
          </div>
          <div className="mt-1 h-[170px] overflow-auto border px-1 py-0.5" style={{ borderColor: 'var(--border-strong)', background: 'var(--input-bg)' }}>
            <label className="flex items-center gap-1.5 py-[1px] cursor-default">
              <input
                type="checkbox"
                checked={allChecked}
                ref={(el) => {
                  if (el) el.indeterminate = !allChecked && someChecked;
                }}
                onChange={() => {
                  if (allChecked) {
                    setChecked(new Set());
                    setBlanks(false);
                  } else {
                    setChecked(new Set(data.values.map((v) => v.text)));
                    setBlanks(true);
                  }
                }}
              />
              {q ? '(Select All Search Results)' : '(Select All)'}
            </label>
            {visible.map((v) => (
              <label key={v.text} className="flex items-center gap-1.5 py-[1px] cursor-default">
                <input
                  type="checkbox"
                  checked={checked.has(v.text)}
                  onChange={() => {
                    const n = new Set(checked);
                    if (n.has(v.text)) n.delete(v.text);
                    else n.add(v.text);
                    setChecked(n);
                  }}
                />
                <span className="truncate">{v.text}</span>
              </label>
            ))}
            {data.hasBlanks && !q && (
              <label className="flex items-center gap-1.5 py-[1px]">
                <input type="checkbox" checked={blanks} onChange={() => setBlanks(!blanks)} />
                (Blanks)
              </label>
            )}
          </div>
          <div className="flex justify-end gap-2 py-2">
            <button className="xl-btn xl-btn-primary" onClick={apply} data-testid="filter-ok">
              OK
            </button>
            <button className="xl-btn" onClick={close}>
              Cancel
            </button>
          </div>
        </div>
      </div>
      {sub && <Menu items={sub.items} x={sub.x} y={sub.y} onClose={() => { setSub(null); close(); }} level={1} minWidth={190} />}
    </>
  );
}

function MenuRow({ icon, label, onClick, onEnter, arrow, disabled, checked }: { icon?: React.ReactNode; label: string; onClick?: () => void; onEnter?: (e: React.MouseEvent) => void; arrow?: boolean; disabled?: boolean; checked?: boolean }) {
  return (
    <div className={'xl-menu-item ' + (disabled ? 'opacity-45 pointer-events-none' : '')} onClick={onClick} onMouseEnter={onEnter}>
      <span className="xl-menu-icon">{checked ? '✓' : icon}</span>
      <span className="flex-1 truncate">{label}</span>
      {arrow && <span className="opacity-60">›</span>}
    </div>
  );
}

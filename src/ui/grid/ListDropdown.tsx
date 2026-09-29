'use client';
import { useEffect, useState } from 'react';
import { transact, S, setState, useStore } from '../../state/store';
import { listItems, validationAt } from '../../state/validation';
import { writeInput } from '../../state/actions/edit';

/** Data validation in-cell dropdown list (also Alt+Down). */
export function ListDropdown({ getRect }: { getRect: (r: number, c: number) => { x: number; y: number; w: number; h: number } | null }) {
  const lm = useStore((s) => s.listMenu)!;
  const sheet = S().wb.activeSheet;
  const dv = validationAt(sheet, lm.r, lm.c);
  const items = dv && dv.type === 'list' ? listItems(sheet, dv) : [];
  const [idx, setIdx] = useState(0);
  const close = () => setState({ listMenu: null });
  const pick = (v: string) => {
    transact('Typing', (tx) => writeInput(tx, sheet, lm.r, lm.c, v));
    close();
  };
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') close();
      else if (e.key === 'ArrowDown') setIdx((i) => Math.min(items.length - 1, i + 1));
      else if (e.key === 'ArrowUp') setIdx((i) => Math.max(0, i - 1));
      else if (e.key === 'Enter' && items[idx] !== undefined) pick(items[idx]);
      else return;
      e.preventDefault();
      e.stopPropagation();
    };
    const onDown = (e: PointerEvent) => {
      if (!(e.target as HTMLElement).closest('[data-list-dd]')) close();
    };
    window.addEventListener('keydown', onKey, true);
    window.addEventListener('pointerdown', onDown, true);
    return () => {
      window.removeEventListener('keydown', onKey, true);
      window.removeEventListener('pointerdown', onDown, true);
    };
  });
  const R = getRect(lm.r, lm.c);
  if (!R || !items.length) return null;
  return (
    <div data-list-dd className="xl-popup absolute z-40 max-h-[160px] overflow-auto text-[12px] py-0.5" style={{ left: R.x, top: R.y + R.h + 1, minWidth: R.w + 18 }} onPointerDown={(e) => e.stopPropagation()}>
      {items.map((it, i) => (
        <div key={i} className={'px-1.5 py-[1px] cursor-default ' + (i === idx ? 'xl-item-active' : '')} onMouseEnter={() => setIdx(i)} onClick={() => pick(it)}>
          {it}
        </div>
      ))}
    </div>
  );
}

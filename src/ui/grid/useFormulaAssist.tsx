'use client';
import { useMemo, useState } from 'react';
import { Engine } from '../../engine/Engine';
import { functionContextAt, partialIdentAt } from '../../model/formula';
import { argList, fnCatalog, fnInfo } from '../../model/functions';
import { S, setState } from '../../state/store';
import type { AutoState } from './editKeys';

/** Function autocomplete + argument-hint state for a formula input. */
export function useFormulaAssist(text: string, caret: number) {
  const [dismissedAt, setDismissedAt] = useState<string | null>(null);
  const [index, setIndex] = useState(0);
  const partial = useMemo(() => partialIdentAt(text, caret), [text, caret]);
  const items = useMemo(() => {
    if (!partial || partial.text.length < 1) return [];
    const p = partial.text.toUpperCase();
    const fns = Engine.functionNames().filter((n) => n.startsWith(p));
    const cat = fnCatalog();
    fns.sort((a, b) => (cat.has(a) === cat.has(b) ? a.localeCompare(b) : cat.has(a) ? -1 : 1));
    const names = S().wb.names.map((n) => n.name).filter((n) => n.toUpperCase().startsWith(p));
    const tables = S().wb.sheets.flatMap((s) => s.tables.map((t) => t.name)).filter((n) => n.toUpperCase().startsWith(p));
    return [...names, ...tables, ...fns].slice(0, 12);
  }, [partial]);
  const key = partial ? `${partial.start}:${partial.text}` : null;
  const open = items.length > 0 && dismissedAt !== key;
  const auto: AutoState | null = open && partial ? { items, index: Math.min(index, items.length - 1), start: partial.start } : null;

  const hint = useMemo(() => {
    if (!text.startsWith('=')) return null;
    const ctx = functionContextAt(text, caret);
    if (!ctx) return null;
    const info = fnInfo(ctx.name);
    return { name: info.name, args: argList(info.args), argIndex: ctx.argIndex };
  }, [text, caret]);

  const accept = (name: string) => {
    const ed = S().edit;
    if (!ed || !partial) return;
    const isFn = Engine.functionNames().includes(name);
    const insert = isFn ? name + '(' : name;
    const t = ed.text.slice(0, partial.start) + insert + ed.text.slice(caret);
    setState({ edit: { ...ed, text: t, caret: partial.start + insert.length, point: undefined } });
    setIndex(0);
  };
  const close = () => setDismissedAt(key);
  const move = (d: number) => setIndex((i) => Math.max(0, Math.min(items.length - 1, i + d)));
  return { auto, hint, accept, close, move, setIndex };
}

export function AutoList({ auto, onPick, onHover, style }: { auto: AutoState; onPick: (n: string) => void; onHover: (i: number) => void; style: React.CSSProperties }) {
  const sel = auto.items[auto.index];
  const info = fnCatalog().get(sel);
  return (
    <div className="absolute z-40 flex items-start gap-1" style={style} onPointerDown={(e) => e.stopPropagation()}>
      <div className="xl-popup min-w-[180px] py-0.5 text-[12px]" role="listbox">
        {auto.items.map((n, i) => (
          <div
            key={n}
            role="option"
            aria-selected={i === auto.index}
            className={'px-2 py-[1px] flex items-center gap-1.5 cursor-default ' + (i === auto.index ? 'xl-item-active' : '')}
            onMouseEnter={() => onHover(i)}
            onMouseDown={(e) => {
              e.preventDefault();
              onPick(n);
            }}
          >
            <span className="inline-block w-3 text-center text-[10px] opacity-70">{fnCatalog().has(n) || Engine.functionNames().includes(n) ? 'ƒx' : '⊞'}</span>
            {n}
          </div>
        ))}
      </div>
      {info && <div className="xl-popup max-w-[260px] px-2 py-1 text-[11.5px]">{info.desc}</div>}
    </div>
  );
}

export function ArgHint({ hint, style }: { hint: { name: string; args: string[]; argIndex: number }; style: React.CSSProperties }) {
  const repeat = hint.args[hint.args.length - 1] === '...';
  const idx = repeat && hint.argIndex >= hint.args.length - 1 ? hint.args.length - 2 : hint.argIndex;
  return (
    <div className="absolute z-40 xl-popup px-1.5 py-[1px] text-[11.5px] whitespace-nowrap" style={style}>
      <span className="underline">{hint.name}</span>(
      {hint.args.map((a, i) => (
        <span key={i}>
          {i > 0 && ', '}
          <span className={i === idx ? 'font-bold' : ''}>{a}</span>
        </span>
      ))}
      )
    </div>
  );
}

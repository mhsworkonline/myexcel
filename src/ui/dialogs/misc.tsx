'use client';
import { useMemo, useState } from 'react';
import { Engine, isErrorVal } from '../../engine/Engine';
import { argList, fnCatalog, fnInfo, FN_CATEGORIES } from '../../model/functions';
import { beginEdit } from '../../state/actions/edit';
import { setHyperlink, setNote } from '../../state/actions/data';
import { closeDialog, S, setState } from '../../state/store';
import { getComputed } from '../../state/values';
import { Dialog, RefInput } from './Dialog';

const recent: string[] = ['SUM', 'AVERAGE', 'IF', 'HYPERLINK', 'COUNT', 'MAX', 'SIN', 'SUMIF', 'PMT', 'STDEV'];

export function InsertFunctionDialog() {
  const [q, setQ] = useState('');
  const [cat, setCat] = useState('Most Recently Used');
  const [sel, setSel] = useState('SUM');
  const cat2 = fnCatalog();
  const list = useMemo(() => {
    if (q) {
      const ql = q.toLowerCase();
      return Engine.functionNames().filter((n) => n.toLowerCase().includes(ql) || (cat2.get(n)?.desc ?? '').toLowerCase().includes(ql));
    }
    if (cat === 'Most Recently Used') return recent;
    if (cat === 'All') return Engine.functionNames();
    return [...cat2.values()].filter((f) => f.cat === cat).map((f) => f.name);
  }, [q, cat, cat2]);
  const info = fnInfo(sel);
  const insert = (name: string) => {
    closeDialog();
    if (!recent.includes(name)) recent.unshift(name);
    setTimeout(() => import('../../state/store').then((m) => m.openDialog('functionArgs', { name })), 0);
  };
  return (
    <Dialog title="Insert Function" width={480} onOk={() => { insert(sel); return false; }} testId="insert-function">
      <div className="mb-1">Search for a function:</div>
      <div className="flex gap-2 mb-2">
        <input className="xl-input flex-1" placeholder="Type a brief description of what you want to do and then click Go" value={q} onChange={(e) => setQ(e.target.value)} />
        <button className="xl-btn !min-w-0">Go</button>
      </div>
      <label className="flex items-center gap-2 mb-2">
        Or select a category:
        <select className="xl-input flex-1" value={cat} onChange={(e) => { setCat(e.target.value); setQ(''); }}>
          {FN_CATEGORIES.map((c) => <option key={c}>{c}</option>)}
        </select>
      </label>
      <div className="mb-1">Select a function:</div>
      <div className="xl-list h-[160px]">
        {list.map((n) => (
          <div key={n} className={sel === n ? 'xl-list-sel' : ''} onClick={() => setSel(n)} onDoubleClick={() => insert(n)}>{n}</div>
        ))}
      </div>
      <div className="mt-2 font-semibold">{info.name}({info.args})</div>
      <div className="mt-1 min-h-[36px] opacity-85">{info.desc || 'No description available.'}</div>
    </Dialog>
  );
}

export function FunctionArgsDialog({ props }: { props: Record<string, unknown> }) {
  const name = props.name as string;
  const info = fnInfo(name);
  const labels = argList(info.args).filter((a) => a !== '...');
  const [vals, setVals] = useState<string[]>(() => labels.map(() => ''));
  const formula = `=${name}(${vals.filter((v, i) => v !== '' || i < vals.findLastIndex((x) => x !== '')).join(',')})`;
  const preview = useMemo(() => {
    try {
      const v = S().engine.evaluate(formula, S().wb.activeSheet);
      if (isErrorVal(v)) return v.error;
      return String(v ?? '');
    } catch {
      return '';
    }
  }, [formula]);
  const evalArg = (t: string) => {
    if (!t) return '';
    const v = S().engine.evaluate(t.startsWith('=') ? t : '=' + t, S().wb.activeSheet);
    return isErrorVal(v) ? v.error : String(v ?? '');
  };
  return (
    <Dialog
      title="Function Arguments"
      width={520}
      onOk={() => {
        const st = S();
        const sheet = st.wb.activeSheet;
        const cur = st.edit;
        if (cur && cur.text.startsWith('=') && cur.text.length > 1) {
          setState({ edit: { ...cur, text: cur.text.slice(0, cur.caret) + formula.slice(1) + cur.text.slice(cur.caret), caret: cur.caret + formula.length - 1 } });
        } else {
          import('../../state/store').then(({ transact }) =>
            import('../../state/actions/edit').then(({ writeInput, autoFormatFormulaResults }) =>
              transact('Insert Function', (tx) => {
                writeInput(tx, sheet, st.sel.active.r, st.sel.active.c, formula);
                autoFormatFormulaResults(tx, sheet, [st.sel.active]);
              }),
            ),
          );
        }
        void getComputed;
      }}
      testId="function-args"
    >
      <fieldset className="xl-fieldset">
        <legend>{name}</legend>
        {labels.map((l, i) => (
          <div key={i} className="flex items-center gap-2 my-1">
            <span className={'w-[110px] truncate ' + (l.startsWith('[') ? '' : 'font-semibold')}>{l.replace(/[[\]]/g, '')}</span>
            <RefInput value={vals[i]} onChange={(v) => setVals(vals.map((x, k) => (k === i ? v.replace(/^=/, '') : x)))} width={200} />
            <span className="truncate opacity-75 text-[11px] flex-1">= {evalArg(vals[i])}</span>
          </div>
        ))}
        {!labels.length && <div className="opacity-70">This function takes no arguments.</div>}
      </fieldset>
      <div className="mt-2 opacity-85">{info.desc}</div>
      <div className="mt-2">Formula result = <b>{preview}</b></div>
      <div className="text-[11px] opacity-60 mt-1">{formula}</div>
    </Dialog>
  );
}

export function NoteDialog({ props }: { props: Record<string, unknown> }) {
  const st = S();
  const sheet = st.wb.activeSheet;
  const { r, c } = st.sel.active;
  const cur = sheet.getCell(r, c)?.note;
  const threaded = !!props.threaded || !!cur?.threaded;
  const [text, setText] = useState(cur?.text ?? '');
  const [reply, setReply] = useState('');
  return (
    <Dialog
      title={threaded ? 'Comment' : 'Note'}
      width={360}
      okLabel={threaded && cur ? 'Reply' : 'Post'}
      onOk={() => {
        if (threaded && cur) {
          if (!reply.trim()) return;
          setNote(r, c, { ...cur, replies: [...(cur.replies ?? []), { author: 'Me', text: reply, time: Date.now() }] });
          return;
        }
        if (!text.trim()) setNote(r, c, undefined);
        else setNote(r, c, { text, author: cur?.author ?? (threaded ? 'Me' : undefined), threaded, time: Date.now() });
      }}
      extraButtons={cur ? <button className="xl-btn" onClick={() => { setNote(r, c, undefined); closeDialog(); }}>Delete</button> : undefined}
      testId="note-dialog"
    >
      {threaded && cur ? (
        <div className="flex flex-col gap-2">
          <div className="p-2 rounded" style={{ background: 'var(--chrome-bg)' }}>
            <div className="font-semibold">{cur.author ?? 'Me'}</div>
            <div className="whitespace-pre-wrap">{cur.text}</div>
          </div>
          {cur.replies?.map((rp, i) => (
            <div key={i} className="p-2 rounded ml-3" style={{ background: 'var(--chrome-bg)' }}>
              <div className="font-semibold">{rp.author}</div>
              <div className="whitespace-pre-wrap">{rp.text}</div>
            </div>
          ))}
          <textarea className="xl-input !h-16 py-1" placeholder="Reply…" value={reply} onChange={(e) => setReply(e.target.value)} />
        </div>
      ) : (
        <textarea className="xl-input w-full !h-28 py-1" style={{ background: '#FFFFE1', color: '#000' }} value={text} onChange={(e) => setText(e.target.value)} placeholder={threaded ? 'Start a conversation' : 'Type a note'} data-testid="note-text" autoFocus />
      )}
    </Dialog>
  );
}

export function HyperlinkDialog() {
  const st = S();
  const sheet = st.wb.activeSheet;
  const { r, c } = st.sel.active;
  const cell = sheet.getCell(r, c);
  const [kind, setKind] = useState<'web' | 'place' | 'email'>(cell?.link?.startsWith('#') ? 'place' : cell?.link?.startsWith('mailto:') ? 'email' : 'web');
  const [text, setText] = useState(String(getComputed(sheet, r, c) ?? ''));
  const [addr, setAddr] = useState(cell?.link?.replace(/^mailto:|^#/, '') ?? '');
  const [place, setPlace] = useState(st.wb.sheets[0].name);
  const [cellRef, setCellRef] = useState('A1');
  return (
    <Dialog
      title={cell?.link ? 'Edit Hyperlink' : 'Insert Hyperlink'}
      width={500}
      onOk={() => {
        const link = kind === 'web' ? addr : kind === 'email' ? 'mailto:' + addr : place ? `#'${place}'!${cellRef}` : `#${cellRef}`;
        if (!addr && kind !== 'place') return;
        setHyperlink(link, text);
      }}
      extraButtons={cell?.link ? <button className="xl-btn" onClick={() => { setHyperlink(undefined); closeDialog(); }}>Remove Link</button> : undefined}
      testId="hyperlink-dialog"
    >
      <div className="flex gap-3">
        <div className="flex flex-col gap-1 w-[120px]">
          {([['web', 'Existing File or Web Page'], ['place', 'Place in This Document'], ['email', 'E-mail Address']] as const).map(([k, l]) => (
            <button key={k} className={'text-left px-2 py-2 rounded ' + (kind === k ? 'xl-item-active' : 'hover:bg-[var(--hover)]')} onClick={() => setKind(k)}>{l}</button>
          ))}
        </div>
        <div className="flex-1 flex flex-col gap-2">
          <label className="flex flex-col gap-1">Text to display:<input className="xl-input" value={text} onChange={(e) => setText(e.target.value)} /></label>
          {kind === 'web' && <label className="flex flex-col gap-1">Address:<input className="xl-input" value={addr} onChange={(e) => setAddr(e.target.value)} placeholder="https://" data-testid="link-address" /></label>}
          {kind === 'email' && <label className="flex flex-col gap-1">E-mail address:<input className="xl-input" value={addr} onChange={(e) => setAddr(e.target.value)} /></label>}
          {kind === 'place' && (
            <>
              <label className="flex flex-col gap-1">Type the cell reference:<input className="xl-input" value={cellRef} onChange={(e) => setCellRef(e.target.value)} /></label>
              <div>Or select a place in this document:</div>
              <div className="xl-list h-[100px]">
                {st.wb.sheets.map((s) => <div key={s.id} className={place === s.name ? 'xl-list-sel' : ''} onClick={() => setPlace(s.name)}>{s.name}</div>)}
                {st.wb.names.map((n) => <div key={n.name} onClick={() => { setPlace(''); setCellRef(n.name); }}>{n.name}</div>)}
              </div>
            </>
          )}
        </div>
      </div>
    </Dialog>
  );
}

export { beginEdit };

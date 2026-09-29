'use client';
import { ChevronDown, ChevronUp, Plus, Trash2, Copy } from 'lucide-react';
import { useState } from 'react';
import { colToName } from '../../model/address';
import type { FilterOp } from '../../model/types';
import { DEFAULT_PASTE, pasteFromSystem, PasteOptions, PasteWhat } from '../../state/actions/clipboard';
import { defaultFind, findAll, FindHit, findNext, FindOptions, gotoHit, replaceAll, replaceCurrent } from '../../state/actions/find';
import { dataRangeForCommand, displayTextAt, guessHeader, setColumnFilter, SortKey, sortRange } from '../../state/actions/sortFilter';
import { closeDialog, S } from '../../state/store';
import { Dialog } from './Dialog';

let lastFind: FindOptions = { ...defaultFind };
let lastReplace = '';

export function FindReplaceDialog({ props }: { props: Record<string, unknown> }) {
  const [tab, setTab] = useState<'find' | 'replace'>((props.tab as 'find' | 'replace') ?? 'find');
  const [o, setO] = useState<FindOptions>(lastFind);
  const [withText, setWith] = useState(lastReplace);
  const [showOpts, setShowOpts] = useState(false);
  const [hits, setHits] = useState<FindHit[] | null>(null);
  const upd = (p: Partial<FindOptions>) => {
    const n = { ...o, ...p };
    setO(n);
    lastFind = n;
  };
  const footer = (
    <div className="flex items-center gap-2 px-4 py-3 flex-wrap">
      {tab === 'replace' && (
        <>
          <button className="xl-btn" onClick={() => { lastReplace = withText; replaceAll(o, withText); }} data-testid="replace-all">Replace All</button>
          <button className="xl-btn" onClick={() => { lastReplace = withText; replaceCurrent(o, withText); }}>Replace</button>
        </>
      )}
      <div className="flex-1" />
      <button className="xl-btn" onClick={() => setHits(findAll(o))} data-testid="find-all">Find All</button>
      <button className="xl-btn xl-btn-primary" onClick={(e) => findNext(o, e.shiftKey)} data-testid="find-next">Find Next</button>
      <button className="xl-btn" onClick={() => closeDialog()}>Close</button>
    </div>
  );
  return (
    <Dialog title="Find and Replace" width={520} footer={footer} modal={false} testId="find-dialog" onOk={() => { findNext(o); return false; }}>
      <div className="flex border-b mb-3" style={{ borderColor: 'var(--border)' }}>
        <button className={'xl-dtab ' + (tab === 'find' ? 'xl-dtab-active' : '')} onClick={() => setTab('find')}>Find</button>
        <button className={'xl-dtab ' + (tab === 'replace' ? 'xl-dtab-active' : '')} onClick={() => setTab('replace')}>Replace</button>
      </div>
      <label className="flex items-center gap-2 my-1">
        <span className="w-24">Find what:</span>
        <input className="xl-input flex-1" value={o.what} onChange={(e) => upd({ what: e.target.value })} data-testid="find-what" />
      </label>
      {tab === 'replace' && (
        <label className="flex items-center gap-2 my-1">
          <span className="w-24">Replace with:</span>
          <input className="xl-input flex-1" value={withText} onChange={(e) => setWith(e.target.value)} data-testid="replace-with" />
        </label>
      )}
      {showOpts && (
        <div className="grid grid-cols-2 gap-2 mt-2">
          <div className="flex flex-col gap-1">
            <label className="flex items-center gap-2"><span className="w-16">Within:</span>
              <select className="xl-input flex-1" value={o.within} onChange={(e) => upd({ within: e.target.value as FindOptions['within'] })}><option value="sheet">Sheet</option><option value="workbook">Workbook</option></select>
            </label>
            <label className="flex items-center gap-2"><span className="w-16">Search:</span>
              <select className="xl-input flex-1" value={o.searchBy} onChange={(e) => upd({ searchBy: e.target.value as FindOptions['searchBy'] })}><option value="rows">By Rows</option><option value="cols">By Columns</option></select>
            </label>
            <label className="flex items-center gap-2"><span className="w-16">Look in:</span>
              <select className="xl-input flex-1" value={o.lookIn} onChange={(e) => upd({ lookIn: e.target.value as FindOptions['lookIn'] })}><option value="formulas">Formulas</option><option value="values">Values</option><option value="notes">Notes</option></select>
            </label>
          </div>
          <div className="flex flex-col gap-1">
            <label className="flex gap-2"><input type="checkbox" checked={o.matchCase} onChange={(e) => upd({ matchCase: e.target.checked })} /> Match case</label>
            <label className="flex gap-2"><input type="checkbox" checked={o.entireCell} onChange={(e) => upd({ entireCell: e.target.checked })} /> Match entire cell contents</label>
          </div>
        </div>
      )}
      <div className="flex justify-end mt-2">
        <button className="xl-btn" onClick={() => setShowOpts(!showOpts)}>{showOpts ? 'Options <<' : 'Options >>'}</button>
      </div>
      {hits && (
        <div className="mt-2">
          <div className="xl-list h-[140px] text-[11.5px]">
            <div className="grid grid-cols-[90px_80px_60px_1fr_1fr] font-semibold sticky top-0" style={{ background: 'var(--chrome-bg)' }}>
              <span>Sheet</span><span>Name</span><span>Cell</span><span>Value</span><span>Formula</span>
            </div>
            {hits.map((h, i) => (
              <div key={i} className="grid grid-cols-[90px_80px_60px_1fr_1fr]" onClick={() => gotoHit(h)}>
                <span className="truncate">{h.sheetName}</span><span /><span>${colToName(h.c)}${h.r + 1}</span><span className="truncate">{h.value}</span><span className="truncate">{h.formula}</span>
              </div>
            ))}
          </div>
          <div className="text-[11px] mt-1">{hits.length} cell(s) found</div>
        </div>
      )}
    </Dialog>
  );
}

interface Level extends SortKey {
  key: number;
}

export function SortDialog() {
  const st = S();
  const sheet = st.wb.activeSheet;
  const af = sheet.autoFilter;
  const a = st.sel.active;
  const inAf = af && a.r >= af.range.r1 && a.r <= af.range.r2 && a.c >= af.range.c1 && a.c <= af.range.c2;
  const rg0 = inAf ? af!.range : dataRangeForCommand();
  const [header, setHeader] = useState(rg0 ? (inAf ? true : guessHeader(sheet, rg0)) : false);
  const [levels, setLevels] = useState<Level[]>([{ key: 1, col: rg0 ? Math.max(rg0.c1, Math.min(rg0.c2, a.c)) : 0, desc: false, by: 'value' }]);
  const [sel, setSel] = useState(0);
  const [caseSensitive, setCase] = useState(false);
  const [ltr, setLtr] = useState(false);
  const [showOpts, setShowOpts] = useState(false);
  if (!rg0) {
    return (
      <Dialog title="Sort" width={380} cancelLabel={null}>
        <div className="py-3">Select a cell or range in your data first.</div>
      </Dialog>
    );
  }
  const rg = rg0;
  const colLabel = (c: number) => (header ? displayTextAt(sheet, rg.r1, c) || `Column ${colToName(c)}` : `Column ${colToName(c)}`);
  const rowLabel = (r: number) => `Row ${r + 1}`;
  const keys = ltr ? Array.from({ length: rg.r2 - rg.r1 + 1 }, (_, i) => rg.r1 + i) : Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => rg.c1 + i);
  const upd = (i: number, p: Partial<Level>) => setLevels(levels.map((l, k) => (k === i ? { ...l, ...p } : l)));
  const lists = [['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'], ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'], ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'], ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'], ['Low', 'Medium', 'High']];
  return (
    <Dialog title="Sort" width={640} onOk={() => sortRange(rg, levels, header, { caseSensitive, leftToRight: ltr })} testId="sort-dialog">
      <div className="flex items-center gap-1.5 mb-2">
        <button className="xl-btn !min-w-0 flex items-center gap-1" onClick={() => setLevels([...levels, { key: Date.now(), col: keys[0], desc: false, by: 'value' }])}><Plus size={13} /> Add Level</button>
        <button className="xl-btn !min-w-0 flex items-center gap-1" disabled={levels.length <= 1} onClick={() => { setLevels(levels.filter((_, i) => i !== sel)); setSel(Math.max(0, sel - 1)); }}><Trash2 size={13} /> Delete Level</button>
        <button className="xl-btn !min-w-0 flex items-center gap-1" onClick={() => setLevels([...levels.slice(0, sel + 1), { ...levels[sel], key: Date.now() }, ...levels.slice(sel + 1)])}><Copy size={13} /> Copy Level</button>
        <button className="xl-btn !min-w-0" disabled={sel === 0} onClick={() => { const n = [...levels]; [n[sel - 1], n[sel]] = [n[sel], n[sel - 1]]; setLevels(n); setSel(sel - 1); }}><ChevronUp size={13} /></button>
        <button className="xl-btn !min-w-0" disabled={sel >= levels.length - 1} onClick={() => { const n = [...levels]; [n[sel + 1], n[sel]] = [n[sel], n[sel + 1]]; setLevels(n); setSel(sel + 1); }}><ChevronDown size={13} /></button>
        <button className="xl-btn !min-w-0" onClick={() => setShowOpts(!showOpts)}>Options...</button>
        <div className="flex-1" />
        <label className="flex items-center gap-1.5"><input type="checkbox" checked={header} disabled={ltr} onChange={(e) => setHeader(e.target.checked)} data-testid="sort-header" /> My data has headers</label>
      </div>
      {showOpts && (
        <div className="flex gap-4 mb-2 p-2 border rounded" style={{ borderColor: 'var(--border)' }}>
          <label className="flex gap-1.5"><input type="checkbox" checked={caseSensitive} onChange={(e) => setCase(e.target.checked)} /> Case sensitive</label>
          <label className="flex gap-1.5"><input type="radio" checked={!ltr} onChange={() => setLtr(false)} /> Sort top to bottom</label>
          <label className="flex gap-1.5"><input type="radio" checked={ltr} onChange={() => { setLtr(true); setHeader(false); setLevels(levels.map((l) => ({ ...l, col: rg.r1 }))); }} /> Sort left to right</label>
        </div>
      )}
      <div className="grid grid-cols-[80px_1fr_1fr_1fr] gap-x-2 font-semibold mb-1">
        <span />
        <span>{ltr ? 'Row' : 'Column'}</span>
        <span>Sort On</span>
        <span>Order</span>
      </div>
      <div className="min-h-[140px]">
        {levels.map((l, i) => (
          <div key={l.key} className={'grid grid-cols-[80px_1fr_1fr_1fr] gap-x-2 items-center py-0.5 px-0.5 rounded ' + (sel === i ? 'xl-item-active' : '')} onClick={() => setSel(i)}>
            <span>{i === 0 ? 'Sort by' : 'Then by'}</span>
            <select className="xl-input" value={l.col} onChange={(e) => upd(i, { col: +e.target.value })} data-testid={`sort-col-${i}`}>
              {keys.map((k) => (
                <option key={k} value={k}>{ltr ? rowLabel(k) : colLabel(k)}</option>
              ))}
            </select>
            <select className="xl-input" value={l.by} onChange={(e) => upd(i, { by: e.target.value as Level['by'] })}>
              <option value="value">Cell Values</option>
              <option value="cellColor">Cell Color</option>
              <option value="fontColor">Font Color</option>
            </select>
            {l.by === 'value' ? (
              <select
                className="xl-input"
                value={l.customList ? 'custom:' + lists.findIndex((x) => x === l.customList) : l.desc ? 'desc' : 'asc'}
                onChange={(e) => {
                  const v = e.target.value;
                  if (v.startsWith('custom:')) upd(i, { customList: lists[+v.slice(7)], desc: false });
                  else upd(i, { desc: v === 'desc', customList: undefined });
                }}
                data-testid={`sort-order-${i}`}
              >
                <option value="asc">A to Z / Smallest to Largest</option>
                <option value="desc">Z to A / Largest to Smallest</option>
                {lists.map((lst, k) => (
                  <option key={k} value={'custom:' + k}>{lst.slice(0, 3).join(', ')}...</option>
                ))}
              </select>
            ) : (
              <input type="color" value={l.color ?? '#FFFF00'} onChange={(e) => upd(i, { color: e.target.value.toUpperCase() })} />
            )}
          </div>
        ))}
      </div>
    </Dialog>
  );
}

const OPS: [FilterOp, string][] = [
  ['eq', 'equals'],
  ['ne', 'does not equal'],
  ['gt', 'is greater than'],
  ['gte', 'is greater than or equal to'],
  ['lt', 'is less than'],
  ['lte', 'is less than or equal to'],
  ['beginsWith', 'begins with'],
  ['endsWith', 'ends with'],
  ['contains', 'contains'],
  ['notContains', 'does not contain'],
];

export function CustomFilterDialog({ props }: { props: Record<string, unknown> }) {
  const col = props.col as number;
  const sheet = S().wb.activeSheet;
  const [op1, setOp1] = useState<FilterOp>((props.op as FilterOp) ?? 'eq');
  const [v1, setV1] = useState('');
  const [and, setAnd] = useState(true);
  const [op2, setOp2] = useState<FilterOp | ''>((props.op2 as FilterOp) ?? '');
  const [v2, setV2] = useState('');
  const header = sheet.autoFilter ? displayTextAt(sheet, sheet.autoFilter.range.r1, col) : '';
  return (
    <Dialog
      title="Custom AutoFilter"
      width={460}
      onOk={() => setColumnFilter(col, { type: 'custom', and, c1: { op: op1, val: v1 }, c2: op2 ? { op: op2, val: v2 } : undefined })}
    >
      <div className="mb-2">Show rows where: <b>{header}</b></div>
      <div className="flex gap-2 mb-2">
        <select className="xl-input w-[190px]" value={op1} onChange={(e) => setOp1(e.target.value as FilterOp)}>{OPS.map(([k, l]) => <option key={k} value={k}>{l}</option>)}</select>
        <input className="xl-input flex-1" value={v1} onChange={(e) => setV1(e.target.value)} data-testid="cf-v1" />
      </div>
      <div className="flex gap-4 mb-2 pl-6">
        <label className="flex gap-1"><input type="radio" checked={and} onChange={() => setAnd(true)} /> And</label>
        <label className="flex gap-1"><input type="radio" checked={!and} onChange={() => setAnd(false)} /> Or</label>
      </div>
      <div className="flex gap-2">
        <select className="xl-input w-[190px]" value={op2} onChange={(e) => setOp2(e.target.value as FilterOp)}><option value="" />{OPS.map(([k, l]) => <option key={k} value={k}>{l}</option>)}</select>
        <input className="xl-input flex-1" value={v2} onChange={(e) => setV2(e.target.value)} />
      </div>
      <div className="text-[11px] opacity-70 mt-3">Use ? to represent any single character. Use * to represent any series of characters.</div>
    </Dialog>
  );
}

export function Top10Dialog({ props }: { props: Record<string, unknown> }) {
  const [top, setTop] = useState(true);
  const [n, setN] = useState(10);
  const [pct, setPct] = useState(false);
  return (
    <Dialog title="Top 10 AutoFilter" width={360} onOk={() => setColumnFilter(props.col as number, { type: 'top10', top, n, percent: pct })}>
      <div className="flex gap-2 items-center py-2">
        <select className="xl-input" value={top ? 'top' : 'bottom'} onChange={(e) => setTop(e.target.value === 'top')}><option value="top">Top</option><option value="bottom">Bottom</option></select>
        <input type="number" className="xl-input w-16" value={n} min={1} onChange={(e) => setN(Math.max(1, +e.target.value))} />
        <select className="xl-input" value={pct ? 'pct' : 'items'} onChange={(e) => setPct(e.target.value === 'pct')}><option value="items">Items</option><option value="pct">Percent</option></select>
      </div>
    </Dialog>
  );
}

export function PasteSpecialDialog() {
  const [o, setO] = useState<PasteOptions>({ ...DEFAULT_PASTE });
  const whats: [PasteWhat, string][] = [
    ['all', 'All'],
    ['formulas', 'Formulas'],
    ['values', 'Values'],
    ['formats', 'Formats'],
    ['comments', 'Comments and Notes'],
    ['validation', 'Validation'],
    ['allExceptBorders', 'All except borders'],
    ['columnWidths', 'Column widths'],
    ['formulasAndNumberFormats', 'Formulas and number formats'],
    ['valuesAndNumberFormats', 'Values and number formats'],
  ];
  return (
    <Dialog
      title="Paste Special"
      width={440}
      onOk={() => {
        closeDialog();
        pasteFromSystem(o);
      }}
      extraButtons={<button className="xl-btn" onClick={() => { closeDialog(); pasteFromSystem({ ...o, link: true }); }}>Paste Link</button>}
      testId="paste-special"
    >
      <fieldset className="xl-fieldset">
        <legend>Paste</legend>
        <div className="grid grid-cols-2 gap-x-4">
          {whats.map(([k, l]) => (
            <label key={k} className="flex gap-2 my-[3px]"><input type="radio" checked={o.what === k} onChange={() => setO({ ...o, what: k })} /> {l}</label>
          ))}
        </div>
      </fieldset>
      <fieldset className="xl-fieldset mt-2">
        <legend>Operation</legend>
        <div className="grid grid-cols-2 gap-x-4">
          {(['none', 'add', 'subtract', 'multiply', 'divide'] as const).map((k) => (
            <label key={k} className="flex gap-2 my-[3px]"><input type="radio" checked={o.operation === k} onChange={() => setO({ ...o, operation: k })} /> {k[0].toUpperCase() + k.slice(1)}</label>
          ))}
        </div>
      </fieldset>
      <div className="flex gap-6 mt-2">
        <label className="flex gap-2"><input type="checkbox" checked={o.skipBlanks} onChange={(e) => setO({ ...o, skipBlanks: e.target.checked })} /> Skip blanks</label>
        <label className="flex gap-2"><input type="checkbox" checked={o.transpose} onChange={(e) => setO({ ...o, transpose: e.target.checked })} data-testid="ps-transpose" /> Transpose</label>
      </div>
    </Dialog>
  );
}

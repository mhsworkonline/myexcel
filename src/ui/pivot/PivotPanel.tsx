'use client';
import { ChevronDown, Columns3, Filter, RefreshCw, Rows3, Sigma, X, Search } from 'lucide-react';
import { useMemo, useState } from 'react';
import { parseRange, rangeToA1 } from '../../model/address';
import { primaryRange } from '../../model/selection';
import type { PivotAgg, PivotSpec } from '../../model/types';
import { dataRangeForCommand } from '../../state/actions/sortFilter';
import { AGG_LABEL, createPivot, deletePivot, fieldItems, findPivot, pivotFields, refreshPivot, updatePivot } from '../../state/pivot';
import { alertBox, closeDialog, openDialog, S, setState, useStore } from '../../state/store';
import { Dialog, RefInput } from '../dialogs/Dialog';
import { Menu, MenuItem } from '../Menu';

type Area = 'filters' | 'cols' | 'rows' | 'values';

export function PivotPanel() {
  const id = useStore((s) => s.pivotPanel);
  useStore((s) => s.rev);
  const [q, setQ] = useState('');
  const [menu, setMenu] = useState<{ x: number; y: number; items: MenuItem[] } | null>(null);
  const [dragging, setDragging] = useState<{ field: string; from: Area | 'list' } | null>(null);
  const found = id ? findPivot(id) : null;
  const spec = found?.spec;
  const fields = useMemo(() => (spec ? pivotFields(S().wb, spec) : []), [spec]);
  if (!spec) return null;
  const allNames = [...fields.map((f) => f.name), ...spec.calculatedFields.map((c) => c.name)];
  const inUse = new Set([...spec.rows, ...spec.cols, ...spec.values.map((v) => v.field), ...spec.filters.map((f) => f.field)]);
  const update = (patch: Partial<PivotSpec>) => updatePivot(spec.id, patch);

  const remove = (field: string, area: Area, s: PivotSpec = spec): Partial<PivotSpec> => {
    if (area === 'rows') return { rows: s.rows.filter((x) => x !== field) };
    if (area === 'cols') return { cols: s.cols.filter((x) => x !== field) };
    if (area === 'filters') return { filters: s.filters.filter((x) => x.field !== field) };
    return { values: s.values.filter((x) => x.field !== field) };
  };
  const add = (field: string, area: Area, s: PivotSpec = spec): Partial<PivotSpec> => {
    const numeric = fields.find((f) => f.name === field)?.numeric || spec.calculatedFields.some((c) => c.name === field);
    if (area === 'rows') return { rows: [...s.rows.filter((x) => x !== field), field] };
    if (area === 'cols') return { cols: [...s.cols.filter((x) => x !== field), field] };
    if (area === 'filters') return { filters: [...s.filters.filter((x) => x.field !== field), { field, selected: null }] };
    return { values: [...s.values, { field, agg: numeric ? 'sum' : 'count' }] };
  };
  const toggleField = (field: string) => {
    if (inUse.has(field)) {
      let s = { ...spec };
      for (const a of ['rows', 'cols', 'values', 'filters'] as Area[]) s = { ...s, ...remove(field, a, s) };
      update({ rows: s.rows, cols: s.cols, values: s.values, filters: s.filters });
      return;
    }
    const f = fields.find((x) => x.name === field);
    const numeric = !!f?.numeric && !f.date || spec.calculatedFields.some((c) => c.name === field);
    update(add(field, numeric ? 'values' : 'rows'));
  };
  const onDrop = (to: Area) => {
    if (!dragging) return;
    let s = { ...spec };
    if (dragging.from !== 'list') s = { ...s, ...remove(dragging.field, dragging.from, s) };
    s = { ...s, ...add(dragging.field, to, s) };
    update({ rows: s.rows, cols: s.cols, values: s.values, filters: s.filters });
    setDragging(null);
  };
  const itemMenu = (field: string, area: Area, e: React.MouseEvent) => {
    const r = (e.currentTarget as HTMLElement).getBoundingClientRect();
    const list = area === 'rows' ? spec.rows : area === 'cols' ? spec.cols : area === 'values' ? spec.values.map((v) => v.field) : spec.filters.map((f) => f.field);
    const i = list.indexOf(field);
    const move = (d: number) => {
      const n = [...list];
      [n[i], n[i + d]] = [n[i + d], n[i]];
      if (area === 'rows') update({ rows: n });
      else if (area === 'cols') update({ cols: n });
      else if (area === 'values') update({ values: n.map((f) => spec.values.find((v) => v.field === f)!) });
      else update({ filters: n.map((f) => spec.filters.find((v) => v.field === f)!) });
    };
    const items: MenuItem[] = [
      { label: 'Move Up', disabled: i <= 0, onClick: () => move(-1) },
      { label: 'Move Down', disabled: i >= list.length - 1, onClick: () => move(1) },
      { separator: true },
      ...(area !== 'rows' ? [{ label: 'Move to Row Labels', onClick: () => { let s = { ...spec, ...remove(field, area) }; s = { ...s, ...add(field, 'rows', s) }; update(s); } }] : []),
      ...(area !== 'cols' ? [{ label: 'Move to Column Labels', onClick: () => { let s = { ...spec, ...remove(field, area) }; s = { ...s, ...add(field, 'cols', s) }; update(s); } }] : []),
      ...(area !== 'filters' ? [{ label: 'Move to Report Filter', onClick: () => { let s = { ...spec, ...remove(field, area) }; s = { ...s, ...add(field, 'filters', s) }; update(s); } }] : []),
      { separator: true },
      { label: 'Remove Field', onClick: () => update(remove(field, area)) },
    ];
    if (area === 'values') items.push({ separator: true }, { label: 'Value Field Settings...', onClick: () => openDialog('pivotValueSettings', { id: spec.id, field }) });
    if (area === 'rows' || area === 'cols') {
      const f = fields.find((x) => x.name === field);
      if (f?.numeric) items.push({ separator: true }, { label: 'Group...', onClick: () => openDialog('pivotGroup', { id: spec.id, field, date: f.date }) }, { label: 'Ungroup', disabled: !spec.grouping.some((g) => g.field === field), onClick: () => update({ grouping: spec.grouping.filter((g) => g.field !== field) }) });
    }
    if (area === 'filters') items.push({ separator: true }, { label: 'Select Items...', onClick: () => openDialog('pivotFilter', { id: spec.id, field }) });
    setMenu({ x: r.left, y: r.bottom, items });
  };
  const box = (area: Area, title: string, icon: React.ReactNode, list: { field: string; label: string }[]) => (
    <div className="flex flex-col min-h-0">
      <div className="flex items-center gap-1 text-[11px] font-semibold mb-0.5 opacity-80">{icon}{title}</div>
      <div
        className="flex-1 min-h-[70px] border rounded p-1 overflow-auto"
        style={{ borderColor: dragging ? 'var(--accent)' : 'var(--border)', background: 'var(--input-bg)' }}
        onDragOver={(e) => e.preventDefault()}
        onDrop={() => onDrop(area)}
        data-testid={`pv-area-${area}`}
      >
        {list.map((it) => (
          <div
            key={it.field}
            draggable
            onDragStart={() => setDragging({ field: it.field, from: area })}
            onDragEnd={() => setDragging(null)}
            className="flex items-center justify-between px-1.5 py-0.5 mb-0.5 rounded border cursor-move text-[12px]"
            style={{ borderColor: 'var(--border)', background: 'var(--panel)' }}
          >
            <span className="truncate">{it.label}</span>
            <button className="opacity-70 hover:opacity-100" onClick={(e) => itemMenu(it.field, area, e)} aria-label={`${it.label} options`}><ChevronDown size={12} /></button>
          </div>
        ))}
      </div>
    </div>
  );
  return (
    <div className="w-[300px] shrink-0 flex flex-col border-l text-[12px]" style={{ borderColor: 'var(--border)', background: 'var(--panel)' }} data-testid="pivot-panel">
      <div className="flex items-center px-3 h-9">
        <span className="font-semibold text-[14px] flex-1">PivotTable Fields</span>
        <button className="w-6 h-6 rounded hover:bg-[var(--hover)] flex items-center justify-center" onClick={() => setState({ pivotPanel: null })} aria-label="Close"><X size={14} /></button>
      </div>
      <div className="px-3 opacity-80">Choose fields to add to report:</div>
      <div className="mx-3 mt-1 flex items-center border rounded px-1.5 h-6" style={{ borderColor: 'var(--border-strong)', background: 'var(--input-bg)' }}>
        <input className="flex-1 bg-transparent outline-none" placeholder="Search" value={q} onChange={(e) => setQ(e.target.value)} />
        <Search size={12} className="opacity-60" />
      </div>
      <div className="mx-3 mt-1 border rounded h-[170px] overflow-auto p-1" style={{ borderColor: 'var(--border)' }}>
        {allNames.filter((n) => !q || n.toLowerCase().includes(q.toLowerCase())).map((n) => (
          <label key={n} className="flex items-center gap-1.5 py-[2px] cursor-default" draggable onDragStart={() => setDragging({ field: n, from: 'list' })} onDragEnd={() => setDragging(null)}>
            <input type="checkbox" checked={inUse.has(n)} onChange={() => toggleField(n)} data-testid={`pv-field-${n}`} />
            <span className={inUse.has(n) ? 'font-semibold' : ''}>{n}</span>
          </label>
        ))}
      </div>
      <div className="px-3 mt-2 opacity-80">Drag fields between areas below:</div>
      <div className="grid grid-cols-2 gap-2 p-3 flex-1 min-h-0">
        {box('filters', 'Filters', <Filter size={11} />, spec.filters.map((f) => ({ field: f.field, label: f.field + (f.selected ? ' (filtered)' : '') })))}
        {box('cols', 'Columns', <Columns3 size={11} />, spec.cols.map((f) => ({ field: f, label: f })))}
        {box('rows', 'Rows', <Rows3 size={11} />, spec.rows.map((f) => ({ field: f, label: f })))}
        {box('values', 'Values', <Sigma size={11} />, spec.values.map((v) => ({ field: v.field, label: v.name ?? `${AGG_LABEL[v.agg]} of ${v.field}` })))}
      </div>
      <div className="flex items-center gap-1.5 px-3 pb-3 flex-wrap">
        <button className="xl-btn !min-w-0 flex items-center gap-1" onClick={() => refreshPivot(spec.id)} data-testid="pv-refresh"><RefreshCw size={12} /> Refresh</button>
        <button className="xl-btn !min-w-0" onClick={() => openDialog('pivotCalcField', { id: spec.id })}>Calculated Field...</button>
        <label className="flex items-center gap-1"><input type="checkbox" checked={spec.showGrandTotals} onChange={(e) => update({ showGrandTotals: e.target.checked })} /> Grand totals</label>
        <button className="xl-btn !min-w-0" onClick={() => deletePivot(spec.id)}>Delete</button>
      </div>
      {menu && <Menu x={menu.x} y={menu.y} items={menu.items} onClose={() => setMenu(null)} />}
    </div>
  );
}

export function CreatePivotDialog({ props }: { props: Record<string, unknown> }) {
  const st = S();
  const sheet = st.wb.activeSheet;
  const guess = (props.range as ReturnType<typeof dataRangeForCommand>) ?? dataRangeForCommand() ?? primaryRange(st.sel);
  const [ref, setRef] = useState(`=${sheet.name.includes(' ') ? `'${sheet.name}'` : sheet.name}!${rangeToA1(guess, true)}`);
  const [where, setWhere] = useState<'new' | 'existing'>('new');
  const [loc, setLoc] = useState('');
  return (
    <Dialog
      title="Create PivotTable"
      width={420}
      onOk={() => {
        const txt = ref.replace(/^=/, '');
        const bang = txt.lastIndexOf('!');
        const srcSheet = bang > 0 ? S().wb.sheetByName(txt.slice(0, bang).replace(/^'|'$/g, '')) : sheet;
        const rg = parseRange(txt.slice(bang + 1).replace(/\$/g, ''));
        if (!srcSheet || !rg || rg.r2 <= rg.r1) {
          alertBox('The PivotTable field name is not valid. To create a PivotTable report, you must use data that is organized as a list with labeled columns.');
          return false;
        }
        if (where === 'existing') {
          const lt = loc.replace(/^=/, '');
          const lb = lt.lastIndexOf('!');
          const ls = lb > 0 ? S().wb.sheetByName(lt.slice(0, lb).replace(/^'|'$/g, '')) : sheet;
          const lr = parseRange(lt.slice(lb + 1).replace(/\$/g, ''));
          if (!ls || !lr) {
            alertBox('The destination reference is not valid.');
            return false;
          }
          createPivot(srcSheet.id, rg, { sheetId: ls.id, r: lr.r1, c: lr.c1 });
        } else createPivot(srcSheet.id, rg, { sheetId: null, r: 2, c: 0 });
      }}
      testId="create-pivot"
    >
      <div className="font-semibold mb-1">Choose the data that you want to analyze</div>
      <label className="flex items-center gap-2 mb-3">Table/Range: <RefInput value={ref} onChange={setRef} width={260} /></label>
      <div className="font-semibold mb-1">Choose where you want the PivotTable report to be placed</div>
      <label className="flex gap-2 my-1"><input type="radio" checked={where === 'new'} onChange={() => setWhere('new')} /> New Worksheet</label>
      <label className="flex gap-2 my-1"><input type="radio" checked={where === 'existing'} onChange={() => setWhere('existing')} /> Existing Worksheet</label>
      {where === 'existing' && <label className="flex items-center gap-2 ml-6">Location: <RefInput value={loc} onChange={setLoc} width={220} /></label>}
    </Dialog>
  );
}

export function PivotValueSettingsDialog({ props }: { props: Record<string, unknown> }) {
  const found = findPivot(props.id as string);
  const field = props.field as string;
  const vd = found?.spec.values.find((v) => v.field === field);
  const [agg, setAgg] = useState<PivotAgg>(vd?.agg ?? 'sum');
  const [name, setName] = useState(vd?.name ?? '');
  if (!found || !vd) return null;
  return (
    <Dialog
      title="Value Field Settings"
      width={340}
      onOk={() => updatePivot(found.spec.id, { values: found.spec.values.map((v) => (v.field === field ? { ...v, agg, name: name || undefined } : v)) })}
    >
      <div className="mb-2">Source Name: <b>{field}</b></div>
      <label className="flex items-center gap-2 mb-2">Custom Name: <input className="xl-input flex-1" value={name} placeholder={`${AGG_LABEL[agg]} of ${field}`} onChange={(e) => setName(e.target.value)} /></label>
      <div className="mb-1">Summarize value field by:</div>
      <div className="xl-list h-[150px]">
        {(Object.keys(AGG_LABEL) as PivotAgg[]).map((a) => (
          <div key={a} className={agg === a ? 'xl-list-sel' : ''} onClick={() => setAgg(a)}>{AGG_LABEL[a]}</div>
        ))}
      </div>
    </Dialog>
  );
}

export function PivotGroupDialog({ props }: { props: Record<string, unknown> }) {
  const found = findPivot(props.id as string);
  const field = props.field as string;
  const isDate = !!props.date;
  const cur = found?.spec.grouping.find((g) => g.field === field);
  const [by, setBy] = useState<'years' | 'quarters' | 'months' | 'days'>(cur?.by ?? 'months');
  const [start, setStart] = useState(String(cur?.start ?? 0));
  const [interval, setInterval] = useState(String(cur?.interval ?? 10));
  if (!found) return null;
  return (
    <Dialog
      title="Grouping"
      width={300}
      onOk={() => {
        const g = isDate ? { field, type: 'date' as const, by } : { field, type: 'number' as const, start: +start || 0, interval: Math.max(1e-9, +interval || 10) };
        updatePivot(found.spec.id, { grouping: [...found.spec.grouping.filter((x) => x.field !== field), g] });
      }}
    >
      {isDate ? (
        <>
          <div className="mb-1">By</div>
          <div className="xl-list h-[100px]">
            {(['days', 'months', 'quarters', 'years'] as const).map((b) => (
              <div key={b} className={by === b ? 'xl-list-sel' : ''} onClick={() => setBy(b)}>{b[0].toUpperCase() + b.slice(1)}</div>
            ))}
          </div>
        </>
      ) : (
        <>
          <label className="flex items-center gap-2 my-1">Starting at: <input className="xl-input w-24" value={start} onChange={(e) => setStart(e.target.value)} /></label>
          <label className="flex items-center gap-2 my-1">By: <input className="xl-input w-24" value={interval} onChange={(e) => setInterval(e.target.value)} /></label>
        </>
      )}
    </Dialog>
  );
}

export function PivotCalcFieldDialog({ props }: { props: Record<string, unknown> }) {
  const found = findPivot(props.id as string);
  const [name, setName] = useState('Field1');
  const [formula, setFormula] = useState('= 0');
  if (!found) return null;
  const fields = pivotFields(S().wb, found.spec);
  return (
    <Dialog
      title="Insert Calculated Field"
      width={400}
      okLabel="Add"
      onOk={() => {
        if (!name.trim()) return false;
        const cf = [...found.spec.calculatedFields.filter((c) => c.name !== name), { name: name.trim(), formula }];
        updatePivot(found.spec.id, { calculatedFields: cf, values: [...found.spec.values.filter((v) => v.field !== name), { field: name.trim(), agg: 'sum' }] });
      }}
    >
      <label className="flex items-center gap-2 my-1"><span className="w-16">Name:</span><input className="xl-input flex-1" value={name} onChange={(e) => setName(e.target.value)} /></label>
      <label className="flex items-center gap-2 my-1"><span className="w-16">Formula:</span><input className="xl-input flex-1" value={formula} onChange={(e) => setFormula(e.target.value)} data-testid="calc-formula" /></label>
      <div className="mt-2 mb-1">Fields:</div>
      <div className="xl-list h-[120px]">
        {fields.map((f) => (
          <div key={f.name} onDoubleClick={() => setFormula(formula + (/\s/.test(f.name) ? `'${f.name}'` : f.name))}>{f.name}</div>
        ))}
      </div>
      <div className="text-[11px] opacity-70 mt-1">Double-click a field to insert it. Calculated fields operate on the sums of the underlying fields.</div>
    </Dialog>
  );
}

export function PivotFilterDialog({ props }: { props: Record<string, unknown> }) {
  const found = findPivot(props.id as string);
  const field = props.field as string;
  const items = useMemo(() => (found ? fieldItems(found.spec, field) : []), [found, field]);
  const cur = found?.spec.filters.find((f) => f.field === field);
  const [sel, setSel] = useState<Set<string>>(new Set(cur?.selected ?? items));
  if (!found) return null;
  return (
    <Dialog
      title={field}
      width={280}
      onOk={() => {
        const all = sel.size === items.length;
        updatePivot(found.spec.id, { filters: found.spec.filters.map((f) => (f.field === field ? { ...f, selected: all ? null : [...sel] } : f)) });
      }}
    >
      <label className="flex gap-1.5 mb-1"><input type="checkbox" checked={sel.size === items.length} onChange={(e) => setSel(e.target.checked ? new Set(items) : new Set())} /> (All)</label>
      <div className="xl-list h-[180px]">
        {items.map((it) => (
          <label key={it} className="flex gap-1.5 px-1">
            <input type="checkbox" checked={sel.has(it)} onChange={() => { const n = new Set(sel); if (n.has(it)) n.delete(it); else n.add(it); setSel(n); }} /> {it}
          </label>
        ))}
      </div>
    </Dialog>
  );
}

export { closeDialog };

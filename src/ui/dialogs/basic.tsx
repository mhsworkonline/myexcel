'use client';
import { AlertTriangle, CircleX, Info } from 'lucide-react';
import { useState } from 'react';
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT } from '../../model/sheet';
import { primaryRange } from '../../model/selection';
import { applyCommittedText, commitEdit } from '../../state/actions/edit';
import { clearAutosave, recoverAutosave, saveAs, SaveFormat } from '../../state/actions/file';
import { fillSeries } from '../../state/actions/fill';
import { goTo, goToSpecial, SpecialKind } from '../../state/actions/find';
import { selectedCols, selectedRows, setColumnWidth, setRowHeight } from '../../state/actions/format';
import { copySheet, deleteCells, insertCells, moveSheet, ShiftDir, unhideSheet } from '../../state/actions/structure';
import { bump, closeDialog, S, setState, useStore } from '../../state/store';
import { Dialog, Field } from './Dialog';

type P = Record<string, unknown>;

export function AlertDialog({ props }: { props: P }) {
  const icon = props.icon === 'error' ? <CircleX className="text-[#C42B1C] shrink-0" size={32} /> : props.icon === 'info' ? <Info className="text-[#2B7CD3] shrink-0" size={32} /> : <AlertTriangle className="text-[#E8A200] shrink-0" size={32} />;
  return (
    <Dialog title={(props.title as string) ?? 'MyExcel'} width={440} cancelLabel={null} testId="alert-dialog">
      <div className="flex gap-3 py-2">
        {icon}
        <div className="whitespace-pre-wrap leading-[18px]">{props.message as string}</div>
      </div>
    </Dialog>
  );
}

export function ConfirmDialog({ props }: { props: P }) {
  return (
    <Dialog title={(props.title as string) ?? 'MyExcel'} width={440} okLabel={(props.okLabel as string) ?? 'OK'} onOk={() => { closeDialog(); (props.onOk as () => void)?.(); return false; }} onCancel={() => (props.onCancel as () => void)?.()} testId="confirm-dialog">
      <div className="flex gap-3 py-2">
        <AlertTriangle className="text-[#E8A200] shrink-0" size={32} />
        <div className="whitespace-pre-wrap leading-[18px]">{props.message as string}</div>
      </div>
    </Dialog>
  );
}

export function ValidationErrorDialog({ props }: { props: P }) {
  const style = props.style as 'stop' | 'warning' | 'information';
  const accept = () => {
    const st = S();
    const ed = st.edit;
    if (!ed) return;
    const sheet = st.wb.sheetById(ed.hostSheetId)!;
    applyCommittedText(sheet, ed.r, ed.c, props.text as string, st.sel, { fillSelection: !!props.fillSelection });
    setState({ edit: null });
  };
  const footer =
    style === 'stop' ? (
      <div className="flex justify-end gap-2 px-4 py-3">
        <button className="xl-btn xl-btn-primary" onClick={() => closeDialog()} data-testid="dv-retry">Retry</button>
        <button className="xl-btn" onClick={() => { closeDialog(); import('../../state/actions/edit').then((m) => m.cancelEdit()); }}>Cancel</button>
      </div>
    ) : style === 'warning' ? (
      <div className="flex justify-end gap-2 px-4 py-3">
        <span className="flex-1 self-center">Continue?</span>
        <button className="xl-btn xl-btn-primary" onClick={() => { closeDialog(); accept(); }}>Yes</button>
        <button className="xl-btn" onClick={() => closeDialog()}>No</button>
        <button className="xl-btn" onClick={() => { closeDialog(); import('../../state/actions/edit').then((m) => m.cancelEdit()); }}>Cancel</button>
      </div>
    ) : (
      <div className="flex justify-end gap-2 px-4 py-3">
        <button className="xl-btn xl-btn-primary" onClick={() => { closeDialog(); accept(); }}>OK</button>
        <button className="xl-btn" onClick={() => { closeDialog(); import('../../state/actions/edit').then((m) => m.cancelEdit()); }}>Cancel</button>
      </div>
    );
  const icon = style === 'stop' ? <CircleX className="text-[#C42B1C] shrink-0" size={32} /> : style === 'warning' ? <AlertTriangle className="text-[#E8A200] shrink-0" size={32} /> : <Info className="text-[#2B7CD3] shrink-0" size={32} />;
  return (
    <Dialog title={props.title as string} width={420} footer={footer} testId="validation-error">
      <div className="flex gap-3 py-2">
        {icon}
        <div className="whitespace-pre-wrap">{props.message as string}</div>
      </div>
    </Dialog>
  );
}

export function InsertDeleteCellsDialog({ mode }: { mode: 'insert' | 'delete' }) {
  const [dir, setDir] = useState<ShiftDir>(mode === 'insert' ? 'down' : 'up');
  const opts: [ShiftDir, string][] =
    mode === 'insert'
      ? [['right', 'Shift cells right'], ['down', 'Shift cells down'], ['row', 'Entire row'], ['col', 'Entire column']]
      : [['left', 'Shift cells left'], ['up', 'Shift cells up'], ['row', 'Entire row'], ['col', 'Entire column']];
  return (
    <Dialog title={mode === 'insert' ? 'Insert' : 'Delete'} width={240} onOk={() => (mode === 'insert' ? insertCells(dir) : deleteCells(dir))}>
      <fieldset className="xl-fieldset">
        <legend>{mode === 'insert' ? 'Insert' : 'Delete'}</legend>
        {opts.map(([k, l]) => (
          <label key={k} className="flex items-center gap-2 my-1">
            <input type="radio" checked={dir === k} onChange={() => setDir(k)} /> {l}
          </label>
        ))}
      </fieldset>
    </Dialog>
  );
}

export function SizeDialog({ axis, props }: { axis: 'row' | 'col'; props: P }) {
  const sheet = S().wb.activeSheet;
  const a = S().sel.active;
  const standard = !!props.standard;
  const cur = axis === 'row' ? sheet.rowHeight(a.r) : sheet.colWidth(a.c);
  const init = axis === 'row' ? (cur * 0.75).toFixed(2).replace(/\.?0+$/, '') : ((standard ? DEFAULT_COL_WIDTH : cur) - 5) / 7;
  const [v, setV] = useState(String(typeof init === 'number' ? +init.toFixed(2) : init));
  return (
    <Dialog
      title={axis === 'row' ? 'Row Height' : standard ? 'Standard Width' : 'Column Width'}
      width={260}
      onOk={() => {
        const n = parseFloat(v);
        if (isNaN(n) || n < 0 || (axis === 'row' ? n > 409 : n > 255)) {
          import('../../state/store').then((m) => m.alertBox(axis === 'row' ? 'Row height must be between 0 and 409.' : 'Column width must be between 0 and 255 characters.'));
          return false;
        }
        if (axis === 'row') setRowHeight(Math.round(n / 0.75), selectedRows());
        else setColumnWidth(n === 0 ? 0 : Math.round(n * 7 + 5), standard ? undefined : selectedCols());
      }}
    >
      <Field label={axis === 'row' ? 'Row height:' : standard ? 'Standard column width:' : 'Column width:'}>
        <input className="xl-input w-24" value={v} onChange={(e) => setV(e.target.value)} data-testid="size-input" />
      </Field>
      <div className="text-[11px] opacity-60">Default: {axis === 'row' ? `${DEFAULT_ROW_HEIGHT * 0.75} pt` : '8.43 characters'}</div>
    </Dialog>
  );
}

export function ZoomDialog() {
  const sheet = S().wb.activeSheet;
  const [z, setZ] = useState(String(sheet.zoom));
  const presets = [200, 100, 75, 50, 25];
  return (
    <Dialog
      title="Zoom"
      width={240}
      onOk={() => {
        let n = parseInt(z, 10);
        if (z === 'fit') {
          const rg = primaryRange(S().sel);
          n = Math.floor(Math.min(1200 / ((rg.c2 - rg.c1 + 1) * 64), 600 / ((rg.r2 - rg.r1 + 1) * 20)) * 100);
        }
        sheet.zoom = Math.max(10, Math.min(400, n || 100));
        sheet.touch();
        bump();
      }}
    >
      <fieldset className="xl-fieldset">
        <legend>Magnification</legend>
        {presets.map((p) => (
          <label key={p} className="flex items-center gap-2 my-0.5">
            <input type="radio" checked={z === String(p)} onChange={() => setZ(String(p))} /> {p}%
          </label>
        ))}
        <label className="flex items-center gap-2 my-0.5">
          <input type="radio" checked={z === 'fit'} onChange={() => setZ('fit')} /> Fit selection
        </label>
        <label className="flex items-center gap-2 my-0.5">
          <input type="radio" checked={!presets.map(String).includes(z) && z !== 'fit'} readOnly /> Custom:
          <input className="xl-input w-16" value={presets.map(String).includes(z) || z === 'fit' ? '' : z} onChange={(e) => setZ(e.target.value)} /> %
        </label>
      </fieldset>
    </Dialog>
  );
}

export function SeriesDialog() {
  const rg = primaryRange(S().sel);
  const [rowsOrCols, setRC] = useState<'rows' | 'columns'>(rg.r2 - rg.r1 >= rg.c2 - rg.c1 ? 'columns' : 'rows');
  const [type, setType] = useState<'linear' | 'growth' | 'date' | 'autofill'>('linear');
  const [unit, setUnit] = useState<'day' | 'weekday' | 'month' | 'year'>('day');
  const [step, setStep] = useState('1');
  const [stop, setStop] = useState('');
  return (
    <Dialog title="Series" width={400} onOk={() => fillSeries({ rowsOrCols, type, dateUnit: unit, step: parseFloat(step) || 1, stop: stop ? parseFloat(stop) : undefined, trend: false })}>
      <div className="flex gap-3">
        <fieldset className="xl-fieldset flex-1">
          <legend>Series in</legend>
          <label className="flex gap-2 my-1"><input type="radio" checked={rowsOrCols === 'rows'} onChange={() => setRC('rows')} /> Rows</label>
          <label className="flex gap-2 my-1"><input type="radio" checked={rowsOrCols === 'columns'} onChange={() => setRC('columns')} /> Columns</label>
        </fieldset>
        <fieldset className="xl-fieldset flex-1">
          <legend>Type</legend>
          {(['linear', 'growth', 'date', 'autofill'] as const).map((t) => (
            <label key={t} className="flex gap-2 my-1"><input type="radio" checked={type === t} onChange={() => setType(t)} /> {t === 'autofill' ? 'AutoFill' : t[0].toUpperCase() + t.slice(1)}</label>
          ))}
        </fieldset>
        <fieldset className="xl-fieldset flex-1" disabled={type !== 'date'}>
          <legend>Date unit</legend>
          {(['day', 'weekday', 'month', 'year'] as const).map((t) => (
            <label key={t} className="flex gap-2 my-1"><input type="radio" checked={unit === t} onChange={() => setUnit(t)} /> {t[0].toUpperCase() + t.slice(1)}</label>
          ))}
        </fieldset>
      </div>
      <div className="flex gap-4 mt-3">
        <Field label="Step value:"><input className="xl-input w-20" value={step} onChange={(e) => setStep(e.target.value)} /></Field>
        <Field label="Stop value:"><input className="xl-input w-20" value={stop} onChange={(e) => setStop(e.target.value)} /></Field>
      </div>
    </Dialog>
  );
}

export function MoveCopySheetDialog({ props }: { props: P }) {
  const wb = S().wb;
  const id = (props.sheetId as string) ?? wb.activeSheetId;
  const [before, setBefore] = useState<number>(wb.sheets.length);
  const [copy, setCopy] = useState(false);
  return (
    <Dialog title="Move or Copy" width={320} onOk={() => (copy ? copySheet(id, before) : moveSheet(id, before))}>
      <div className="mb-1">Move selected sheets</div>
      <div className="mb-1">Before sheet:</div>
      <div className="xl-list h-[150px]">
        {wb.sheets.map((s, i) => (
          <div key={s.id} className={before === i ? 'xl-list-sel' : ''} onClick={() => setBefore(i)}>{s.name}</div>
        ))}
        <div className={before === wb.sheets.length ? 'xl-list-sel' : ''} onClick={() => setBefore(wb.sheets.length)}>(move to end)</div>
      </div>
      <label className="flex items-center gap-2 mt-2"><input type="checkbox" checked={copy} onChange={(e) => setCopy(e.target.checked)} /> Create a copy</label>
    </Dialog>
  );
}

export function UnhideSheetDialog() {
  const hidden = S().wb.sheets.filter((s) => s.visibility === 'hidden');
  const [sel, setSel] = useState(hidden[0]?.id);
  return (
    <Dialog title="Unhide" width={300} onOk={() => { if (sel) unhideSheet(sel); }}>
      <div className="mb-1">Unhide sheet:</div>
      <div className="xl-list h-[140px]">
        {hidden.map((s) => (
          <div key={s.id} className={sel === s.id ? 'xl-list-sel' : ''} onClick={() => setSel(s.id)} onDoubleClick={() => { unhideSheet(s.id); closeDialog(); }}>{s.name}</div>
        ))}
        {!hidden.length && <div className="opacity-60">(no hidden sheets)</div>}
      </div>
    </Dialog>
  );
}

export function RecoveryDialog({ props }: { props: P }) {
  return (
    <Dialog
      title="Document Recovery"
      width={420}
      okLabel="Recover"
      cancelLabel="Discard"
      onOk={() => recoverAutosave()}
      onCancel={() => clearAutosave()}
      testId="recovery-dialog"
    >
      <div className="py-2 leading-5">
        MyExcel found an unsaved version of <b>{props.fileName as string}</b> from {new Date(props.savedAt as number).toLocaleString()}.
        <br />
        Do you want to recover it?
      </div>
    </Dialog>
  );
}

export function SaveAsDialog() {
  const [fmt, setFmt] = useState<SaveFormat>('xlsx');
  const file = useStore((s) => s.file);
  return (
    <Dialog title="Save As" width={380} okLabel="Save" onOk={() => { saveAs(fmt); }}>
      <Field label="File name:"><span className="font-semibold">{file.name}</span></Field>
      <Field label="Save as type:">
        <select className="xl-input flex-1" value={fmt} onChange={(e) => setFmt(e.target.value as SaveFormat)} data-testid="saveas-type">
          <option value="xlsx">Excel Workbook (*.xlsx)</option>
          <option value="csv">CSV UTF-8 (Comma delimited) (*.csv)</option>
          <option value="tsv">Text (Tab delimited) (*.tsv)</option>
          <option value="ods">OpenDocument Spreadsheet (*.ods)</option>
          <option value="pdf">PDF (*.pdf)</option>
        </select>
      </Field>
      {(fmt === 'csv' || fmt === 'tsv') && S().wb.sheets.length > 1 && <div className="text-[11px] mt-2 opacity-75">Only the active sheet will be saved in this format.</div>}
    </Dialog>
  );
}

export function GoToDialog() {
  const names = S().wb.names;
  const [ref, setRef] = useState('');
  return (
    <Dialog
      title="Go To"
      width={320}
      onOk={() => {
        if (!ref.trim()) return;
        if (!goTo(ref)) {
          import('../../state/store').then((m) => m.alertBox("Reference isn't valid."));
          return false;
        }
      }}
      extraButtons={<button className="xl-btn" onClick={() => import('../../state/store').then((m) => m.openDialog('gotoSpecial'))}>Special...</button>}
    >
      <div className="mb-1">Go to:</div>
      <div className="xl-list h-[140px] mb-2">
        {names.map((n) => (
          <div key={n.name} className={ref === n.name ? 'xl-list-sel' : ''} onClick={() => setRef(n.name)} onDoubleClick={() => { goTo(n.name); closeDialog(); }}>{n.name}</div>
        ))}
      </div>
      <Field label="Reference:"><input className="xl-input flex-1" value={ref} onChange={(e) => setRef(e.target.value)} data-testid="goto-ref" /></Field>
    </Dialog>
  );
}

export function GoToSpecialDialog() {
  const [kind, setKind] = useState<SpecialKind>('blanks');
  const [types, setTypes] = useState({ numbers: true, text: true, logicals: true, errors: true });
  const items: [SpecialKind, string][] = [
    ['notes', 'Notes'],
    ['constants', 'Constants'],
    ['formulas', 'Formulas'],
    ['blanks', 'Blanks'],
    ['currentRegion', 'Current region'],
    ['rowDiffs', 'Row differences'],
    ['colDiffs', 'Column differences'],
    ['precedents', 'Precedents'],
    ['dependents', 'Dependents'],
    ['lastCell', 'Last cell'],
    ['visible', 'Visible cells only'],
    ['conditionalFormats', 'Conditional formats'],
    ['validation', 'Data validation'],
  ];
  return (
    <Dialog title="Go To Special" width={420} onOk={() => goToSpecial({ kind, ...types })}>
      <fieldset className="xl-fieldset">
        <legend>Select</legend>
        <div className="grid grid-cols-2 gap-x-6">
          {items.map(([k, l]) => (
            <div key={k}>
              <label className="flex gap-2 my-[3px]"><input type="radio" checked={kind === k} onChange={() => setKind(k)} /> {l}</label>
              {k === 'formulas' && (
                <div className="pl-6">
                  {(['numbers', 'text', 'logicals', 'errors'] as const).map((t) => (
                    <label key={t} className={'flex gap-2 my-[2px] ' + (kind === 'constants' || kind === 'formulas' ? '' : 'opacity-40')}>
                      <input type="checkbox" disabled={!(kind === 'constants' || kind === 'formulas')} checked={types[t]} onChange={(e) => setTypes({ ...types, [t]: e.target.checked })} /> {t[0].toUpperCase() + t.slice(1)}
                    </label>
                  ))}
                </div>
              )}
            </div>
          ))}
        </div>
      </fieldset>
    </Dialog>
  );
}

export { commitEdit };

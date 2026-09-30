'use client';
import { useState } from 'react';
import { addrToA1, parseRange } from '../../model/address';
import { primaryRange } from '../../model/selection';
import { newId } from '../../model/sheet';
import type { CellValue } from '../../model/types';
import { selectionRefText } from '../../state/actions/formulas';
import { resolveReference } from '../../state/actions/find';
import { alertBox, closeDialog, openDialog, S, transact, useStore } from '../../state/store';
import { getScalar } from '../../state/values';
import {
  addScenario,
  commitCellValue,
  Constraint,
  dataTable,
  deleteScenario,
  goalSeek,
  keepSolverSolution,
  parseCellRef,
  scenarioSummary,
  showScenario,
  solve,
  SolverSpec,
} from '../../state/whatif';
import { Dialog, RefInput } from './Dialog';

const activeRef = () => '$' + addrToA1(S().sel.active.r, S().sel.active.c).replace(/(\d)/, '$$$1');

export function GoalSeekDialog() {
  const [setCell, setSetCell] = useState(activeRef());
  const [to, setTo] = useState('');
  const [chg, setChg] = useState('');
  return (
    <Dialog
      title="Goal Seek"
      width={320}
      onOk={() => {
        const target = Number(to);
        if (to === '' || isNaN(target)) {
          alertBox('The To value must be a number.');
          return false;
        }
        const res = goalSeek(setCell, target, chg);
        if (typeof res === 'string') {
          alertBox(res);
          return false;
        }
        closeDialog();
        openDialog('goalSeekStatus', { res, chg, setCell });
        return false;
      }}
      testId="goal-seek"
    >
      <label className="flex items-center gap-2 my-1"><span className="w-28">Set cell:</span><RefInput value={setCell} onChange={setSetCell} width={150} testId="gs-set" /></label>
      <label className="flex items-center gap-2 my-1"><span className="w-28">To value:</span><input className="xl-input w-[150px]" value={to} onChange={(e) => setTo(e.target.value)} data-testid="gs-to" /></label>
      <label className="flex items-center gap-2 my-1"><span className="w-28">By changing cell:</span><RefInput value={chg} onChange={setChg} width={150} testId="gs-chg" /></label>
    </Dialog>
  );
}

export function GoalSeekStatusDialog({ props }: { props: Record<string, unknown> }) {
  const res = props.res as { found: boolean; value: number; target: number; achieved: number | null };
  const ref = parseCellRef(props.chg as string);
  const setRef = props.setCell as string;
  return (
    <Dialog
      title="Goal Seek Status"
      width={340}
      onOk={() => {
        if (ref) commitCellValue(ref.sheet, ref.r, ref.c, res.value, 'Goal Seek');
      }}
      testId="goal-seek-status"
    >
      <div className="leading-6">
        Goal Seeking with Cell {setRef.replace(/\$/g, '')}
        <br />
        {res.found ? 'found a solution.' : 'may not have found a solution.'}
        <div className="grid grid-cols-2 mt-2">
          <span>Target value:</span><span>{res.target}</span>
          <span>Current value:</span><span>{res.achieved === null ? '#N/A' : +res.achieved.toPrecision(10)}</span>
        </div>
      </div>
    </Dialog>
  );
}

let lastSolver: SolverSpec | null = null;

export function SolverDialog() {
  const [spec, setSpec] = useState<SolverSpec>(lastSolver ?? { objective: activeRef(), goal: 'max', targetValue: 0, variables: '', constraints: [], nonNegative: true });
  const [newC, setNewC] = useState<Constraint>({ ref: '', op: '<=', value: '' });
  const [sel, setSel] = useState<number | null>(null);
  const upd = (p: Partial<SolverSpec>) => setSpec({ ...spec, ...p });
  return (
    <Dialog
      title="Solver Parameters"
      width={520}
      okLabel="Solve"
      onOk={() => {
        lastSolver = spec;
        const res = solve(spec);
        closeDialog();
        openDialog('solverResults', { res, spec });
        return false;
      }}
      testId="solver"
    >
      <label className="flex items-center gap-2 my-1"><span className="w-28">Set Objective:</span><RefInput value={spec.objective} onChange={(v) => upd({ objective: v })} width={180} testId="solver-obj" /></label>
      <div className="flex items-center gap-4 my-2">
        <span className="w-28">To:</span>
        {(['max', 'min', 'value'] as const).map((g) => (
          <label key={g} className="flex gap-1"><input type="radio" checked={spec.goal === g} onChange={() => upd({ goal: g })} /> {g === 'max' ? 'Max' : g === 'min' ? 'Min' : 'Value Of:'}</label>
        ))}
        <input className="xl-input w-20" disabled={spec.goal !== 'value'} value={spec.targetValue} onChange={(e) => upd({ targetValue: Number(e.target.value) || 0 })} />
      </div>
      <label className="flex items-center gap-2 my-1"><span className="w-28">By Changing Variable Cells:</span><RefInput value={spec.variables} onChange={(v) => upd({ variables: v })} width={180} testId="solver-vars" /></label>
      <div className="mt-2 mb-1">Subject to the Constraints:</div>
      <div className="flex gap-2">
        <div className="xl-list h-[110px] flex-1">
          {spec.constraints.map((c, i) => (
            <div key={i} className={sel === i ? 'xl-list-sel' : ''} onClick={() => setSel(i)}>{c.ref.replace(/^=/, '')} {c.op} {c.op === 'int' || c.op === 'bin' ? '' : c.value}</div>
          ))}
        </div>
        <div className="flex flex-col gap-1 w-[190px]">
          <RefInput value={newC.ref} onChange={(v) => setNewC({ ...newC, ref: v })} width={190} />
          <div className="flex gap-1">
            <select className="xl-input" value={newC.op} onChange={(e) => setNewC({ ...newC, op: e.target.value as Constraint['op'] })}>
              {['<=', '>=', '=', 'int', 'bin'].map((o) => <option key={o}>{o}</option>)}
            </select>
            <input className="xl-input flex-1 min-w-0" value={newC.value} onChange={(e) => setNewC({ ...newC, value: e.target.value })} placeholder="value or cell" />
          </div>
          <button className="xl-btn" onClick={() => { if (newC.ref) { upd({ constraints: [...spec.constraints, newC] }); setNewC({ ref: '', op: '<=', value: '' }); } }}>Add</button>
          <button className="xl-btn" disabled={sel === null} onClick={() => { upd({ constraints: spec.constraints.filter((_, i) => i !== sel) }); setSel(null); }}>Delete</button>
        </div>
      </div>
      <label className="flex gap-2 mt-2"><input type="checkbox" checked={spec.nonNegative} onChange={(e) => upd({ nonNegative: e.target.checked })} /> Make Unconstrained Variables Non-Negative</label>
      <div className="text-[11px] opacity-70 mt-2">Solving method: Nelder–Mead with constraint penalties and pattern-search polish (suitable for small smooth or linear models).</div>
    </Dialog>
  );
}

export function SolverResultsDialog({ props }: { props: Record<string, unknown> }) {
  const res = props.res as ReturnType<typeof solve>;
  const spec = props.spec as SolverSpec;
  const [keep, setKeep] = useState(true);
  return (
    <Dialog title="Solver Results" width={420} onOk={() => { if (keep && res.values.length) keepSolverSolution(spec, res.values); }} testId="solver-results">
      <div className="mb-2">{res.message}</div>
      <label className="flex gap-2 my-1"><input type="radio" checked={keep} onChange={() => setKeep(true)} /> Keep Solver Solution</label>
      <label className="flex gap-2 my-1"><input type="radio" checked={!keep} onChange={() => setKeep(false)} /> Restore Original Values</label>
      {res.objective !== null && <div className="mt-2 opacity-80">Objective value: {+res.objective.toPrecision(10)}</div>}
    </Dialog>
  );
}

export function ScenarioManagerDialog() {
  useStore((s) => s.rev);
  const sheet = S().wb.activeSheet;
  const [sel, setSel] = useState<string | null>(sheet.scenarios[0]?.id ?? null);
  const cur = sheet.scenarios.find((s) => s.id === sel);
  return (
    <Dialog title="Scenario Manager" width={460} okLabel="Close" cancelLabel={null} testId="scenario-manager">
      <div className="flex gap-3">
        <div className="flex-1">
          <div className="mb-1">Scenarios:</div>
          <div className="xl-list h-[160px]">
            {sheet.scenarios.map((s) => <div key={s.id} className={sel === s.id ? 'xl-list-sel' : ''} onClick={() => setSel(s.id)} onDoubleClick={() => showScenario(s.id)}>{s.name}</div>)}
            {!sheet.scenarios.length && <div className="opacity-60 whitespace-normal">No Scenarios defined. Choose Add to add scenarios.</div>}
          </div>
        </div>
        <div className="flex flex-col gap-1 w-[100px]">
          <button className="xl-btn" onClick={() => openDialog('scenarioEdit', {})}>Add...</button>
          <button className="xl-btn" disabled={!cur} onClick={() => cur && deleteScenario(cur.id)}>Delete</button>
          <button className="xl-btn" disabled={!cur} onClick={() => cur && openDialog('scenarioEdit', { id: cur.id })}>Edit...</button>
          <button className="xl-btn" disabled={!sheet.scenarios.length} onClick={() => openDialog('scenarioSummary')}>Summary...</button>
        </div>
      </div>
      {cur && (
        <div className="mt-2 text-[11px] opacity-80">Changing cells: {cur.cells.map((a) => addrToA1(a.r, a.c)).join(', ')}{cur.comment ? ` — ${cur.comment}` : ''}</div>
      )}
      <div className="flex justify-end mt-2">
        <button className="xl-btn xl-btn-primary" disabled={!cur} onClick={() => cur && showScenario(cur.id)} data-testid="scenario-show">Show</button>
      </div>
    </Dialog>
  );
}

export function ScenarioEditDialog({ props }: { props: Record<string, unknown> }) {
  const sheet = S().wb.activeSheet;
  const existing = sheet.scenarios.find((s) => s.id === props.id);
  const [name, setName] = useState(existing?.name ?? '');
  const [cells, setCells] = useState(existing ? existing.cells.map((a) => addrToA1(a.r, a.c, true, true)).join(',') : '=' + selectionRefText().split('!').pop());
  const [comment, setComment] = useState(existing?.comment ?? `Created by user on ${new Date().toLocaleDateString()}`);
  const [step, setStep] = useState<1 | 2>(1);
  const res = resolveReference(cells.replace(/^=/, '').split(',')[0]);
  const addrs: { r: number; c: number }[] = [];
  if (res) for (let r = res.range.r1; r <= Math.min(res.range.r2, res.range.r1 + 31); r++) for (let c = res.range.c1; c <= Math.min(res.range.c2, res.range.c1 + 31); c++) addrs.push({ r, c });
  const [values, setValues] = useState<string[]>(() => (existing ? existing.values.map((v) => String(v ?? '')) : addrs.map((a) => String(getScalar(sheet, a.r, a.c) ?? ''))));
  if (step === 2) {
    return (
      <Dialog
        title="Scenario Values"
        width={340}
        onOk={() => {
          const vals: CellValue[] = values.map((v) => (v === '' ? null : isNaN(Number(v)) ? v : Number(v)));
          if (!addScenario(name, cells, vals, comment, existing?.id)) return false;
          setTimeout(() => openDialog('scenarioManager'), 0);
        }}
      >
        <div className="mb-2">Enter values for each of the changing cells.</div>
        {addrs.map((a, i) => (
          <label key={i} className="flex items-center gap-2 my-1"><span className="w-16">{addrToA1(a.r, a.c, true, true)}</span><input className="xl-input flex-1" value={values[i] ?? ''} onChange={(e) => setValues(values.map((v, k) => (k === i ? e.target.value : v)))} /></label>
        ))}
      </Dialog>
    );
  }
  return (
    <Dialog
      title={existing ? 'Edit Scenario' : 'Add Scenario'}
      width={380}
      onOk={() => {
        if (!name.trim() || !res) {
          alertBox('Enter a scenario name and changing cells.');
          return false;
        }
        if (!existing) setValues(addrs.map((a) => String(getScalar(sheet, a.r, a.c) ?? '')));
        setStep(2);
        return false;
      }}
    >
      <label className="flex flex-col gap-1 my-1">Scenario name:<input className="xl-input" value={name} onChange={(e) => setName(e.target.value)} data-testid="scenario-name" /></label>
      <label className="flex flex-col gap-1 my-1">Changing cells:<RefInput value={cells} onChange={setCells} width={330} /></label>
      <label className="flex flex-col gap-1 my-1">Comment:<textarea className="xl-input !h-16 py-1" value={comment} onChange={(e) => setComment(e.target.value)} /></label>
    </Dialog>
  );
}

export function ScenarioSummaryDialog() {
  const [ref, setRef] = useState('=' + selectionRefText().split('!').pop());
  return (
    <Dialog title="Scenario Summary" width={320} onOk={() => scenarioSummary(ref)}>
      <div className="mb-1">Result cells:</div>
      <RefInput value={ref} onChange={setRef} width={280} />
    </Dialog>
  );
}

export function DataTableDialog() {
  const [rowIn, setRowIn] = useState('');
  const [colIn, setColIn] = useState('');
  return (
    <Dialog
      title="Data Table"
      width={320}
      onOk={() => {
        const err = dataTable(primaryRange(S().sel), rowIn, colIn);
        if (err) {
          alertBox(err);
          return false;
        }
      }}
      testId="data-table"
    >
      <label className="flex items-center gap-2 my-1"><span className="w-28">Row input cell:</span><RefInput value={rowIn} onChange={setRowIn} width={150} /></label>
      <label className="flex items-center gap-2 my-1"><span className="w-28">Column input cell:</span><RefInput value={colIn} onChange={setColIn} width={150} testId="dt-col" /></label>
    </Dialog>
  );
}

export function SparklineDialog({ props }: { props: Record<string, unknown> }) {
  const type = (props.type as 'line' | 'column' | 'winloss') ?? 'line';
  const [data, setData] = useState('=' + selectionRefText().split('!').pop());
  const [loc, setLoc] = useState('');
  return (
    <Dialog
      title="Create Sparklines"
      width={380}
      onOk={() => {
        const sheet = S().wb.activeSheet;
        const drg = parseRange(data.replace(/^=/, '').replace(/\$/g, '').replace(/^.*!/, ''));
        const lrg = parseRange(loc.replace(/^=/, '').replace(/\$/g, '').replace(/^.*!/, ''));
        if (!drg || !lrg) {
          alertBox('Data Range or Location Range is not valid.');
          return false;
        }
        const items: { r: number; c: number; ref: string }[] = [];
        if (lrg.c1 === lrg.c2) {
          // one sparkline per data row
          for (let i = 0; i <= lrg.r2 - lrg.r1; i++) {
            const r = drg.r1 + i;
            if (r > drg.r2) break;
            items.push({ r: lrg.r1 + i, c: lrg.c1, ref: `${addrToA1(r, drg.c1)}:${addrToA1(r, drg.c2)}` });
          }
        } else {
          for (let i = 0; i <= lrg.c2 - lrg.c1; i++) {
            const c = drg.c1 + i;
            if (c > drg.c2) break;
            items.push({ r: lrg.r1, c: lrg.c1 + i, ref: `${addrToA1(drg.r1, c)}:${addrToA1(drg.r2, c)}` });
          }
        }
        if (!items.length) {
          alertBox('Location reference is not valid.');
          return false;
        }
        transact('Insert Sparklines', (tx) => tx.setMeta(sheet, 'sparklines', sheet.sparklines.concat([{ id: newId('sp'), type, items, color: '#4472C4', negColor: '#C00000', markers: false, highPoint: false, lowPoint: false }])));
      }}
      testId="sparklines"
    >
      <div className="mb-1">Choose the data that you want</div>
      <label className="flex items-center gap-2 my-1"><span className="w-24">Data Range:</span><RefInput value={data} onChange={setData} width={220} /></label>
      <div className="mb-1 mt-2">Choose where you want the sparklines to be placed</div>
      <label className="flex items-center gap-2 my-1"><span className="w-24">Location Range:</span><RefInput value={loc} onChange={setLoc} width={220} testId="spark-loc" /></label>
    </Dialog>
  );
}

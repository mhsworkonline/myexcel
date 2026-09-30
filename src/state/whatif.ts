// What-If Analysis: Goal Seek, Solver-lite, Scenario Manager, Data Tables.
import { isErrorVal } from '../engine/Engine';
import { addrToA1, CellAddr, Range } from '../model/address';
import { newId, Sheet } from '../model/sheet';
import type { Cell, CellValue, Scenario } from '../model/types';
import { FnCommand } from '../model/commands';
import { resolveReference } from './actions/find';
import { writeInput } from './actions/edit';
import { alertBox, S, transact } from './store';

function num(v: unknown): number | null {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  return null;
}

/** Temporarily set a cell's value (not undoable; used inside solvers). */
function setRaw(sheet: Sheet, r: number, c: number, v: number): void {
  const cur = sheet.getCell(r, c) ?? {};
  const next: Cell = { ...cur, v };
  delete next.f;
  delete next.e;
  S().wb.setCells(sheet, [{ r, c, cell: next }]);
}

function restoreRaw(sheet: Sheet, r: number, c: number, cell: Cell | undefined): void {
  S().wb.setCells(sheet, [{ r, c, cell }]);
}

function readNum(sheet: Sheet, r: number, c: number): number | null {
  return num(S().engine.getValue(sheet, r, c));
}

export function parseCellRef(text: string): { sheet: Sheet; r: number; c: number } | null {
  const res = resolveReference(text.replace(/^=/, ''));
  if (!res) return null;
  const sheet = S().wb.sheetById(res.sheetId)!;
  return { sheet, r: res.range.r1, c: res.range.c1 };
}

export interface GoalSeekResult {
  found: boolean;
  value: number;
  target: number;
  achieved: number | null;
  iterations: number;
}

/** Secant method with bracketing fallback. Leaves the changing cell at the best value (not yet committed). */
export function goalSeek(setCellRef: string, toValue: number, changingRef: string): GoalSeekResult | string {
  const setC = parseCellRef(setCellRef);
  const chg = parseCellRef(changingRef);
  if (!setC || !chg) return 'Reference is not valid.';
  if (!setC.sheet.getCell(setC.r, setC.c)?.f) return 'Cell must contain a formula.';
  const chgCell = chg.sheet.getCell(chg.r, chg.c);
  if (chgCell?.f) return 'Cell must contain a value.';
  S().engine.attachIfNeeded();
  const f = (x: number) => {
    setRaw(chg.sheet, chg.r, chg.c, x);
    const y = readNum(setC.sheet, setC.r, setC.c);
    return y === null ? NaN : y - toValue;
  };
  let x0 = num(chgCell?.v) ?? 0;
  let x1 = x0 === 0 ? 1 : x0 * 1.01 + 0.01;
  let f0 = f(x0);
  let f1 = f(x1);
  let best = Math.abs(f0) < Math.abs(f1) ? x0 : x1;
  let bestErr = Math.min(Math.abs(f0), Math.abs(f1));
  let it = 0;
  const tol = Math.max(1e-7, Math.abs(toValue) * 1e-9);
  for (; it < 200 && bestErr > tol; it++) {
    if (!isFinite(f0) || !isFinite(f1) || f1 === f0) {
      x1 = x1 + (it + 1) * (Math.abs(x1) + 1);
      f1 = f(x1);
      continue;
    }
    const x2 = x1 - (f1 * (x1 - x0)) / (f1 - f0);
    if (!isFinite(x2)) break;
    x0 = x1;
    f0 = f1;
    x1 = x2;
    f1 = f(x1);
    if (Math.abs(f1) < bestErr) {
      bestErr = Math.abs(f1);
      best = x1;
    }
  }
  f(best);
  const achieved = readNum(setC.sheet, setC.r, setC.c);
  restoreRaw(chg.sheet, chg.r, chg.c, chgCell);
  return { found: bestErr <= Math.max(tol, 1e-6 * Math.max(1, Math.abs(toValue))), value: best, target: toValue, achieved, iterations: it };
}

export function commitCellValue(sheet: Sheet, r: number, c: number, v: number, label: string): void {
  transact(label, (tx) => {
    const cur = sheet.getCell(r, c) ?? {};
    const next: Cell = { ...cur, v: +v.toPrecision(15) };
    delete next.f;
    delete next.e;
    tx.setCell(sheet, r, c, next);
  });
}

// ---------- Solver-lite ----------

export interface Constraint {
  ref: string; // cell or range
  op: '<=' | '>=' | '=' | 'int' | 'bin';
  value: string; // number or cell ref
}

export interface SolverSpec {
  objective: string;
  goal: 'max' | 'min' | 'value';
  targetValue: number;
  variables: string; // range
  constraints: Constraint[];
  nonNegative: boolean;
}

function rangeCells(text: string): { sheet: Sheet; cells: CellAddr[] } | null {
  const res = resolveReference(text.replace(/^=/, ''));
  if (!res) return null;
  const sheet = S().wb.sheetById(res.sheetId)!;
  const cells: CellAddr[] = [];
  const rg = res.range;
  for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 50); r++) for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 50); c++) cells.push({ r, c });
  return { sheet, cells };
}

export interface SolverResult {
  ok: boolean;
  message: string;
  values: number[];
  objective: number | null;
}

/** Derivative-free penalty method (Nelder–Mead + pattern search). Suitable for small models. */
export function solve(spec: SolverSpec): SolverResult {
  const obj = parseCellRef(spec.objective);
  const vars = rangeCells(spec.variables);
  if (!obj || !vars || !vars.cells.length) return { ok: false, message: 'Objective or variable cells are not valid.', values: [], objective: null };
  S().engine.attachIfNeeded();
  const sheet = vars.sheet;
  const originals = vars.cells.map((a) => sheet.getCell(a.r, a.c));
  const x0 = vars.cells.map((a) => num(sheet.getCell(a.r, a.c)?.v) ?? 0);
  const intIdx = new Set<number>();
  const binIdx = new Set<number>();
  type C = { cells: { sheet: Sheet; cells: CellAddr[] }; op: Constraint['op']; rhs: () => number };
  const cons: C[] = [];
  for (const c of spec.constraints) {
    const cells = rangeCells(c.ref);
    if (!cells) continue;
    if (c.op === 'int' || c.op === 'bin') {
      cells.cells.forEach((a) => {
        const k = vars.cells.findIndex((v) => v.r === a.r && v.c === a.c);
        if (k >= 0) (c.op === 'int' ? intIdx : binIdx).add(k);
      });
      continue;
    }
    const rhsRef = isNaN(Number(c.value)) ? parseCellRef(c.value) : null;
    const rhsNum = Number(c.value);
    cons.push({ cells, op: c.op, rhs: () => (rhsRef ? readNum(rhsRef.sheet, rhsRef.r, rhsRef.c) ?? 0 : rhsNum) });
  }
  const apply = (x: number[]) => {
    S().wb.setCells(
      sheet,
      vars.cells.map((a, i) => {
        const cur = sheet.getCell(a.r, a.c) ?? {};
        const n: Cell = { ...cur, v: x[i] };
        delete n.f;
        return { r: a.r, c: a.c, cell: n };
      }),
    );
  };
  const round = (x: number[]) => x.map((v, i) => (binIdx.has(i) ? (v >= 0.5 ? 1 : 0) : intIdx.has(i) ? Math.round(v) : v));
  const violation = (x: number[]) => {
    let pen = 0;
    x.forEach((v, i) => {
      if (spec.nonNegative && v < 0) pen += -v;
      if (binIdx.has(i)) pen += Math.min(Math.abs(v), Math.abs(v - 1)) * 0.1;
    });
    for (const c of cons) {
      const rhs = c.rhs();
      for (const a of c.cells.cells) {
        const lhs = readNum(c.cells.sheet, a.r, a.c) ?? 0;
        if (c.op === '<=' && lhs > rhs) pen += lhs - rhs;
        if (c.op === '>=' && lhs < rhs) pen += rhs - lhs;
        if (c.op === '=') pen += Math.abs(lhs - rhs);
      }
    }
    return pen;
  };
  const score = (xRaw: number[]) => {
    const x = round(xRaw);
    apply(x);
    const y = readNum(obj.sheet, obj.r, obj.c);
    if (y === null) return 1e30;
    const base = spec.goal === 'max' ? -y : spec.goal === 'min' ? y : Math.abs(y - spec.targetValue);
    return base + violation(x) * 1e6;
  };
  const n = x0.length;
  // Nelder–Mead
  let simplex = [x0.slice()];
  for (let i = 0; i < n; i++) {
    const p = x0.slice();
    p[i] = p[i] !== 0 ? p[i] * 1.5 : 1;
    simplex.push(p);
  }
  let vals = simplex.map(score);
  for (let iter = 0; iter < 400 * n; iter++) {
    const order = vals.map((v, i) => i).sort((a, b) => vals[a] - vals[b]);
    simplex = order.map((i) => simplex[i]);
    vals = order.map((i) => vals[i]);
    if (Math.abs(vals[n] - vals[0]) < 1e-10 * (1 + Math.abs(vals[0])) && iter > 20) break;
    const centroid = new Array(n).fill(0);
    for (let i = 0; i < n; i++) for (let j = 0; j < n; j++) centroid[j] += simplex[i][j] / n;
    const worst = simplex[n];
    const refl = centroid.map((c, j) => c + (c - worst[j]));
    const fr = score(refl);
    if (fr < vals[0]) {
      const exp = centroid.map((c, j) => c + 2 * (c - worst[j]));
      const fe = score(exp);
      if (fe < fr) {
        simplex[n] = exp;
        vals[n] = fe;
      } else {
        simplex[n] = refl;
        vals[n] = fr;
      }
    } else if (fr < vals[n - 1]) {
      simplex[n] = refl;
      vals[n] = fr;
    } else {
      const con = centroid.map((c, j) => c + 0.5 * (worst[j] - c));
      const fc = score(con);
      if (fc < vals[n]) {
        simplex[n] = con;
        vals[n] = fc;
      } else {
        for (let i = 1; i <= n; i++) {
          simplex[i] = simplex[i].map((v, j) => simplex[0][j] + 0.5 * (v - simplex[0][j]));
          vals[i] = score(simplex[i]);
        }
      }
    }
  }
  let best = simplex[0];
  let bestVal = vals[0];
  // pattern search polish
  let step = 1;
  for (let k = 0; k < 60 && step > 1e-7; k++) {
    let improved = false;
    for (let i = 0; i < n; i++)
      for (const d of [step, -step]) {
        const t = best.slice();
        t[i] += d * (Math.abs(t[i]) > 1 ? Math.abs(t[i]) : 1);
        const v = score(t);
        if (v < bestVal) {
          best = t;
          bestVal = v;
          improved = true;
        }
      }
    if (!improved) step /= 2;
  }
  const finalX = round(best);
  apply(finalX);
  const feasible = violation(finalX) < 1e-6;
  const objective = readNum(obj.sheet, obj.r, obj.c);
  // restore originals; caller decides to keep
  S().wb.setCells(sheet, vars.cells.map((a, i) => ({ r: a.r, c: a.c, cell: originals[i] })));
  return {
    ok: feasible,
    message: feasible ? 'Solver found a solution. All Constraints and optimality conditions are satisfied.' : 'Solver could not find a feasible solution.',
    values: finalX.map((v) => +v.toPrecision(12)),
    objective,
  };
}

export function keepSolverSolution(spec: SolverSpec, values: number[]): void {
  const vars = rangeCells(spec.variables);
  if (!vars) return;
  transact('Solver', (tx) => {
    vars.cells.forEach((a, i) => {
      const cur = vars.sheet.getCell(a.r, a.c) ?? {};
      const n: Cell = { ...cur, v: values[i] };
      delete n.f;
      delete n.e;
      tx.setCell(vars.sheet, a.r, a.c, n);
    });
  });
}

// ---------- Scenarios ----------

export function addScenario(name: string, cellsRef: string, values: CellValue[], comment?: string, replaceId?: string): boolean {
  const sheet = S().wb.activeSheet;
  const res = rangeCells(cellsRef);
  if (!res || res.sheet !== sheet) {
    alertBox('Changing cells must be on the current sheet.');
    return false;
  }
  const sc: Scenario = { id: replaceId ?? newId('sc'), name, cells: res.cells, values, comment };
  const list = replaceId ? sheet.scenarios.map((s) => (s.id === replaceId ? sc : s)) : sheet.scenarios.concat([sc]);
  transact('Scenario', (tx) => tx.setMeta(sheet, 'scenarios', list));
  return true;
}

export function deleteScenario(id: string): void {
  const sheet = S().wb.activeSheet;
  transact('Delete Scenario', (tx) => tx.setMeta(sheet, 'scenarios', sheet.scenarios.filter((s) => s.id !== id)));
}

export function showScenario(id: string): void {
  const sheet = S().wb.activeSheet;
  const sc = sheet.scenarios.find((s) => s.id === id);
  if (!sc) return;
  transact(`Show Scenario ${sc.name}`, (tx) => {
    sc.cells.forEach((a, i) => {
      const v = sc.values[i];
      writeInput(tx, sheet, a.r, a.c, v === null || v === undefined ? '' : String(v));
    });
  });
}

export function scenarioSummary(resultRef: string): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  const results = rangeCells(resultRef);
  const scs = sheet.scenarios;
  if (!scs.length) return;
  const changing = scs[0].cells;
  const summary = new Sheet(st.wb.uniqueSheetName('Scenario Summary'));
  const bold = st.wb.styles.intern({ bold: true });
  const hdr = st.wb.styles.intern({ bold: true, fillColor: '#D9E1F2' });
  const set = (r: number, c: number, v: CellValue, s?: number) => summary.setCellRaw(r, c, { v, ...(s ? { s } : {}) });
  set(0, 1, 'Scenario Summary', bold);
  set(2, 1, '', hdr);
  set(2, 2, 'Current Values:', hdr);
  scs.forEach((s, i) => set(2, 3 + i, s.name, hdr));
  set(3, 1, 'Changing Cells:', bold);
  changing.forEach((a, k) => {
    set(4 + k, 2 - 0, addrToA1(a.r, a.c, true, true));
    const cur = S().engine.getValue(sheet, a.r, a.c);
    set(4 + k, 2, isErrorVal(cur) ? cur.error : (cur as CellValue));
    set(4 + k, 1, addrToA1(a.r, a.c, true, true));
    scs.forEach((s, i) => set(4 + k, 3 + i, s.values[k] ?? null));
  });
  let row = 4 + changing.length;
  if (results) {
    set(row, 1, 'Result Cells:', bold);
    row++;
    // evaluate each scenario by applying values temporarily
    const originals = changing.map((a) => sheet.getCell(a.r, a.c));
    results.cells.forEach((a, k) => set(row + k, 1, addrToA1(a.r, a.c, true, true)));
    results.cells.forEach((a, k) => {
      const v = S().engine.getValue(results.sheet, a.r, a.c);
      set(row + k, 2, isErrorVal(v) ? v.error : (v as CellValue));
    });
    scs.forEach((s, i) => {
      S().wb.setCells(sheet, s.cells.map((a, k) => ({ r: a.r, c: a.c, cell: { ...(sheet.getCell(a.r, a.c) ?? {}), v: s.values[k] ?? undefined, f: undefined } as Cell })));
      results.cells.forEach((a, k) => {
        const v = S().engine.getValue(results.sheet, a.r, a.c);
        set(row + k, 3 + i, isErrorVal(v) ? v.error : (v as CellValue));
      });
    });
    S().wb.setCells(sheet, changing.map((a, k) => ({ r: a.r, c: a.c, cell: originals[k] })));
  }
  summary.colWidths.set(1, 110);
  summary.colWidths.set(2, 100);
  const prev = st.wb.activeSheetId;
  const idx = st.wb.sheets.indexOf(sheet);
  transact('Scenario Summary', (tx) =>
    tx.run(
      new FnCommand(
        'Scenario Summary',
        (w) => {
          w.insertSheet(summary, idx);
          w.activeSheetId = summary.id;
        },
        (w) => {
          w.removeSheet(summary);
          w.activeSheetId = prev;
        },
      ),
    ),
  );
}

// ---------- Data Tables ----------

/**
 * One- or two-variable data table over `range` (Excel layout):
 *  - column input: formulas in the top row (right of the corner), inputs down the left column
 *  - row input: formulas in the left column (below the corner), inputs across the top row
 *  - both: formula in the corner, row inputs across the top, column inputs down the left
 */
export function dataTable(range: Range, rowInputRef: string, colInputRef: string): string | null {
  const sheet = S().wb.activeSheet;
  const rowIn = rowInputRef ? parseCellRef(rowInputRef) : null;
  const colIn = colInputRef ? parseCellRef(colInputRef) : null;
  if (!rowIn && !colIn) return 'Input cell reference is not valid.';
  S().engine.attachIfNeeded();
  const out: { r: number; c: number; v: CellValue }[] = [];
  const evalAt = (fr: number, fc: number) => {
    const v = S().engine.getValue(sheet, fr, fc);
    return isErrorVal(v) ? v.error : (v as CellValue);
  };
  const withInputs = (vals: [ReturnType<typeof parseCellRef>, CellValue][], fn: () => void) => {
    const saved = vals.map(([ref]) => (ref ? ref.sheet.getCell(ref.r, ref.c) : undefined));
    vals.forEach(([ref, v]) => ref && S().wb.setCells(ref.sheet, [{ r: ref.r, c: ref.c, cell: { ...(ref.sheet.getCell(ref.r, ref.c) ?? {}), v: v ?? undefined, f: undefined } as Cell }]));
    fn();
    vals.forEach(([ref], i) => ref && S().wb.setCells(ref.sheet, [{ r: ref.r, c: ref.c, cell: saved[i] }]));
  };
  const cellVal = (r: number, c: number) => S().engine.getValue(sheet, r, c) as CellValue;
  if (rowIn && colIn) {
    for (let r = range.r1 + 1; r <= range.r2; r++)
      for (let c = range.c1 + 1; c <= range.c2; c++)
        withInputs([[rowIn, cellVal(range.r1, c)], [colIn, cellVal(r, range.c1)]], () => out.push({ r, c, v: evalAt(range.r1, range.c1) }));
  } else if (colIn) {
    for (let r = range.r1 + 1; r <= range.r2; r++)
      withInputs([[colIn, cellVal(r, range.c1)]], () => {
        for (let c = range.c1 + 1; c <= range.c2; c++) out.push({ r, c, v: evalAt(range.r1, c) });
      });
  } else if (rowIn) {
    for (let c = range.c1 + 1; c <= range.c2; c++)
      withInputs([[rowIn, cellVal(range.r1, c)]], () => {
        for (let r = range.r1 + 1; r <= range.r2; r++) out.push({ r, c, v: evalAt(r, range.c1) });
      });
  }
  transact('Data Table', (tx) => {
    for (const o of out) {
      const cur = sheet.getCell(o.r, o.c) ?? {};
      const n: Cell = { ...cur };
      delete n.f;
      delete n.e;
      if (o.v === null || o.v === undefined) delete n.v;
      else n.v = o.v;
      if (typeof o.v === 'string' && o.v.startsWith('#')) n.e = true;
      tx.setCell(sheet, o.r, o.c, n);
    }
  });
  return null;
}

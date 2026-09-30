// Pivot table engine: builds a snapshot grid from a source range and writes it into cells.
import { isErrorVal } from '../engine/Engine';
import { Range } from '../model/address';
import { FnCommand, Tx } from '../model/commands';
import { formatValue, serialToParts } from '../model/numfmt';
import { newId, Sheet } from '../model/sheet';
import type { Cell, CellValue, PivotAgg, PivotSpec } from '../model/types';
import type { Workbook } from '../model/workbook';
import { S, transact, setState } from './store';
import { getScalar } from './values';
import { measureColumn } from './actions/format';

export interface PivotField {
  name: string;
  col: number;
  numeric: boolean;
  date: boolean;
}

export function pivotFields(wb: Workbook, spec: Pick<PivotSpec, 'sourceSheetId' | 'sourceRange'>): PivotField[] {
  const src = wb.sheetById(spec.sourceSheetId);
  if (!src) return [];
  const rg = spec.sourceRange;
  const out: PivotField[] = [];
  const seen = new Set<string>();
  for (let c = rg.c1; c <= rg.c2; c++) {
    let name = String(getScalar(src, rg.r1, c) ?? '').trim() || `Column${c - rg.c1 + 1}`;
    let k = 2;
    const base = name;
    while (seen.has(name)) name = `${base}${k++}`;
    seen.add(name);
    let nums = 0;
    let dates = 0;
    let total = 0;
    for (let r = rg.r1 + 1; r <= Math.min(rg.r2, rg.r1 + 200); r++) {
      const v = getScalar(src, r, c);
      if (v === null || v === '') continue;
      total++;
      if (typeof v === 'number') {
        nums++;
        const fmt = wb.styles.get(src.styleIdAt(r, c)).numFmt ?? '';
        if (/[dmy]/i.test(fmt.replace(/"[^"]*"|\[[^\]]*\]/g, '')) && !/0/.test(fmt)) dates++;
      }
    }
    out.push({ name, col: c, numeric: total > 0 && nums === total, date: total > 0 && dates === total });
  }
  return out;
}

const QUARTERS = ['Qtr1', 'Qtr2', 'Qtr3', 'Qtr4'];
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

type Key = string;

function labelOf(v: CellValue, field: PivotField, spec: PivotSpec, wb: Workbook, src: Sheet, r: number): { label: string; sort: number | string } {
  const g = spec.grouping.find((x) => x.field === field.name);
  if (g && typeof v === 'number') {
    if (g.type === 'date') {
      const p = serialToParts(v);
      switch (g.by) {
        case 'years': return { label: String(p.y), sort: p.y };
        case 'quarters': return { label: QUARTERS[Math.floor((p.m - 1) / 3)], sort: Math.floor((p.m - 1) / 3) };
        case 'days': return { label: `${p.d}-${MONTHS[p.m - 1]}`, sort: p.m * 100 + p.d };
        default: return { label: MONTHS[p.m - 1], sort: p.m };
      }
    }
    const start = g.start ?? 0;
    const step = g.interval || 10;
    const bin = Math.floor((v - start) / step);
    const lo = start + bin * step;
    return { label: `${lo}-${lo + step - 1}`, sort: lo };
  }
  if (v === null || v === '') return { label: '(blank)', sort: '￿' };
  if (isErrorVal(v as never)) return { label: String(v), sort: String(v) };
  if (typeof v === 'number') {
    const fmt = wb.styles.get(src.styleIdAt(r, field.col)).numFmt;
    return { label: formatValue(v, fmt).text, sort: v };
  }
  if (typeof v === 'boolean') return { label: v ? 'TRUE' : 'FALSE', sort: v ? 1 : 0 };
  return { label: String(v), sort: String(v).toLowerCase() };
}

class Acc {
  sum = 0;
  count = 0;
  countNums = 0;
  min = Infinity;
  max = -Infinity;
  prod = 1;
  sq = 0;
  add(v: CellValue): void {
    if (v === null || v === '') return;
    this.count++;
    if (typeof v === 'number') {
      this.countNums++;
      this.sum += v;
      this.sq += v * v;
      this.prod *= v;
      if (v < this.min) this.min = v;
      if (v > this.max) this.max = v;
    }
  }
  value(agg: PivotAgg): number | null {
    switch (agg) {
      case 'sum': return this.countNums ? this.sum : this.count ? 0 : null;
      case 'count': return this.count || null;
      case 'countNums': return this.countNums || null;
      case 'average': return this.countNums ? this.sum / this.countNums : null;
      case 'max': return this.countNums ? this.max : null;
      case 'min': return this.countNums ? this.min : null;
      case 'product': return this.countNums ? this.prod : null;
      case 'stdDev': return this.countNums > 1 ? Math.sqrt((this.sq - (this.sum * this.sum) / this.countNums) / (this.countNums - 1)) : null;
      case 'var': return this.countNums > 1 ? (this.sq - (this.sum * this.sum) / this.countNums) / (this.countNums - 1) : null;
    }
    return null;
  }
}

export const AGG_LABEL: Record<PivotAgg, string> = { sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min', product: 'Product', countNums: 'Count Numbers', stdDev: 'StdDev', var: 'Var' };

export type PivotCellKind = 'filterLabel' | 'filterValue' | 'corner' | 'colHeader' | 'rowHeader' | 'rowLabel' | 'rowLabelBold' | 'value' | 'valueBold' | 'grandLabel' | 'grand' | 'blank';

export interface PivotGrid {
  cells: CellValue[][];
  kinds: PivotCellKind[][];
  indent: number[][];
  numFmts: (string | undefined)[][];
}

export function computePivot(wb: Workbook, spec: PivotSpec): PivotGrid {
  const src = wb.sheetById(spec.sourceSheetId);
  const grid: PivotGrid = { cells: [], kinds: [], indent: [], numFmts: [] };
  if (!src) return grid;
  const fields = pivotFields(wb, spec);
  const fieldByName = new Map(fields.map((f) => [f.name, f]));
  const rg = spec.sourceRange;
  const calc = new Map(spec.calculatedFields.map((c) => [c.name, c.formula]));
  const valueDefs = spec.values.filter((v) => fieldByName.has(v.field) || calc.has(v.field));
  const rowFields = spec.rows.map((n) => fieldByName.get(n)).filter((f): f is PivotField => !!f);
  const colFields = spec.cols.map((n) => fieldByName.get(n)).filter((f): f is PivotField => !!f);
  // records
  type Rec = { r: number; rowLabels: { label: string; sort: number | string }[]; colLabels: { label: string; sort: number | string }[] };
  const recs: Rec[] = [];
  for (let r = rg.r1 + 1; r <= rg.r2; r++) {
    let any = false;
    for (let c = rg.c1; c <= rg.c2; c++) {
      const v = getScalar(src, r, c);
      if (v !== null && v !== '') {
        any = true;
        break;
      }
    }
    if (!any) continue;
    let pass = true;
    for (const f of spec.filters) {
      if (!f.selected) continue;
      const fd = fieldByName.get(f.field);
      if (!fd) continue;
      if (!f.selected.includes(labelOf(getScalar(src, r, fd.col), fd, spec, wb, src, r).label)) {
        pass = false;
        break;
      }
    }
    if (!pass) continue;
    recs.push({
      r,
      rowLabels: rowFields.map((f) => labelOf(getScalar(src, r, f.col), f, spec, wb, src, r)),
      colLabels: colFields.map((f) => labelOf(getScalar(src, r, f.col), f, spec, wb, src, r)),
    });
  }
  // distinct sorted key paths
  const sortKey = (a: { sort: number | string }, b: { sort: number | string }) =>
    typeof a.sort === 'number' && typeof b.sort === 'number' ? a.sort - b.sort : String(a.sort).localeCompare(String(b.sort));
  const buildTree = (get: (r: Rec) => { label: string; sort: number | string }[]) => {
    const paths = new Map<Key, { labels: { label: string; sort: number | string }[] }>();
    for (const rec of recs) {
      const ls = get(rec);
      for (let d = 1; d <= ls.length; d++) {
        const k = ls.slice(0, d).map((x) => x.label).join('\u0001');
        if (!paths.has(k)) paths.set(k, { labels: ls.slice(0, d) });
      }
    }
    const list = [...paths.entries()].map(([k, v]) => ({ key: k, labels: v.labels }));
    list.sort((a, b) => {
      const n = Math.min(a.labels.length, b.labels.length);
      for (let i = 0; i < n; i++) {
        const c = sortKey(a.labels[i], b.labels[i]);
        if (c !== 0) return c;
      }
      return a.labels.length - b.labels.length;
    });
    return list;
  };
  const rowPaths = rowFields.length ? buildTree((r) => r.rowLabels) : [];
  const colLeaves = colFields.length ? buildTree((r) => r.colLabels).filter((p) => p.labels.length === colFields.length) : [];
  // accumulators keyed by rowPath|colLeaf|valueIndex
  const accs = new Map<string, Acc[]>();
  const accFor = (rk: string, ck: string) => {
    const k = rk + '\u0002' + ck;
    let a = accs.get(k);
    if (!a) accs.set(k, (a = []));
    return a;
  };
  const fieldNamesInCalc = (formula: string) => fields.filter((f) => formula.includes(f.name)).map((f) => f.name);
  const neededFields = new Set<string>();
  valueDefs.forEach((v) => {
    if (calc.has(v.field)) fieldNamesInCalc(calc.get(v.field)!).forEach((n) => neededFields.add(n));
  });
  const calcFieldList = [...neededFields];
  for (const rec of recs) {
    const rks = [''];
    for (let d = 1; d <= rec.rowLabels.length; d++) rks.push(rec.rowLabels.slice(0, d).map((x) => x.label).join('\u0001'));
    const cks = [''];
    if (colFields.length) cks.push(rec.colLabels.map((x) => x.label).join('\u0001'));
    for (const rk of rks)
      for (const ck of cks) {
        const a = accFor(rk, ck);
        valueDefs.forEach((vd, i) => {
          if (calc.has(vd.field)) return;
          const f = fieldByName.get(vd.field)!;
          (a[i] ??= new Acc()).add(getScalar(src, rec.r, f.col));
        });
        calcFieldList.forEach((n, j) => {
          const f = fieldByName.get(n)!;
          (a[valueDefs.length + j] ??= new Acc()).add(getScalar(src, rec.r, f.col));
        });
      }
  }
  const engine = S().engine;
  const valueAt = (rk: string, ck: string, i: number): number | null => {
    const a = accs.get(rk + '\u0002' + ck);
    if (!a) return null;
    const vd = valueDefs[i];
    if (calc.has(vd.field)) {
      let f = calc.get(vd.field)!;
      const names = [...calcFieldList].sort((x, y) => y.length - x.length);
      for (const n of names) {
        const acc = a[valueDefs.length + calcFieldList.indexOf(n)];
        const num = acc ? acc.sum : 0;
        f = f.split(`'${n}'`).join(String(num)).split(n).join(String(num));
      }
      const v = engine.evaluate(f.startsWith('=') ? f : '=' + f, src);
      return typeof v === 'number' ? v : null;
    }
    return a[i]?.value(vd.agg) ?? null;
  };
  const vName = (vd: PivotSpec['values'][0]) => vd.name ?? (calc.has(vd.field) ? `Sum of ${vd.field}` : `${AGG_LABEL[vd.agg]} of ${vd.field}`);
  const srcFmt = (vd: PivotSpec['values'][0]) => {
    const f = fieldByName.get(vd.field);
    if (!f || vd.agg === 'count' || vd.agg === 'countNums') return undefined;
    return wb.styles.get(src.styleIdAt(rg.r1 + 1, f.col)).numFmt;
  };

  const push = (cells: CellValue[], kinds: PivotCellKind[], indent: number[] = [], fmts: (string | undefined)[] = []) => {
    grid.cells.push(cells);
    grid.kinds.push(kinds);
    grid.indent.push(indent);
    grid.numFmts.push(fmts);
  };
  // filter area
  for (const f of spec.filters) {
    const sel = f.selected;
    push([f.field, sel === null ? '(All)' : sel.length === 1 ? sel[0] : '(Multiple Items)'], ['filterLabel', 'filterValue']);
  }
  if (spec.filters.length) push([], []);
  const nVals = Math.max(1, valueDefs.length);
  const colCombos: { ck: string; label: string; vi: number }[] = [];
  if (colFields.length) {
    for (const leaf of colLeaves) for (let i = 0; i < nVals; i++) colCombos.push({ ck: leaf.key, label: leaf.labels.map((l) => l.label).join(' - ') + (nVals > 1 ? ` ${valueDefs[i] ? vName(valueDefs[i]) : ''}` : ''), vi: i });
    if (spec.showGrandTotals) for (let i = 0; i < nVals; i++) colCombos.push({ ck: '', label: nVals > 1 ? `Total ${valueDefs[i] ? vName(valueDefs[i]) : ''}` : 'Grand Total', vi: i });
  } else for (let i = 0; i < nVals; i++) colCombos.push({ ck: '', label: valueDefs[i] ? vName(valueDefs[i]) : '', vi: i });
  // header rows
  if (colFields.length) {
    push([nVals === 1 && valueDefs[0] ? vName(valueDefs[0]) : 'Values', 'Column Labels'], ['corner', 'colHeader']);
    push([rowFields.length ? 'Row Labels' : '', ...colCombos.map((c) => c.label)], ['rowHeader', ...colCombos.map(() => 'colHeader' as PivotCellKind)]);
  } else if (valueDefs.length) {
    push([rowFields.length ? 'Row Labels' : '', ...colCombos.map((c) => c.label)], ['rowHeader', ...colCombos.map(() => 'colHeader' as PivotCellKind)]);
  } else if (rowFields.length) push(['Row Labels'], ['rowHeader']);
  // body
  for (const p of rowPaths) {
    const depth = p.labels.length;
    const isLeaf = depth === rowFields.length;
    const cells: CellValue[] = [p.labels[depth - 1].label];
    const kinds: PivotCellKind[] = [isLeaf ? 'rowLabel' : 'rowLabelBold'];
    const fmts: (string | undefined)[] = [undefined];
    for (const cc of colCombos) {
      cells.push(valueDefs.length ? valueAt(p.key, cc.ck, cc.vi) : null);
      kinds.push(isLeaf ? 'value' : 'valueBold');
      fmts.push(valueDefs[cc.vi] ? srcFmt(valueDefs[cc.vi]) : undefined);
    }
    push(cells, kinds, [depth - 1], fmts);
  }
  if (!rowFields.length && valueDefs.length) {
    push(['Total', ...colCombos.map((cc) => valueAt('', cc.ck, cc.vi))], ['grandLabel', ...colCombos.map(() => 'grand' as PivotCellKind)], [], [undefined, ...colCombos.map((cc) => srcFmt(valueDefs[cc.vi]))]);
  } else if (rowFields.length && spec.showGrandTotals) {
    push(['Grand Total', ...colCombos.map((cc) => (valueDefs.length ? valueAt('', cc.ck, cc.vi) : null))], ['grandLabel', ...colCombos.map(() => 'grand' as PivotCellKind)], [], [undefined, ...colCombos.map((cc) => (valueDefs[cc.vi] ? srcFmt(valueDefs[cc.vi]) : undefined))]);
  }
  return grid;
}

const PIVOT_HEADER_FILL = '#DDEBF7';
const PIVOT_BORDER = { style: 'thin' as const, color: '#9BC2E6' };

/** Write the pivot output into its target sheet (inside a transaction). */
export function writePivot(tx: Tx, wb: Workbook, spec: PivotSpec, host: Sheet): PivotSpec {
  const g = computePivot(wb, spec);
  // clear previous output
  if (spec.lastRange) host.forEachInRange(spec.lastRange, (r, c) => tx.setCell(host, r, c, undefined));
  const { r: r0, c: c0 } = spec.target;
  const styles = wb.styles;
  const sid = (patch: Parameters<typeof styles.intern>[0]) => styles.intern(patch);
  const kindStyle: Record<PivotCellKind, number> = {
    filterLabel: sid({ bold: true }),
    filterValue: sid({}),
    corner: sid({ bold: true, fillColor: PIVOT_HEADER_FILL, bBottom: PIVOT_BORDER }),
    colHeader: sid({ bold: true, fillColor: PIVOT_HEADER_FILL, bBottom: PIVOT_BORDER }),
    rowHeader: sid({ bold: true, fillColor: PIVOT_HEADER_FILL, bBottom: PIVOT_BORDER }),
    rowLabel: sid({}),
    rowLabelBold: sid({ bold: true }),
    value: sid({}),
    valueBold: sid({ bold: true }),
    grandLabel: sid({ bold: true, fillColor: PIVOT_HEADER_FILL, bTop: PIVOT_BORDER }),
    grand: sid({ bold: true, fillColor: PIVOT_HEADER_FILL, bTop: PIVOT_BORDER }),
    blank: 0,
  };
  let maxC = 0;
  g.cells.forEach((row, i) => {
    maxC = Math.max(maxC, row.length);
    row.forEach((v, j) => {
      const kind = g.kinds[i][j];
      let s = kindStyle[kind] ?? 0;
      const ind = g.indent[i]?.[j];
      if (ind) s = styles.merge(s, { indent: ind, hAlign: 'left' });
      const nf = g.numFmts[i]?.[j];
      if (nf) s = styles.merge(s, { numFmt: nf });
      const cell: Cell = {};
      if (v !== null && v !== undefined) cell.v = v;
      if (s) cell.s = s;
      tx.setCell(host, r0 + i, c0 + j, Object.keys(cell).length ? cell : undefined);
    });
  });
  const lastRange: Range = { r1: r0, c1: c0, r2: r0 + Math.max(0, g.cells.length - 1), c2: c0 + Math.max(0, maxC - 1) };
  // Excel autofits pivot columns on every update
  tx.flush();
  const widths = new Map(host.colWidths);
  for (let c = lastRange.c1; c <= lastRange.c2; c++) {
    const w = measureColumn(host, c, (r) => r >= lastRange.r1 && r <= lastRange.r2);
    widths.set(c, Math.max(w, 64));
  }
  tx.setMeta(host, 'colWidths', widths);
  return { ...spec, lastRange };
}

export function createPivot(sourceSheetId: string, sourceRange: Range, target: { sheetId: string | null; r: number; c: number }): void {
  const wb = S().wb;
  let hostId = target.sheetId;
  transact('Create PivotTable', (tx) => {
    let host: Sheet;
    if (!hostId) {
      host = new Sheet(wb.uniqueSheetName());
      const idx = Math.max(0, wb.sheets.findIndex((s) => s.id === sourceSheetId));
      const prev = wb.activeSheetId;
      tx.run(
        new FnCommand(
          'Insert Sheet',
          (w: Workbook) => {
            w.insertSheet(host, idx);
            w.activeSheetId = host.id;
          },
          (w: Workbook) => {
            w.removeSheet(host);
            w.activeSheetId = prev;
          },
        ),
      );
      hostId = host.id;
    } else host = wb.sheetById(hostId)!;
    const n = wb.sheets.flatMap((s) => s.pivots).length + 1;
    let spec: PivotSpec = {
      id: newId('pv'),
      name: `PivotTable${n}`,
      sourceSheetId,
      sourceRange,
      target: { r: target.r, c: target.c },
      rows: [],
      cols: [],
      values: [],
      filters: [],
      calculatedFields: [],
      grouping: [],
      showGrandTotals: true,
    };
    spec = writePivot(tx, wb, spec, host);
    tx.setMeta(host, 'pivots', host.pivots.concat([spec]));
    if (wb.activeSheetId !== host.id) wb.activeSheetId = host.id;
    setState({ pivotPanel: spec.id });
  }, { ranges: [{ r1: target.r, c1: target.c, r2: target.r, c2: target.c }], active: { r: target.r, c: target.c }, anchor: { r: target.r, c: target.c } });
}

export function findPivot(id: string): { sheet: Sheet; spec: PivotSpec } | null {
  for (const s of S().wb.sheets) {
    const p = s.pivots.find((x) => x.id === id);
    if (p) return { sheet: s, spec: p };
  }
  return null;
}

export function pivotAt(sheet: Sheet, r: number, c: number): PivotSpec | undefined {
  return sheet.pivots.find((p) => {
    const rg = p.lastRange ?? { r1: p.target.r, c1: p.target.c, r2: p.target.r + 2, c2: p.target.c + 2 };
    return r >= rg.r1 && r <= rg.r2 + 1 && c >= rg.c1 && c <= rg.c2 + 1;
  });
}

/** Update a pivot's layout (fields/areas/etc.) and rewrite its output. */
export function updatePivot(id: string, patch: Partial<PivotSpec>, label = 'PivotTable'): void {
  const found = findPivot(id);
  if (!found) return;
  const wb = S().wb;
  transact(label, (tx) => {
    const next = writePivot(tx, wb, { ...found.spec, ...patch }, found.sheet);
    tx.setMeta(found.sheet, 'pivots', found.sheet.pivots.map((p) => (p.id === id ? next : p)));
  });
}

export function refreshPivot(id: string): void {
  updatePivot(id, {}, 'Refresh');
}

export function refreshAllPivots(): void {
  for (const s of S().wb.sheets) for (const p of s.pivots) refreshPivot(p.id);
}

export function deletePivot(id: string): void {
  const found = findPivot(id);
  if (!found) return;
  transact('Delete PivotTable', (tx) => {
    if (found.spec.lastRange) found.sheet.forEachInRange(found.spec.lastRange, (r, c) => tx.setCell(found.sheet, r, c, undefined));
    tx.setMeta(found.sheet, 'pivots', found.sheet.pivots.filter((p) => p.id !== id));
  });
  setState({ pivotPanel: null });
}

/** Distinct labels of a field (for filter pickers). */
export function fieldItems(spec: PivotSpec, fieldName: string): string[] {
  const wb = S().wb;
  const src = wb.sheetById(spec.sourceSheetId);
  if (!src) return [];
  const f = pivotFields(wb, spec).find((x) => x.name === fieldName);
  if (!f) return [];
  const set = new Map<string, number | string>();
  for (let r = spec.sourceRange.r1 + 1; r <= spec.sourceRange.r2; r++) {
    const l = labelOf(getScalar(src, r, f.col), f, spec, wb, src, r);
    set.set(l.label, l.sort);
  }
  return [...set.entries()].sort((a, b) => (typeof a[1] === 'number' && typeof b[1] === 'number' ? a[1] - b[1] : String(a[1]).localeCompare(String(b[1])))).map((x) => x[0]);
}

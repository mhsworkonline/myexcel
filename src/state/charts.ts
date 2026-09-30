// Chart data and actions.
import { isFullCols, quoteSheetName, Range, rangeToA1 } from '../model/address';
import { primaryRange } from '../model/selection';
import { newId, Sheet } from '../model/sheet';
import type { CellValue, ChartSeries, ChartSpec, ChartType } from '../model/types';
import { sheetLayout } from '../model/layout';
import { dataRangeForCommand } from './actions/sortFilter';
import { alertBox, S, setState, transact } from './store';
import { rangeValues, resolveRef } from './sparkline';
import { displayOf, getScalar } from './values';

export const PALETTES: string[][] = [
  ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E', '#636363', '#997300'],
  ['#5B9BD5', '#A5A5A5', '#4472C4', '#264478', '#8FAADC', '#C9C9C9', '#1F4E79', '#2E75B6'],
  ['#ED7D31', '#F4B183', '#C55A11', '#F8CBAD', '#843C0C', '#FBE5D6'],
  ['#70AD47', '#A9D18E', '#548235', '#C5E0B4', '#375623', '#E2F0D9'],
  ['#FFC000', '#FFD966', '#BF9000', '#FFE699', '#7F6000', '#FFF2CC'],
  ['#7030A0', '#B4A7D6', '#5B2C86', '#D9D2E9', '#3F1D5C', '#E4DFEC'],
];

export interface SeriesData {
  name: string;
  values: (number | null)[];
  xValues?: (number | null)[];
  color: string;
  seriesType?: ChartSeries['seriesType'];
  secondary?: boolean;
}

export interface ChartData {
  categories: string[];
  series: SeriesData[];
}

function toNum(v: CellValue): number | null {
  return typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : null;
}

function labelText(sheet: Sheet, r: number, c: number): string {
  const st = S().wb.styles.get(sheet.styleIdAt(r, c));
  const d = displayOf(sheet, r, c, st, 20);
  return d.left !== undefined ? (d.left + (d.right ?? '')).trim() : d.text;
}

function rangeLabels(ref: string | undefined, host: Sheet): string[] {
  if (!ref) return [];
  const res = resolveRef(S().wb, host, ref);
  if (!res) return [];
  const out: string[] = [];
  const { sheet, range } = res;
  for (let r = range.r1; r <= Math.min(range.r2, range.r1 + 5000); r++) for (let c = range.c1; c <= Math.min(range.c2, range.c1 + 500); c++) out.push(labelText(sheet, r, c));
  return out;
}

export function chartData(spec: ChartSpec, host: Sheet): ChartData {
  const wb = S().wb;
  const palette = spec.colors ?? PALETTES[(spec.style - 1) % PALETTES.length] ?? PALETTES[0];
  const series: SeriesData[] = spec.series.map((s, i) => {
    const res = resolveRef(wb, host, s.values);
    const values = res ? rangeValues(res.sheet, res.range).map(toNum) : [];
    let xValues: (number | null)[] | undefined;
    if (s.xValues) {
      const xr = resolveRef(wb, host, s.xValues);
      xValues = xr ? rangeValues(xr.sheet, xr.range).map(toNum) : undefined;
    }
    let name = s.name ?? `Series${i + 1}`;
    if (s.nameRef) {
      const nr = resolveRef(wb, host, s.nameRef);
      if (nr) name = labelText(nr.sheet, nr.range.r1, nr.range.c1) || name;
    }
    return { name, values, xValues, color: s.color ?? palette[i % palette.length], seriesType: s.seriesType, secondary: s.secondaryAxis };
  });
  let categories = rangeLabels(spec.series[0]?.categories, host);
  const n = Math.max(0, ...series.map((s) => s.values.length));
  if (!categories.length) categories = Array.from({ length: n }, (_, i) => String(i + 1));
  return { categories, series };
}

function refOf(sheet: Sheet, rg: Range): string {
  return `${quoteSheetName(sheet.name)}!${rangeToA1(rg, true)}`;
}

/** Build series definitions from a data range, like Excel's chart wizard. */
export function seriesFromRange(sheet: Sheet, rg: Range, type: ChartType, byRow?: boolean): { series: ChartSeries[]; title?: string } {
  const isText = (r: number, c: number) => {
    const v = getScalar(sheet, r, c);
    return typeof v === 'string' || v === null;
  };
  // header row: first row has text in the value columns
  let hasHeaderRow = false;
  for (let c = rg.c1; c <= rg.c2; c++) if (typeof getScalar(sheet, rg.r1, c) === 'string') hasHeaderRow = true;
  if (rg.r1 === rg.r2) hasHeaderRow = false;
  // category column: first column all text (below header)
  let hasCatCol = rg.c2 > rg.c1;
  for (let r = rg.r1 + (hasHeaderRow ? 1 : 0); r <= rg.r2 && hasCatCol; r++) if (!isText(r, rg.c1)) hasCatCol = false;
  if (type === 'scatter') hasCatCol = rg.c2 > rg.c1;
  const dataR1 = rg.r1 + (hasHeaderRow ? 1 : 0);
  const dataC1 = rg.c1 + (hasCatCol ? 1 : 0);
  const rows = rg.r2 - dataR1 + 1;
  const cols = rg.c2 - dataC1 + 1;
  const series: ChartSeries[] = [];
  const asRows = byRow ?? (cols > rows && type !== 'scatter' && type !== 'pie' && type !== 'doughnut');
  if (!asRows) {
    const cats = hasCatCol ? refOf(sheet, { r1: dataR1, r2: rg.r2, c1: rg.c1, c2: rg.c1 }) : undefined;
    for (let c = dataC1; c <= rg.c2; c++) {
      const s: ChartSeries = { values: refOf(sheet, { r1: dataR1, r2: rg.r2, c1: c, c2: c }), categories: type === 'scatter' ? undefined : cats };
      if (type === 'scatter' && cats) s.xValues = cats;
      if (hasHeaderRow) s.nameRef = refOf(sheet, { r1: rg.r1, r2: rg.r1, c1: c, c2: c });
      series.push(s);
    }
  } else {
    const cats = hasHeaderRow ? refOf(sheet, { r1: rg.r1, r2: rg.r1, c1: dataC1, c2: rg.c2 }) : undefined;
    for (let r = dataR1; r <= rg.r2; r++) {
      const s: ChartSeries = { values: refOf(sheet, { r1: r, r2: r, c1: dataC1, c2: rg.c2 }), categories: cats };
      if (hasCatCol) s.nameRef = refOf(sheet, { r1: r, r2: r, c1: rg.c1, c2: rg.c1 });
      series.push(s);
    }
  }
  if (type === 'pie' || type === 'doughnut') series.splice(1);
  if (type === 'combo' && series.length > 1) series.forEach((s, i) => (s.seriesType = i === series.length - 1 ? 'line' : 'column'));
  const title = series.length === 1 && hasHeaderRow ? String(getScalar(sheet, rg.r1, asRows ? rg.c1 : dataC1) ?? '') : undefined;
  return { series, title: title || undefined };
}

export function chartSourceRange(): { sheet: Sheet; range: Range } | null {
  const st = S();
  const sheet = st.wb.activeSheet;
  let rg = primaryRange(st.sel);
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) {
    const cr = dataRangeForCommand();
    if (!cr) return null;
    rg = cr;
  } else if (isFullCols(rg)) {
    const used = sheet.usedRange();
    if (!used) return null;
    rg = { ...rg, r1: used.r1, r2: used.r2 };
  }
  return { sheet, range: rg };
}

export function insertChart(type: ChartType): void {
  const src = chartSourceRange();
  if (!src) {
    alertBox('Select a range of cells with data before inserting a chart.', 'MyExcel', 'info');
    return;
  }
  const { sheet, range } = src;
  const { series, title } = seriesFromRange(sheet, range, type);
  if (!series.length) {
    alertBox('The selected range has no numeric data to chart.', 'MyExcel', 'info');
    return;
  }
  const { rows, cols } = sheetLayout(sheet);
  const x = cols.offset(range.c2 + 2);
  const y = rows.offset(range.r1);
  const spec: ChartSpec = {
    id: newId('ch'),
    type,
    title: title ?? (series.length === 1 ? undefined : 'Chart Title'),
    showTitle: true,
    legend: type === 'pie' || type === 'doughnut' ? 'right' : series.length > 1 ? 'bottom' : 'none',
    dataLabels: false,
    gridlines: type !== 'pie' && type !== 'doughnut',
    style: 1,
    series,
    sourceRange: refOf(sheet, range),
    anchor: { x, y, w: 480, h: 288 },
  };
  transact('Insert Chart', (tx) => tx.setMeta(sheet, 'charts', sheet.charts.concat([spec])));
  setState({ selectedChartId: spec.id, ribbonTab: 'Chart Design' as never });
}

export function updateChart(id: string, patch: Partial<ChartSpec>, label = 'Chart'): void {
  const sheet = S().wb.activeSheet;
  if (!sheet.charts.some((c) => c.id === id)) return;
  transact(label, (tx) => tx.setMeta(sheet, 'charts', sheet.charts.map((c) => (c.id === id ? { ...c, ...patch } : c))));
}

export function deleteChart(id: string): void {
  const sheet = S().wb.activeSheet;
  transact('Delete Chart', (tx) => tx.setMeta(sheet, 'charts', sheet.charts.filter((c) => c.id !== id)));
  setState({ selectedChartId: null });
}

export function selectedChart(): ChartSpec | undefined {
  const st = S();
  return st.selectedChartId ? st.wb.activeSheet.charts.find((c) => c.id === st.selectedChartId) : undefined;
}

export function switchRowCol(id: string): void {
  const sheet = S().wb.activeSheet;
  const ch = sheet.charts.find((c) => c.id === id);
  if (!ch?.sourceRange) return;
  const res = resolveRef(S().wb, sheet, ch.sourceRange);
  if (!res) return;
  const firstVals = resolveRef(S().wb, sheet, ch.series[0]?.values ?? '');
  const currentlyByRow = !!firstVals && firstVals.range.r1 === firstVals.range.r2 && ch.series.length > 1;
  const { series } = seriesFromRange(res.sheet, res.range, ch.type, !currentlyByRow);
  updateChart(id, { series }, 'Switch Row/Column');
}

export function changeChartType(id: string, type: ChartType): void {
  const ch = S().wb.activeSheet.charts.find((c) => c.id === id);
  if (!ch) return;
  let series = ch.series.map((s) => ({ ...s, seriesType: undefined as ChartSeries['seriesType'] }));
  if (type === 'combo' && series.length > 1) series = series.map((s, i) => ({ ...s, seriesType: i === series.length - 1 ? 'line' : 'column' }));
  updateChart(id, { type, series, gridlines: type !== 'pie' && type !== 'doughnut' }, 'Change Chart Type');
}

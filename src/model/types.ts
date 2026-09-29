import type { Range } from './address';

export type CellValue = number | string | boolean | null;

export interface Note {
  text: string;
  author?: string;
  /** Threaded comment (Review > New Comment) vs legacy note. */
  threaded?: boolean;
  replies?: { author: string; text: string; time: number }[];
  time?: number;
}

export interface Cell {
  /** Constant value, or the cached value for formula cells loaded from files. */
  v?: CellValue;
  /** Formula text including leading '='. */
  f?: string;
  /** Style id in the workbook StyleTable. */
  s?: number;
  /** v is an error literal such as '#N/A'. */
  e?: boolean;
  note?: Note;
  link?: string;
}

// ---------- Conditional formatting ----------
export interface CFStyle {
  fontColor?: string;
  fillColor?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  numFmt?: string;
  borderColor?: string;
}

export interface CFVO {
  type: 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula' | 'autoMin' | 'autoMax';
  value?: number | string;
  gte?: boolean;
}

export type CFType =
  | 'cellIs'
  | 'containsText'
  | 'notContainsText'
  | 'beginsWith'
  | 'endsWith'
  | 'timePeriod'
  | 'top10'
  | 'aboveAverage'
  | 'duplicateValues'
  | 'uniqueValues'
  | 'expression'
  | 'dataBar'
  | 'colorScale'
  | 'iconSet'
  | 'containsBlanks'
  | 'notContainsBlanks'
  | 'containsErrors'
  | 'notContainsErrors';

export type CFOperator =
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'greaterThan'
  | 'lessThan'
  | 'greaterThanOrEqual'
  | 'lessThanOrEqual';

export interface CFRule {
  id: string;
  ranges: Range[];
  priority: number;
  stopIfTrue?: boolean;
  type: CFType;
  operator?: CFOperator;
  formulas?: string[];
  text?: string;
  timePeriod?: 'today' | 'yesterday' | 'tomorrow' | 'last7Days' | 'thisWeek' | 'lastWeek' | 'nextWeek' | 'thisMonth' | 'lastMonth' | 'nextMonth';
  rank?: number;
  percent?: boolean;
  bottom?: boolean;
  aboveAverage?: boolean;
  equalAverage?: boolean;
  stdDev?: number;
  style?: CFStyle;
  dataBar?: { color: string; gradient: boolean; min: CFVO; max: CFVO; showValue: boolean; negColor?: string };
  colorScale?: { cfvos: CFVO[]; colors: string[] };
  iconSet?: { set: string; cfvos: CFVO[]; reverse?: boolean; showValue?: boolean };
}

// ---------- Data validation ----------
export type DVType = 'any' | 'whole' | 'decimal' | 'list' | 'date' | 'time' | 'textLength' | 'custom';

export interface DataValidation {
  id: string;
  ranges: Range[];
  type: DVType;
  operator?: CFOperator;
  formula1?: string;
  formula2?: string;
  allowBlank: boolean;
  showDropdown: boolean;
  showInput: boolean;
  promptTitle?: string;
  prompt?: string;
  showError: boolean;
  errorStyle: 'stop' | 'warning' | 'information';
  errorTitle?: string;
  error?: string;
}

// ---------- Tables ----------
export type TotalFn = 'none' | 'sum' | 'average' | 'count' | 'countNums' | 'max' | 'min' | 'stdDev' | 'var' | 'custom';

export interface TableColumn {
  name: string;
  totalFunction?: TotalFn;
  totalLabel?: string;
}

export interface TableDef {
  id: string;
  name: string;
  range: Range;
  headerRow: boolean;
  totalRow: boolean;
  style: string;
  bandedRows: boolean;
  bandedCols: boolean;
  firstCol: boolean;
  lastCol: boolean;
  showFilterButton: boolean;
  columns: TableColumn[];
}

// ---------- AutoFilter ----------
export type FilterOp = 'eq' | 'ne' | 'gt' | 'gte' | 'lt' | 'lte' | 'beginsWith' | 'endsWith' | 'contains' | 'notContains';

export type ColumnFilter =
  | { type: 'values'; values: string[]; blanks: boolean }
  | { type: 'custom'; and: boolean; c1: { op: FilterOp; val: string }; c2?: { op: FilterOp; val: string } }
  | { type: 'color'; color: string; font?: boolean }
  | { type: 'top10'; top: boolean; percent: boolean; n: number }
  | { type: 'dynamic'; kind: 'aboveAverage' | 'belowAverage' };

export interface AutoFilter {
  range: Range;
  /** keyed by absolute column index */
  filters: Record<number, ColumnFilter>;
  sort?: { col: number; desc: boolean };
}

// ---------- Protection ----------
export interface SheetProtection {
  passwordHash?: string;
  allow: {
    selectLocked: boolean;
    selectUnlocked: boolean;
    formatCells: boolean;
    formatColumns: boolean;
    formatRows: boolean;
    insertColumns: boolean;
    insertRows: boolean;
    insertHyperlinks: boolean;
    deleteColumns: boolean;
    deleteRows: boolean;
    sort: boolean;
    autoFilter: boolean;
    pivotTables: boolean;
    editObjects: boolean;
    editScenarios: boolean;
  };
}

// ---------- Charts ----------
export type ChartType = 'column' | 'stackedColumn' | 'bar' | 'stackedBar' | 'line' | 'lineMarkers' | 'pie' | 'doughnut' | 'area' | 'stackedArea' | 'scatter' | 'combo';

export interface ChartSeries {
  name?: string;
  /** e.g. Sheet1!$B$2:$B$10 */
  values: string;
  categories?: string;
  /** Scatter X values. */
  xValues?: string;
  color?: string;
  /** For combo charts. */
  seriesType?: 'column' | 'line' | 'area';
  secondaryAxis?: boolean;
  nameRef?: string;
}

export interface ChartSpec {
  id: string;
  type: ChartType;
  title?: string;
  showTitle: boolean;
  legend: 'right' | 'bottom' | 'top' | 'left' | 'none';
  dataLabels: boolean;
  gridlines: boolean;
  xTitle?: string;
  yTitle?: string;
  style: number;
  series: ChartSeries[];
  sourceRange?: string;
  anchor: { x: number; y: number; w: number; h: number }; // px relative to sheet origin at 100% zoom
  colors?: string[];
}

// ---------- Pivot tables ----------
export type PivotAgg = 'sum' | 'count' | 'average' | 'max' | 'min' | 'product' | 'countNums' | 'stdDev' | 'var';

export interface PivotSpec {
  id: string;
  name: string;
  sourceSheetId: string;
  sourceRange: Range;
  target: { r: number; c: number };
  rows: string[];
  cols: string[];
  values: { field: string; agg: PivotAgg; name?: string }[];
  filters: { field: string; selected: string[] | null }[];
  calculatedFields: { name: string; formula: string }[];
  grouping: { field: string; type: 'date' | 'number'; by?: 'years' | 'quarters' | 'months' | 'days'; start?: number; end?: number; interval?: number }[];
  showGrandTotals: boolean;
  lastRange?: Range;
}

// ---------- Sparklines ----------
export interface SparklineGroup {
  id: string;
  type: 'line' | 'column' | 'winloss';
  /** Each item maps a location cell to a data range reference (A1, possibly sheet-qualified). */
  items: { r: number; c: number; ref: string }[];
  color: string;
  negColor?: string;
  markers?: boolean;
  highPoint?: boolean;
  lowPoint?: boolean;
}

// ---------- Page setup ----------
export interface PageSetup {
  orientation: 'portrait' | 'landscape';
  paperSize: 'letter' | 'legal' | 'a4' | 'a3' | 'tabloid';
  scale: number;
  fitToPage: boolean;
  fitToWidth: number;
  fitToHeight: number;
  margins: { top: number; bottom: number; left: number; right: number; header: number; footer: number };
  header: { left: string; center: string; right: string };
  footer: { left: string; center: string; right: string };
  printArea?: Range;
  printTitleRows?: [number, number];
  printTitleCols?: [number, number];
  gridlines: boolean;
  headings: boolean;
  centerH: boolean;
  centerV: boolean;
  pageOrder: 'downThenOver' | 'overThenDown';
}

export interface Scenario {
  id: string;
  name: string;
  cells: { r: number; c: number }[];
  values: CellValue[];
  comment?: string;
}

export interface DefinedName {
  name: string;
  /** Formula text without '=' e.g. Sheet1!$A$1:$B$4 */
  ref: string;
  /** undefined = workbook scope; otherwise sheet id */
  scope?: string;
  comment?: string;
}

export function defaultPageSetup(): PageSetup {
  return {
    orientation: 'portrait',
    paperSize: 'letter',
    scale: 100,
    fitToPage: false,
    fitToWidth: 1,
    fitToHeight: 1,
    margins: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7, header: 0.3, footer: 0.3 },
    header: { left: '', center: '', right: '' },
    footer: { left: '', center: '', right: '' },
    gridlines: false,
    headings: false,
    centerH: false,
    centerV: false,
    pageOrder: 'downThenOver',
  };
}

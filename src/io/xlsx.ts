// .xlsx import/export via ExcelJS.
import ExcelJS from 'exceljs';
import { addrToA1, colToName, parseA1, parseRange, Range, rangeToA1 } from '../model/address';
import { partsToSerial } from '../model/numfmt';
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT, newId, Sheet } from '../model/sheet';
import { BorderEdge, BorderStyleName, CellStyle, DEFAULT_FONT, DEFAULT_FONT_SIZE, HAlign, Underline, VAlign } from '../model/styles';
import type { Cell, CFRule, CFVO, DataValidation, DefinedName, TableDef } from '../model/types';
import { Workbook } from '../model/workbook';

export interface ValueReader {
  /** Computed value of a formula cell for writing cached results. */
  (sheet: Sheet, r: number, c: number): unknown;
}

// ---------- unit conversions ----------

export function charsToPx(w: number): number {
  return Math.trunc(((256 * w + Math.trunc(128 / 7)) / 256) * 7);
}
export function pxToChars(px: number): number {
  return Math.round((px / 7) * 256) / 256;
}
export const ptToPxH = (pt: number) => Math.round((pt * 4) / 3);
export const pxToPtH = (px: number) => Math.round(px * 0.75 * 4) / 4;

// ---------- colors ----------

const THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
const INDEXED = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  r /= 255;
  g /= 255;
  b /= 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  let h = 0;
  let s = 0;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h /= 6;
  }
  return [h, s, l];
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const hue = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [hue(p, q, h + 1 / 3) * 255, hue(p, q, h) * 255, hue(p, q, h - 1 / 3) * 255];
}

export function applyTint(hex: string, tint: number): string {
  if (!tint) return hex;
  const n = parseInt(hex, 16);
  const [h, s, l] = rgbToHsl((n >> 16) & 255, (n >> 8) & 255, n & 255);
  const nl = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const [r, g, b] = hslToRgb(h, s, Math.max(0, Math.min(1, nl)));
  return [r, g, b].map((v) => Math.round(v).toString(16).padStart(2, '0')).join('').toUpperCase();
}

function readColor(c: Partial<ExcelJS.Color> & { indexed?: number; tint?: number } | undefined): string | undefined {
  if (!c) return undefined;
  if (c.argb) {
    const a = c.argb.length === 8 ? c.argb.slice(2) : c.argb;
    return '#' + a.toUpperCase();
  }
  if (c.theme !== undefined) {
    const base = THEME[c.theme] ?? '000000';
    return '#' + applyTint(base, c.tint ?? 0);
  }
  if (c.indexed !== undefined) {
    if (c.indexed === 64) return undefined; // system foreground
    return '#' + (INDEXED[c.indexed] ?? '000000');
  }
  return undefined;
}

function argb(hex: string | undefined): { argb: string } | undefined {
  if (!hex) return undefined;
  return { argb: 'FF' + hex.replace('#', '').toUpperCase() };
}

// ---------- style conversion ----------

function readBorder(b: Partial<ExcelJS.Border> | undefined): BorderEdge | undefined {
  if (!b || !b.style) return undefined;
  return { style: b.style as BorderStyleName, color: readColor(b.color) ?? '#000000' };
}

function readStyle(cell: ExcelJS.Cell | ExcelJS.Row | ExcelJS.Column): CellStyle {
  const s: CellStyle = {};
  const f = cell.font;
  if (f) {
    if (f.name && f.name !== DEFAULT_FONT) s.fontName = f.name;
    if (f.size && f.size !== DEFAULT_FONT_SIZE) s.fontSize = f.size;
    if (f.bold) s.bold = true;
    if (f.italic) s.italic = true;
    if (f.underline) s.underline = (f.underline === true ? 'single' : f.underline) as Underline;
    if (f.strike) s.strike = true;
    const col = readColor(f.color as ExcelJS.Color);
    if (col && col !== '#000000') s.fontColor = col;
    if (f.vertAlign === 'superscript' || f.vertAlign === 'subscript') s.vertAlign = f.vertAlign;
  }
  const fill = cell.fill as (ExcelJS.FillPattern | { type: 'gradient' }) | undefined;
  if (fill && fill.type === 'pattern' && fill.pattern && fill.pattern !== ('none' as string)) {
    const col = readColor(fill.fgColor as ExcelJS.Color) ?? readColor(fill.bgColor as ExcelJS.Color);
    if (col) s.fillColor = col;
    if (fill.pattern !== 'solid') {
      s.patternType = fill.pattern;
      const pc = readColor(fill.bgColor as ExcelJS.Color);
      if (pc) s.patternColor = pc;
    }
  } else if (fill && fill.type === 'gradient') {
    const g = fill as unknown as ExcelJS.FillGradientAngle;
    const col = readColor(g.stops?.[0]?.color as ExcelJS.Color);
    if (col) s.fillColor = col;
  }
  const b = cell.border;
  if (b) {
    const t = readBorder(b.top);
    const bo = readBorder(b.bottom);
    const l = readBorder(b.left);
    const r = readBorder(b.right);
    if (t) s.bTop = t;
    if (bo) s.bBottom = bo;
    if (l) s.bLeft = l;
    if (r) s.bRight = r;
    const d = readBorder(b.diagonal);
    if (d && b.diagonal?.up) s.bDiagUp = d;
    if (d && b.diagonal?.down) s.bDiagDown = d;
  }
  const a = cell.alignment;
  if (a) {
    if (a.horizontal && a.horizontal !== 'general' as unknown) s.hAlign = a.horizontal as HAlign;
    if (a.vertical && a.vertical !== 'bottom') s.vAlign = a.vertical as VAlign;
    if (a.wrapText) s.wrap = true;
    if (a.shrinkToFit) s.shrink = true;
    if (a.indent) s.indent = a.indent;
    if (a.textRotation) s.rotation = a.textRotation === 'vertical' ? 255 : (a.textRotation as number);
  }
  if (cell.numFmt && cell.numFmt !== 'General') s.numFmt = cell.numFmt;
  const p = cell.protection;
  if (p) {
    if (p.locked === false) s.locked = false;
    if (p.hidden) s.hidden = true;
  }
  return s;
}

let writeDocFont = DEFAULT_FONT;

function writeStyle(target: ExcelJS.Cell | ExcelJS.Column | ExcelJS.Row, s: CellStyle): void {
  const font: Partial<ExcelJS.Font> = {};
  if (s.fontName) font.name = s.fontName;
  else if (writeDocFont !== DEFAULT_FONT) font.name = writeDocFont;
  if (s.fontSize) font.size = s.fontSize;
  if (s.bold) font.bold = true;
  if (s.italic) font.italic = true;
  if (s.underline) font.underline = s.underline === 'single' ? true : s.underline;
  if (s.strike) font.strike = true;
  if (s.fontColor) font.color = argb(s.fontColor);
  if (s.vertAlign) font.vertAlign = s.vertAlign;
  if (Object.keys(font).length) target.font = { name: writeDocFont, size: DEFAULT_FONT_SIZE, ...font } as ExcelJS.Font;
  if (s.fillColor) {
    target.fill = s.patternType
      ? { type: 'pattern', pattern: s.patternType as ExcelJS.FillPatterns, fgColor: argb(s.fillColor), bgColor: argb(s.patternColor ?? '#FFFFFF') }
      : { type: 'pattern', pattern: 'solid', fgColor: argb(s.fillColor), bgColor: { indexed: 64 } as unknown as ExcelJS.Color };
  }
  const edge = (e?: BorderEdge): Partial<ExcelJS.Border> | undefined => (e ? { style: e.style as ExcelJS.BorderStyle, color: argb(e.color ?? '#000000') } : undefined);
  if (s.bTop || s.bBottom || s.bLeft || s.bRight || s.bDiagUp || s.bDiagDown) {
    const border: Partial<ExcelJS.Borders> = {};
    if (s.bTop) border.top = edge(s.bTop);
    if (s.bBottom) border.bottom = edge(s.bBottom);
    if (s.bLeft) border.left = edge(s.bLeft);
    if (s.bRight) border.right = edge(s.bRight);
    const diag = s.bDiagUp ?? s.bDiagDown;
    if (diag) border.diagonal = { ...edge(diag), up: !!s.bDiagUp, down: !!s.bDiagDown } as ExcelJS.BorderDiagonal;
    target.border = border as ExcelJS.Borders;
  }
  const al: Partial<ExcelJS.Alignment> = {};
  if (s.hAlign) al.horizontal = s.hAlign as ExcelJS.Alignment['horizontal'];
  if (s.vAlign) al.vertical = s.vAlign as ExcelJS.Alignment['vertical'];
  if (s.wrap) al.wrapText = true;
  if (s.shrink) al.shrinkToFit = true;
  if (s.indent) al.indent = s.indent;
  if (s.rotation) al.textRotation = s.rotation === 255 ? 'vertical' : s.rotation;
  if (Object.keys(al).length) target.alignment = al;
  if (s.numFmt) target.numFmt = s.numFmt;
  if (s.locked === false || s.hidden) target.protection = { locked: s.locked !== false, hidden: !!s.hidden };
}

// ---------- reading ----------

function dateToSerial(d: Date): number {
  const ms = d.getTime() - Date.UTC(1899, 11, 30);
  return +(ms / 86400000).toFixed(10);
}

function compressAddresses(addrs: string[]): Range[] {
  const cells = addrs.map((a) => parseA1(a)).filter((x): x is { r: number; c: number } => !!x);
  cells.sort((a, b) => a.r - b.r || a.c - b.c);
  const runs: Range[] = [];
  for (const { r, c } of cells) {
    const last = runs[runs.length - 1];
    if (last && last.r1 === r && last.c2 === c - 1) last.c2 = c;
    else runs.push({ r1: r, r2: r, c1: c, c2: c });
  }
  const out: Range[] = [];
  for (const run of runs) {
    const prev = out.find((o) => o.r2 === run.r1 - 1 && o.c1 === run.c1 && o.c2 === run.c2);
    if (prev) prev.r2 = run.r2;
    else out.push(run);
  }
  return out;
}

function parseSqref(s: string): Range[] {
  return s
    .split(/\s+/)
    .map((x) => parseRange(x.replace(/\$/g, '')))
    .filter((x): x is Range => !!x);
}

function readCfvo(v: { type: string; value?: number | string } | undefined): CFVO {
  if (!v) return { type: 'min' };
  return { type: v.type as CFVO['type'], value: v.value };
}

function readCF(ws: ExcelJS.Worksheet, dxfStyle: (st: Partial<ExcelJS.Style> | undefined) => CFRule['style']): CFRule[] {
  const out: CFRule[] = [];
  const list = (ws as unknown as { conditionalFormattings?: { ref: string; rules: Record<string, unknown>[] }[] }).conditionalFormattings ?? [];
  for (const cf of list) {
    const ranges = parseSqref(cf.ref);
    for (const rule of cf.rules) {
      const type = rule.type as string;
      const base: CFRule = { id: newId('cf'), ranges, priority: (rule.priority as number) ?? out.length + 1, type: 'expression' };
      if (rule.stopIfTrue) base.stopIfTrue = true;
      const style = dxfStyle(rule.style as Partial<ExcelJS.Style>);
      switch (type) {
        case 'expression':
          out.push({ ...base, type: 'expression', formulas: ((rule.formulae as string[]) ?? []).map((f) => '=' + String(f)), style });
          break;
        case 'cellIs':
          out.push({ ...base, type: 'cellIs', operator: rule.operator as CFRule['operator'], formulas: ((rule.formulae as unknown[]) ?? []).map((f) => '=' + String(f)), style });
          break;
        case 'top10':
          out.push({ ...base, type: 'top10', rank: (rule.rank as number) ?? 10, percent: !!rule.percent, bottom: !!rule.bottom, style });
          break;
        case 'aboveAverage':
          out.push({ ...base, type: 'aboveAverage', aboveAverage: rule.aboveAverage !== false, style });
          break;
        case 'containsText': {
          const op = (rule.operator as string) ?? 'containsText';
          const formula = ((rule.formulae as string[]) ?? [])[0];
          let text = rule.text as string | undefined;
          if (!text && formula) text = /"([^"]*)"/.exec(formula)?.[1];
          if (op === 'containsBlanks' || op === 'notContainsBlanks' || op === 'containsErrors' || op === 'notContainsErrors') out.push({ ...base, type: op as CFRule['type'], style });
          else out.push({ ...base, type: op === 'containsText' ? 'containsText' : (op as CFRule['type']), text, style });
          break;
        }
        case 'timePeriod':
          out.push({ ...base, type: 'timePeriod', timePeriod: rule.timePeriod as CFRule['timePeriod'], style });
          break;
        case 'dataBar': {
          const cfvo = (rule.cfvo as { type: string; value?: number }[]) ?? [];
          out.push({
            ...base,
            type: 'dataBar',
            dataBar: { color: readColor(rule.color as ExcelJS.Color) ?? '#638EC6', gradient: rule.gradient !== false, min: readCfvo(cfvo[0]), max: readCfvo(cfvo[1]), showValue: rule.showValue !== false },
          });
          break;
        }
        case 'colorScale': {
          const cfvo = (rule.cfvo as { type: string; value?: number }[]) ?? [];
          const colors = ((rule.color as ExcelJS.Color[]) ?? []).map((c) => readColor(c) ?? '#FFFFFF');
          out.push({ ...base, type: 'colorScale', colorScale: { cfvos: cfvo.map(readCfvo), colors } });
          break;
        }
        case 'iconSet': {
          const cfvo = (rule.cfvo as { type: string; value?: number }[]) ?? [];
          out.push({ ...base, type: 'iconSet', iconSet: { set: (rule.iconSet as string) ?? '3TrafficLights1', cfvos: cfvo.map(readCfvo), reverse: !!rule.reverse, showValue: rule.showValue !== false } });
          break;
        }
        default:
          break;
      }
    }
  }
  return out;
}

function readDxf(st: Partial<ExcelJS.Style> | undefined): CFRule['style'] {
  if (!st) return {};
  const out: NonNullable<CFRule['style']> = {};
  if (st.font) {
    const col = readColor(st.font.color as ExcelJS.Color);
    if (col) out.fontColor = col;
    if (st.font.bold) out.bold = true;
    if (st.font.italic) out.italic = true;
    if (st.font.underline) out.underline = true;
    if (st.font.strike) out.strike = true;
  }
  const fill = st.fill as ExcelJS.FillPattern | undefined;
  if (fill && fill.type === 'pattern') {
    const col = readColor(fill.bgColor as ExcelJS.Color) ?? readColor(fill.fgColor as ExcelJS.Color);
    if (col) out.fillColor = col;
  }
  if (st.numFmt) out.numFmt = st.numFmt;
  if (st.border) {
    const col = readColor((st.border.top ?? st.border.bottom)?.color as ExcelJS.Color);
    if (col) out.borderColor = col;
  }
  return out;
}

export async function readXlsx(data: ArrayBuffer): Promise<Workbook> {
  const x = new ExcelJS.Workbook();
  await x.xlsx.load(data);
  const wb = new Workbook();
  const styles = wb.styles;
  for (const ws of x.worksheets) {
    const sheet = new Sheet(ws.name);
    sheet.visibility = ws.state === 'hidden' ? 'hidden' : ws.state === 'veryHidden' ? 'veryHidden' : 'visible';
    const tab = (ws.properties as { tabColor?: ExcelJS.Color }).tabColor;
    if (tab) sheet.tabColor = readColor(tab);
    const view = ws.views?.[0] as (Partial<ExcelJS.WorksheetView> & { xSplit?: number; ySplit?: number; showGridLines?: boolean; zoomScale?: number; showRowColHeaders?: boolean; rightToLeft?: boolean }) | undefined;
    if (view) {
      if (view.state === 'frozen') sheet.freeze = { rows: view.ySplit ?? 0, cols: view.xSplit ?? 0 };
      if (view.showGridLines === false) sheet.showGridlines = false;
      if (view.showRowColHeaders === false) sheet.showHeaders = false;
      if (view.zoomScale) sheet.zoom = view.zoomScale;
      if (view.rightToLeft) sheet.rtl = true;
    }
    // sheet-wide default sizes (sheetFormatPr)
    const fmtPr = ws.properties as { defaultRowHeight?: number; defaultColWidth?: number } | undefined;
    if (fmtPr?.defaultRowHeight) sheet.defaultRowHeight = ptToPxH(fmtPr.defaultRowHeight);
    if (fmtPr?.defaultColWidth) sheet.defaultColWidth = charsToPx(fmtPr.defaultColWidth);
    // columns
    const colCount = ws.columnCount;
    for (let c = 1; c <= Math.max(colCount, ws.columns?.length ?? 0); c++) {
      const col = ws.getColumn(c);
      if (col.width !== undefined && col.width !== null) {
        const px = charsToPx(col.width);
        if (px !== sheet.defaultColWidth) sheet.colWidths.set(c - 1, px);
      }
      if (col.hidden) sheet.hiddenCols.add(c - 1);
      const cs = readStyle(col);
      if (Object.keys(cs).length) sheet.colStyles.set(c - 1, styles.intern(cs));
    }
    ws.eachRow({ includeEmpty: true }, (row, rn) => {
      const r = rn - 1;
      if (row.height && ptToPxH(row.height) !== sheet.defaultRowHeight && (row as unknown as { customHeight?: boolean }).customHeight !== false) sheet.rowHeights.set(r, ptToPxH(row.height));
      if (row.hidden) sheet.hiddenRows.add(r);
      const rstyle = (row as unknown as { _style?: unknown; style?: unknown }).style ? readStyle(row) : {};
      if (Object.keys(rstyle).length && (row as unknown as { model?: { style?: unknown } }).model?.style) sheet.rowStyles.set(r, styles.intern(rstyle));
      row.eachCell({ includeEmpty: true }, (cell, cn) => {
        const c = cn - 1;
        if (cell.isMerged && cell.master !== cell) {
          const st = readStyle(cell);
          if (Object.keys(st).length) sheet.setCellRaw(r, c, { s: styles.intern(st) });
          return;
        }
        const out: Cell = {};
        const st = readStyle(cell);
        if (Object.keys(st).length) out.s = styles.intern(st);
        const v = cell.value as unknown;
        const setVal = (val: unknown) => {
          if (val === null || val === undefined) return;
          if (val instanceof Date) out.v = dateToSerial(val);
          else if (typeof val === 'number' || typeof val === 'boolean' || typeof val === 'string') out.v = val;
          else if (typeof val === 'object' && 'error' in (val as object)) {
            out.v = (val as { error: string }).error;
            out.e = true;
          } else if (typeof val === 'object' && 'richText' in (val as object)) out.v = (val as { richText: { text: string }[] }).richText.map((t) => t.text).join('');
        };
        if (v && typeof v === 'object' && !(v instanceof Date)) {
          const o = v as Record<string, unknown>;
          if ('formula' in o || 'sharedFormula' in o) {
            const f = cell.formula;
            if (f) out.f = '=' + f;
            setVal(o.result);
          } else if ('hyperlink' in o) {
            const text = o.text as unknown;
            setVal(typeof text === 'object' && text && 'richText' in (text as object) ? (text as { richText: { text: string }[] }).richText.map((t) => t.text).join('') : text);
            out.link = String(o.hyperlink);
          } else setVal(v);
        } else setVal(v);
        const note = cell.note as unknown;
        if (note) {
          const text = typeof note === 'string' ? note : ((note as { texts?: { text: string }[] }).texts ?? []).map((t) => t.text).join('');
          const m = /^([^:\n]{1,60}):\n/.exec(text);
          out.note = m ? { text: text.slice(m[0].length), author: m[1] } : { text };
        }
        if (Object.keys(out).length) sheet.setCellRaw(r, c, out);
      });
    });
    // merges
    const merges = (ws.model as unknown as { merges?: string[] }).merges ?? [];
    sheet.merges = merges.map((m) => parseRange(m)).filter((x): x is Range => !!x);
    // autofilter
    const af = ws.autoFilter as unknown;
    if (af) {
      let rg: Range | null = null;
      if (typeof af === 'string') rg = parseRange(af.replace(/\$/g, ''));
      else if (typeof af === 'object') {
        const o = af as { from: unknown; to: unknown };
        const s = (p: unknown) => (typeof p === 'string' ? p : `${colToName(((p as { column: number }).column ?? 1) - 1)}${(p as { row: number }).row}`);
        rg = parseRange(`${s(o.from)}:${s(o.to)}`.replace(/\$/g, ''));
      }
      if (rg) sheet.autoFilter = { range: rg, filters: {} };
    }
    // data validations
    const dvModel = ((ws as unknown as { dataValidations: unknown }).dataValidations as unknown as { model: Record<string, ExcelJS.DataValidation | undefined> }).model ?? {};
    const groups = new Map<ExcelJS.DataValidation, string[]>();
    for (const [addr, dv] of Object.entries(dvModel)) {
      if (!dv) continue;
      const g = groups.get(dv);
      if (g) g.push(addr);
      else groups.set(dv, [addr]);
    }
    for (const [dv, addrs] of groups) {
      const ranges: Range[] = [];
      const singles: string[] = [];
      for (const a of addrs) {
        if (a.includes(':')) {
          const rg = parseRange(a.replace(/\$/g, ''));
          if (rg) ranges.push(rg);
        } else singles.push(a);
      }
      ranges.push(...compressAddresses(singles));
      const f = (x: unknown) => (x === undefined || x === null ? undefined : typeof x === 'string' && (dv.type === 'list' && /^".*"$/.test(x)) ? x : '=' + String(x));
      const f1 = dv.formulae?.[0];
      sheet.validations.push({
        id: newId('dv'),
        ranges,
        type: (dv.type as DataValidation['type']) ?? 'any',
        operator: dv.operator as DataValidation['operator'],
        formula1: dv.type === 'list' && typeof f1 === 'string' && /^".*"$/.test(f1) ? f1.slice(1, -1) : f(f1 instanceof Date ? dateToSerial(f1) : f1),
        formula2: f(dv.formulae?.[1] instanceof Date ? dateToSerial(dv.formulae[1] as Date) : dv.formulae?.[1]),
        allowBlank: dv.allowBlank !== false,
        showDropdown: dv.type === 'list' ? !(dv as unknown as { showDropDown?: boolean }).showDropDown : false,
        showInput: dv.showInputMessage !== false && !!(dv.prompt || dv.promptTitle),
        promptTitle: dv.promptTitle,
        prompt: dv.prompt,
        showError: dv.showErrorMessage !== false,
        errorStyle: (dv.errorStyle as DataValidation['errorStyle']) ?? 'stop',
        errorTitle: dv.errorTitle,
        error: dv.error,
      });
    }
    // conditional formats
    sheet.conditionalFormats = readCF(ws, readDxf);
    // tables
    const tables = (ws as unknown as { tables?: Record<string, { table?: { name: string; displayName?: string; ref?: string; tableRef?: string; headerRow?: boolean; totalsRow?: boolean; style?: { theme?: string; showRowStripes?: boolean; showColumnStripes?: boolean; showFirstColumn?: boolean; showLastColumn?: boolean }; columns?: { name: string; totalsRowFunction?: string; totalsRowLabel?: string }[] } }> }).tables ?? {};
    for (const t of Object.values(tables)) {
      const m = t.table ?? (t as unknown as NonNullable<typeof t.table>);
      const ref = m.tableRef ?? m.ref;
      if (!ref) continue;
      let rg = parseRange(ref.replace(/\$/g, ''));
      if (!rg) continue;
      const cols = m.columns ?? [];
      if (rg.r1 === rg.r2 && rg.c1 === rg.c2 && cols.length > 1) rg = { ...rg, c2: rg.c1 + cols.length - 1 };
      sheet.tables.push({
        id: newId('tb'),
        name: m.displayName ?? m.name,
        range: rg,
        headerRow: m.headerRow !== false,
        totalRow: !!m.totalsRow,
        style: m.style?.theme ?? 'TableStyleMedium2',
        bandedRows: m.style?.showRowStripes !== false,
        bandedCols: !!m.style?.showColumnStripes,
        firstCol: !!m.style?.showFirstColumn,
        lastCol: !!m.style?.showLastColumn,
        showFilterButton: true,
        columns: cols.map((c) => ({ name: c.name, totalFunction: (c.totalsRowFunction as TableDef['columns'][0]['totalFunction']) ?? 'none', totalLabel: c.totalsRowLabel })),
      });
    }
    // protection
    const prot = (ws as unknown as { sheetProtection?: Record<string, unknown> }).sheetProtection;
    if (prot && prot.sheet) {
      const allow = (k: string, def: boolean) => (prot[k] === undefined ? def : !prot[k]);
      sheet.protection = {
        passwordHash: (prot.password as string | undefined) ?? (prot.hashValue as string | undefined),
        allow: {
          selectLocked: prot.selectLockedCells !== true,
          selectUnlocked: prot.selectUnlockedCells !== true,
          formatCells: allow('formatCells', false),
          formatColumns: allow('formatColumns', false),
          formatRows: allow('formatRows', false),
          insertColumns: allow('insertColumns', false),
          insertRows: allow('insertRows', false),
          insertHyperlinks: allow('insertHyperlinks', false),
          deleteColumns: allow('deleteColumns', false),
          deleteRows: allow('deleteRows', false),
          sort: allow('sort', false),
          autoFilter: allow('autoFilter', false),
          pivotTables: allow('pivotTables', false),
          editObjects: allow('objects', false),
          editScenarios: allow('scenarios', false),
        },
      };
    }
    // page setup
    const ps = ws.pageSetup as Partial<ExcelJS.PageSetup> & { printArea?: string; printTitlesRow?: string };
    if (ps) {
      sheet.pageSetup.orientation = ps.orientation === 'landscape' ? 'landscape' : 'portrait';
      if (ps.scale) sheet.pageSetup.scale = ps.scale;
      if (ps.fitToPage) {
        sheet.pageSetup.fitToPage = true;
        sheet.pageSetup.fitToWidth = ps.fitToWidth ?? 1;
        sheet.pageSetup.fitToHeight = ps.fitToHeight ?? 1;
      }
      const paper: Record<number, typeof sheet.pageSetup.paperSize> = { 1: 'letter', 5: 'legal', 9: 'a4', 8: 'a3', 3: 'tabloid' };
      if (ps.paperSize && paper[ps.paperSize]) sheet.pageSetup.paperSize = paper[ps.paperSize];
      if (ps.margins) sheet.pageSetup.margins = { ...sheet.pageSetup.margins, ...ps.margins };
      if (ps.printArea) {
        const rg = parseRange(ps.printArea.split(',')[0].replace(/\$/g, '').replace(/^.*!/, ''));
        if (rg) sheet.pageSetup.printArea = rg;
      }
      if (ps.printTitlesRow) {
        const m = /(\d+):\$?(\d+)/.exec(ps.printTitlesRow);
        if (m) sheet.pageSetup.printTitleRows = [+m[1] - 1, +m[2] - 1];
      }
      if (ps.showGridLines) sheet.pageSetup.gridlines = true;
      if (ps.horizontalCentered) sheet.pageSetup.centerH = true;
      if (ps.verticalCentered) sheet.pageSetup.centerV = true;
    }
    const hf = ws.headerFooter as Partial<ExcelJS.HeaderFooter> | undefined;
    if (hf) {
      sheet.pageSetup.header = parseHF(hf.oddHeader);
      sheet.pageSetup.footer = parseHF(hf.oddFooter);
    }
    const breaks = (ws as unknown as { rowBreaks?: { id: number }[] }).rowBreaks;
    if (breaks?.length) sheet.rowBreaks = breaks.map((b) => b.id);
    sheet.touch();
    wb.sheets.push(sheet);
  }
  // defined names
  const names: DefinedName[] = [];
  const dn = (x.definedNames as unknown as { model: { name: string; ranges: string[]; localSheetId?: number }[] }).model ?? [];
  for (const d of dn) {
    if (d.name.startsWith('_xlnm')) {
      if (d.name === '_xlnm.Print_Area' || d.name === '_xlnm._FilterDatabase') continue;
      continue;
    }
    const scope = d.localSheetId !== undefined ? wb.sheets[d.localSheetId]?.id : undefined;
    names.push({ name: d.name, ref: d.ranges.join(','), scope });
  }
  wb.names = names;
  wb.activeSheetId = wb.sheets[(x.views?.[0] as { activeTab?: number } | undefined)?.activeTab ?? 0]?.id ?? wb.sheets[0]?.id;
  if (!wb.sheets.length) {
    const s = new Sheet('Sheet1');
    wb.sheets.push(s);
    wb.activeSheetId = s.id;
  }
  const act = wb.sheetById(wb.activeSheetId);
  if (!act || act.visibility !== 'visible') wb.activeSheetId = wb.sheets.find((s) => s.visibility === 'visible')!.id;
  wb.props = { title: x.title, author: x.creator, created: x.created?.getTime(), modified: x.modified?.getTime() };
  return wb;
}

function parseHF(s: string | undefined): { left: string; center: string; right: string } {
  const out = { left: '', center: '', right: '' };
  if (!s) return out;
  const re = /&([LCR])/g;
  let cur = 'center' as 'left' | 'center' | 'right';
  let last = 0;
  let m: RegExpExecArray | null;
  const parts: [typeof cur, string][] = [];
  while ((m = re.exec(s))) {
    parts.push([cur, s.slice(last, m.index)]);
    cur = m[1] === 'L' ? 'left' : m[1] === 'C' ? 'center' : 'right';
    last = m.index + 2;
  }
  parts.push([cur, s.slice(last)]);
  for (const [k, v] of parts) if (v) out[k] += v;
  return out;
}

function buildHF(h: { left: string; center: string; right: string }): string | undefined {
  const s = (h.left ? '&L' + h.left : '') + (h.center ? '&C' + h.center : '') + (h.right ? '&R' + h.right : '');
  return s || undefined;
}

// ---------- writing ----------

function cfToExcel(cf: CFRule, tl: string): Record<string, unknown> | null {
  const dxf = (s: CFRule['style']): Partial<ExcelJS.Style> | undefined => {
    if (!s) return undefined;
    const st: Partial<ExcelJS.Style> = {};
    const font: Partial<ExcelJS.Font> = {};
    if (s.fontColor) font.color = argb(s.fontColor);
    if (s.bold) font.bold = true;
    if (s.italic) font.italic = true;
    if (s.underline) font.underline = true;
    if (s.strike) font.strike = true;
    if (Object.keys(font).length) st.font = font;
    if (s.fillColor) st.fill = { type: 'pattern', pattern: 'solid', bgColor: argb(s.fillColor) };
    if (s.numFmt) st.numFmt = s.numFmt;
    if (s.borderColor) {
      const e = { style: 'thin' as const, color: argb(s.borderColor) };
      st.border = { top: e, bottom: e, left: e, right: e };
    }
    return st;
  };
  const stripEq = (f: string) => (f.startsWith('=') ? f.slice(1) : f);
  const base = { priority: cf.priority, stopIfTrue: cf.stopIfTrue };
  switch (cf.type) {
    case 'cellIs':
      return { ...base, type: 'cellIs', operator: cf.operator, formulae: (cf.formulas ?? []).map(stripEq), style: dxf(cf.style) };
    case 'expression':
      return { ...base, type: 'expression', formulae: (cf.formulas ?? []).map(stripEq), style: dxf(cf.style) };
    case 'containsText':
      return { ...base, type: 'containsText', operator: 'containsText', text: cf.text, formulae: [`NOT(ISERROR(SEARCH("${cf.text ?? ''}",${tl})))`], style: dxf(cf.style) };
    case 'notContainsText':
      return { ...base, type: 'expression', formulae: [`ISERROR(SEARCH("${cf.text ?? ''}",${tl}))`], style: dxf(cf.style) };
    case 'beginsWith':
      return { ...base, type: 'expression', formulae: [`LEFT(${tl},LEN("${cf.text ?? ''}"))="${cf.text ?? ''}"`], style: dxf(cf.style) };
    case 'endsWith':
      return { ...base, type: 'expression', formulae: [`RIGHT(${tl},LEN("${cf.text ?? ''}"))="${cf.text ?? ''}"`], style: dxf(cf.style) };
    case 'containsBlanks':
    case 'notContainsBlanks':
    case 'containsErrors':
    case 'notContainsErrors':
      return { ...base, type: 'containsText', operator: cf.type, style: dxf(cf.style) };
    case 'timePeriod':
      return { ...base, type: 'timePeriod', timePeriod: cf.timePeriod, style: dxf(cf.style) };
    case 'top10':
      return { ...base, type: 'top10', rank: cf.rank ?? 10, percent: !!cf.percent, bottom: !!cf.bottom, style: dxf(cf.style) };
    case 'aboveAverage':
      return { ...base, type: 'aboveAverage', aboveAverage: cf.aboveAverage !== false, style: dxf(cf.style) };
    case 'duplicateValues':
    case 'uniqueValues': {
      const rg = cf.ranges[0];
      const abs = `$${colToName(rg.c1)}$${rg.r1 + 1}:$${colToName(rg.c2)}$${rg.r2 + 1}`;
      return { ...base, type: 'expression', formulae: [`COUNTIF(${abs},${tl})${cf.type === 'duplicateValues' ? '>1' : '=1'}`], style: dxf(cf.style) };
    }
    case 'dataBar':
      return {
        ...base,
        type: 'dataBar',
        cfvo: [cf.dataBar!.min, cf.dataBar!.max].map((v) => ({ type: v.type === 'autoMin' ? 'min' : v.type === 'autoMax' ? 'max' : v.type, value: v.value })),
        color: argb(cf.dataBar!.color),
        gradient: cf.dataBar!.gradient,
        showValue: cf.dataBar!.showValue,
      };
    case 'colorScale':
      return { ...base, type: 'colorScale', cfvo: cf.colorScale!.cfvos.map((v) => ({ type: v.type, value: v.value })), color: cf.colorScale!.colors.map(argb) };
    case 'iconSet':
      return { ...base, type: 'iconSet', iconSet: cf.iconSet!.set, cfvo: cf.iconSet!.cfvos.map((v) => ({ type: v.type, value: v.value })), reverse: cf.iconSet!.reverse, showValue: cf.iconSet!.showValue };
  }
  return null;
}

export async function writeXlsx(wb: Workbook, getValue: ValueReader): Promise<ArrayBuffer> {
  const x = new ExcelJS.Workbook();
  writeDocFont = wb.props.defaultFont ?? DEFAULT_FONT;
  x.creator = wb.props.author ?? 'MyExcel';
  x.created = wb.props.created ? new Date(wb.props.created) : new Date();
  x.modified = new Date();
  const styles = wb.styles;
  const activeIdx = Math.max(0, wb.sheets.findIndex((s) => s.id === wb.activeSheetId));
  x.views = [{ x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, activeTab: activeIdx, visibility: 'visible' }];
  for (const sheet of wb.sheets) {
    const views: Partial<ExcelJS.WorksheetView>[] = [];
    const v: Record<string, unknown> = { showGridLines: sheet.showGridlines, zoomScale: sheet.zoom, showRowColHeaders: sheet.showHeaders, rightToLeft: sheet.rtl };
    if (sheet.freeze.rows || sheet.freeze.cols) Object.assign(v, { state: 'frozen', xSplit: sheet.freeze.cols, ySplit: sheet.freeze.rows, topLeftCell: addrToA1(sheet.freeze.rows, sheet.freeze.cols) });
    views.push(v as Partial<ExcelJS.WorksheetView>);
    const ws = x.addWorksheet(sheet.name, {
      properties: { tabColor: sheet.tabColor ? argb(sheet.tabColor) : undefined, defaultRowHeight: pxToPtH(sheet.defaultRowHeight), defaultColWidth: sheet.defaultColWidth !== DEFAULT_COL_WIDTH ? pxToChars(sheet.defaultColWidth) : undefined } as Partial<ExcelJS.WorksheetProperties>,
      views: views as ExcelJS.WorksheetView[],
      state: sheet.visibility,
    });
    // columns
    const colIdx = new Set<number>([...sheet.colWidths.keys(), ...sheet.hiddenCols, ...sheet.colStyles.keys()]);
    for (const c of [...colIdx].sort((a, b) => a - b)) {
      if (c > 16383) continue;
      const col = ws.getColumn(c + 1);
      const w = sheet.colWidths.get(c);
      if (w !== undefined) col.width = pxToChars(w);
      if (sheet.hiddenCols.has(c)) col.hidden = true;
      const cs = sheet.colStyles.get(c);
      if (cs) writeStyle(col, styles.get(cs));
    }
    const rowIdx = new Set<number>([...sheet.rows.keys(), ...sheet.rowHeights.keys(), ...sheet.hiddenRows, ...sheet.rowStyles.keys()]);
    for (const r of [...rowIdx].sort((a, b) => a - b)) {
      const row = ws.getRow(r + 1);
      const h = sheet.rowHeights.get(r);
      if (h !== undefined) row.height = pxToPtH(h);
      if (sheet.hiddenRows.has(r)) {
        row.hidden = true;
        if (h === undefined) row.height = 15; // ExcelJS drops rows with neither cells nor height
      }
      const rs = sheet.rowStyles.get(r);
      if (rs) writeStyle(row, styles.get(rs));
      const cells = sheet.rows.get(r);
      if (!cells) continue;
      for (const [c, cell] of cells) {
        const xc = row.getCell(c + 1);
        if (cell.f) {
          const res = getValue(sheet, r, c);
          let result: unknown = res;
          if (res && typeof res === 'object' && 'error' in (res as object)) result = { error: (res as { error: string }).error === '#CYCLE!' ? '#REF!' : (res as { error: string }).error };
          xc.value = { formula: cell.f.slice(1), result: result as ExcelJS.CellFormulaValue['result'] } as ExcelJS.CellFormulaValue;
        } else if (cell.v !== undefined && cell.v !== null) {
          if (cell.e) xc.value = { error: cell.v as ExcelJS.ValueType } as unknown as ExcelJS.CellErrorValue;
          else if (cell.link) xc.value = { text: String(cell.v), hyperlink: cell.link };
          else xc.value = cell.v;
        } else if (cell.link) xc.value = { text: cell.link, hyperlink: cell.link };
        if (cell.s) writeStyle(xc, styles.get(cell.s));
        else if (writeDocFont !== DEFAULT_FONT) writeStyle(xc, {});
        if (cell.note) xc.note = (cell.note.author ? `${cell.note.author}:\n` : '') + cell.note.text;
      }
    }
    for (const m of sheet.merges) ws.mergeCells(m.r1 + 1, m.c1 + 1, m.r2 + 1, m.c2 + 1);
    if (sheet.autoFilter && !sheet.tables.some((t) => t.range.r1 === sheet.autoFilter!.range.r1 && t.range.c1 === sheet.autoFilter!.range.c1)) ws.autoFilter = rangeToA1(sheet.autoFilter.range);
    // validations
    const dvs = ((ws as unknown as { dataValidations: unknown }).dataValidations as unknown as { model: Record<string, unknown> }).model;
    for (const dv of sheet.validations) {
      if (dv.type === 'any') continue;
      const f = (s: string | undefined) => (s === undefined ? undefined : s.startsWith('=') ? s.slice(1) : dv.type === 'list' ? `"${s}"` : s);
      const model = {
        type: dv.type,
        operator: dv.type === 'list' || dv.type === 'custom' ? undefined : dv.operator ?? 'between',
        allowBlank: dv.allowBlank,
        showInputMessage: dv.showInput,
        showErrorMessage: dv.showError,
        promptTitle: dv.promptTitle,
        prompt: dv.prompt,
        errorStyle: dv.errorStyle,
        errorTitle: dv.errorTitle,
        error: dv.error,
        formulae: [f(dv.formula1), f(dv.formula2)].filter((q) => q !== undefined),
        showDropDown: dv.type === 'list' ? !dv.showDropdown : undefined,
      };
      for (const rg of dv.ranges) dvs[rangeToA1(rg)] = model;
    }
    // conditional formats
    const byRef = new Map<string, Record<string, unknown>[]>();
    for (const cf of [...sheet.conditionalFormats].sort((a, b) => a.priority - b.priority)) {
      const ref = cf.ranges.map((rg) => rangeToA1(rg)).join(' ');
      const tl = addrToA1(cf.ranges[0].r1, cf.ranges[0].c1);
      const rule = cfToExcel(cf, tl);
      if (!rule) continue;
      const l = byRef.get(ref) ?? [];
      l.push(rule);
      byRef.set(ref, l);
    }
    for (const [ref, rules] of byRef) ws.addConditionalFormatting({ ref, rules: rules as unknown as ExcelJS.ConditionalFormattingRule[] });
    // tables
    for (const t of sheet.tables) {
      const rows: unknown[][] = [];
      const dataStart = t.range.r1 + (t.headerRow ? 1 : 0);
      const dataEnd = t.range.r2 - (t.totalRow ? 1 : 0);
      for (let r = dataStart; r <= dataEnd; r++) {
        const row: unknown[] = [];
        for (let c = t.range.c1; c <= t.range.c2; c++) {
          const cell = sheet.getCell(r, c);
          if (cell?.f) {
            const res = getValue(sheet, r, c);
            row.push({ formula: cell.f.slice(1), result: res && typeof res === 'object' ? undefined : res });
          } else row.push(cell?.v ?? null);
        }
        rows.push(row);
      }
      if (!rows.length) rows.push(t.columns.map(() => null));
      try {
        ws.addTable({
          name: t.name.replace(/[^A-Za-z0-9_]/g, '_'),
          displayName: t.name.replace(/[^A-Za-z0-9_]/g, '_'),
          ref: addrToA1(t.range.r1, t.range.c1),
          headerRow: t.headerRow,
          totalsRow: t.totalRow,
          style: { theme: t.style as ExcelJS.TableStyleProperties['theme'], showRowStripes: t.bandedRows, showColumnStripes: t.bandedCols, showFirstColumn: t.firstCol, showLastColumn: t.lastCol },
          columns: t.columns.map((c) => ({ name: c.name, filterButton: t.showFilterButton, totalsRowFunction: (c.totalFunction && c.totalFunction !== 'none' ? c.totalFunction : undefined) as ExcelJS.TableColumnProperties['totalsRowFunction'], totalsRowLabel: c.totalLabel })),
          rows,
        });
      } catch {
        /* invalid table definition – cells are still written */
      }
    }
    // re-apply cell styles in table area (addTable rewrites values but keeps styles)
    // protection
    if (sheet.protection) {
      const a = sheet.protection.allow;
      await ws.protect('', {
        selectLockedCells: a.selectLocked,
        selectUnlockedCells: a.selectUnlocked,
        formatCells: a.formatCells,
        formatColumns: a.formatColumns,
        formatRows: a.formatRows,
        insertColumns: a.insertColumns,
        insertRows: a.insertRows,
        insertHyperlinks: a.insertHyperlinks,
        deleteColumns: a.deleteColumns,
        deleteRows: a.deleteRows,
        sort: a.sort,
        autoFilter: a.autoFilter,
        pivotTables: a.pivotTables,
        objects: a.editObjects,
        scenarios: a.editScenarios,
      });
      if (sheet.protection.passwordHash) {
        const sp = (ws as unknown as { sheetProtection: Record<string, unknown> }).sheetProtection;
        sp.hashValue = sheet.protection.passwordHash;
        sp.algorithmName = 'SHA-512';
      }
    }
    // page setup
    const ps = sheet.pageSetup;
    const paper: Record<string, number> = { letter: 1, legal: 5, a4: 9, a3: 8, tabloid: 3 };
    ws.pageSetup = {
      ...ws.pageSetup,
      orientation: ps.orientation,
      paperSize: paper[ps.paperSize] as ExcelJS.PageSetup['paperSize'],
      scale: ps.scale,
      fitToPage: ps.fitToPage,
      fitToWidth: ps.fitToWidth,
      fitToHeight: ps.fitToHeight,
      margins: ps.margins,
      showGridLines: ps.gridlines,
      horizontalCentered: ps.centerH,
      verticalCentered: ps.centerV,
      printArea: ps.printArea ? rangeToA1(ps.printArea, true) : undefined,
      printTitlesRow: ps.printTitleRows ? `$${ps.printTitleRows[0] + 1}:$${ps.printTitleRows[1] + 1}` : undefined,
    };
    const hdr = buildHF(ps.header);
    const ftr = buildHF(ps.footer);
    if (hdr || ftr) ws.headerFooter = { oddHeader: hdr, oddFooter: ftr } as ExcelJS.HeaderFooter;
    for (const b of sheet.rowBreaks) ws.getRow(b + 1).addPageBreak();
  }
  for (const n of wb.names) {
    try {
      const refs = n.ref.split(',');
      for (const ref of refs) x.definedNames.add(ref, n.name);
    } catch {
      /* formulas that aren't ranges cannot be expressed via ExcelJS */
    }
  }
  const buf = await x.xlsx.writeBuffer();
  return buf as ArrayBuffer;
}

export { partsToSerial };

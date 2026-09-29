import { isFullCols, isFullRows, MAX_COLS, MAX_ROWS, Range } from '../../model/address';
import type { Tx } from '../../model/commands';
import { adjustDecimals, FMT } from '../../model/numfmt';
import { primaryRange } from '../../model/selection';
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT, Sheet } from '../../model/sheet';
import { BorderEdge, BorderStyleName, CellStyle, DEFAULT_FONT_SIZE, StylePatch } from '../../model/styles';
import { lineHeight, measureText, wrapText } from '../measure';
import { bump, openDialog, S, setState, transact } from '../store';
import { displayOf, getComputed } from '../values';
import { guardRanges, guardSheet } from './edit';

export function activeStyle(): CellStyle {
  const st = S();
  const sheet = st.wb.activeSheet;
  return st.wb.styles.get(sheet.styleIdAt(st.sel.active.r, st.sel.active.c));
}

type StyleFn = (base: CellStyle, r: number, c: number) => StylePatch | null;

/** Apply a per-cell style transform to ranges (handles whole rows/cols via row/col styles). */
export function applyStyleFn(tx: Tx, sheet: Sheet, ranges: Range[], fn: StyleFn): void {
  const wb = S().wb;
  const styles = wb.styles;
  const mergeId = (id: number, r: number, c: number) => {
    const p = fn(styles.get(id), r, c);
    return p ? styles.merge(id, p) : id;
  };
  for (const rg of ranges) {
    const fullC = isFullCols(rg);
    const fullR = isFullRows(rg);
    if (fullC || fullR) {
      if (fullC && fullR) {
        // Select All: style every column
        const cs = new Map(sheet.colStyles);
        for (let c = 0; c < MAX_COLS; c++) cs.set(c, mergeId(sheet.colStyles.get(c) ?? 0, -1, c));
        tx.setMeta(sheet, 'colStyles', cs);
        const rs = new Map(sheet.rowStyles);
        for (const [r, id] of sheet.rowStyles) rs.set(r, mergeId(id, r, -1));
        tx.setMeta(sheet, 'rowStyles', rs);
      } else if (fullC) {
        const cs = new Map(sheet.colStyles);
        for (let c = rg.c1; c <= rg.c2; c++) cs.set(c, mergeId(sheet.colStyles.get(c) ?? 0, -1, c));
        tx.setMeta(sheet, 'colStyles', cs);
        // intersections with styled rows need concrete cells
        for (const [r, rid] of sheet.rowStyles) {
          for (let c = rg.c1; c <= rg.c2; c++) {
            const cell = sheet.getCell(r, c);
            if (cell?.s !== undefined) continue;
            tx.setCell(sheet, r, c, { ...(cell ?? {}), s: mergeId(rid, r, c) });
          }
        }
      } else {
        const rs = new Map(sheet.rowStyles);
        for (let r = rg.r1; r <= rg.r2; r++) rs.set(r, mergeId(sheet.rowStyles.get(r) ?? 0, r, -1));
        tx.setMeta(sheet, 'rowStyles', rs);
        for (const [c, cid] of sheet.colStyles) {
          for (let r = rg.r1; r <= rg.r2; r++) {
            const cell = sheet.getCell(r, c);
            if (cell?.s !== undefined) continue;
            if (sheet.rowStyles.has(r)) continue;
            tx.setCell(sheet, r, c, { ...(cell ?? {}), s: mergeId(cid, r, c) });
          }
        }
      }
      // existing cells inside
      sheet.forEachInRange(rg, (r, c, cell) => {
        const id = cell.s ?? sheet.styleIdAt(r, c);
        tx.setCell(sheet, r, c, { ...cell, s: mergeId(id, r, c) });
      });
      continue;
    }
    for (let r = rg.r1; r <= rg.r2; r++) {
      for (let c = rg.c1; c <= rg.c2; c++) {
        const cell = sheet.getCell(r, c);
        const id = sheet.styleIdAt(r, c);
        const nid = mergeId(id, r, c);
        if (nid !== id || (cell?.s === undefined && nid !== 0)) tx.setCell(sheet, r, c, { ...(cell ?? {}), s: nid });
      }
    }
  }
}

export function applyStyle(patch: StylePatch, label = 'Format Cells', ranges?: Range[]): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  if (sheet.protection && !sheet.protection.allow.formatCells) {
    guardSheet(null);
    return;
  }
  transact(label, (tx) => {
    applyStyleFn(tx, sheet, ranges ?? st.sel.ranges, () => patch);
    if (patch.fontSize !== undefined || patch.wrap !== undefined || patch.fontName !== undefined) autoRowHeights(tx, sheet, ranges ?? st.sel.ranges);
  });
}

export function toggleStyle(key: 'bold' | 'italic' | 'strike'): void {
  const cur = !!activeStyle()[key];
  applyStyle({ [key]: cur ? null : true }, key[0].toUpperCase() + key.slice(1));
}

export function toggleUnderline(kind: 'single' | 'double' = 'single'): void {
  const cur = activeStyle().underline;
  applyStyle({ underline: cur === kind ? null : kind }, 'Underline');
}

const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
export { FONT_SIZES };

export function growFont(dir: 1 | -1): void {
  const cur = activeStyle().fontSize ?? DEFAULT_FONT_SIZE;
  let next: number;
  if (dir > 0) next = FONT_SIZES.find((s) => s > cur) ?? cur + 10;
  else next = [...FONT_SIZES].reverse().find((s) => s < cur) ?? Math.max(1, cur - 1);
  applyStyle({ fontSize: next }, 'Font Size');
}

// ---------- number formats ----------

export function setNumFmt(fmt: string): void {
  applyStyle({ numFmt: fmt === 'General' ? null : fmt }, 'Number Format');
}

export function changeDecimals(delta: 1 | -1): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  const { r, c } = st.sel.active;
  const v = getComputed(sheet, r, c);
  const base = activeStyle().numFmt;
  const fmt = adjustDecimals(base, delta, typeof v === 'number' ? v : undefined);
  applyStyle({ numFmt: fmt }, delta > 0 ? 'Increase Decimal' : 'Decrease Decimal');
}

export function commaStyle(): void {
  setNumFmt(FMT.comma);
}

// ---------- borders ----------

export type BorderPreset =
  | 'bottom'
  | 'top'
  | 'left'
  | 'right'
  | 'none'
  | 'all'
  | 'outside'
  | 'thickOutside'
  | 'bottomDouble'
  | 'thickBottom'
  | 'topBottom'
  | 'topThickBottom'
  | 'topDoubleBottom'
  | 'inside'
  | 'insideH'
  | 'insideV'
  | 'diagUp'
  | 'diagDown';

export function applyBorder(preset: BorderPreset, style: BorderStyleName = 'thin', color = '#000000'): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  const edge = (s: BorderStyleName): BorderEdge => ({ style: s, color });
  transact('Borders', (tx) => {
    for (const rg of st.sel.ranges) {
      const r2 = Math.min(rg.r2, rg.r1 + 20000);
      const c2 = Math.min(rg.c2, rg.c1 + 2000);
      applyStyleFn(tx, sheet, [{ ...rg, r2, c2 }], (_b, r, c) => {
        const top = r === rg.r1;
        const bottom = r === r2;
        const left = c === rg.c1;
        const right = c === c2;
        const p: StylePatch = {};
        switch (preset) {
          case 'none':
            return { bTop: null, bBottom: null, bLeft: null, bRight: null, bDiagDown: null, bDiagUp: null };
          case 'all':
            return { bTop: edge(style), bBottom: edge(style), bLeft: edge(style), bRight: edge(style) };
          case 'bottom':
            if (bottom) p.bBottom = edge(style);
            break;
          case 'top':
            if (top) p.bTop = edge(style);
            break;
          case 'left':
            if (left) p.bLeft = edge(style);
            break;
          case 'right':
            if (right) p.bRight = edge(style);
            break;
          case 'outside':
          case 'thickOutside': {
            const s = preset === 'thickOutside' ? 'medium' : style;
            if (top) p.bTop = edge(s);
            if (bottom) p.bBottom = edge(s);
            if (left) p.bLeft = edge(s);
            if (right) p.bRight = edge(s);
            break;
          }
          case 'bottomDouble':
            if (bottom) p.bBottom = edge('double');
            break;
          case 'thickBottom':
            if (bottom) p.bBottom = edge('medium');
            break;
          case 'topBottom':
            if (top) p.bTop = edge('thin');
            if (bottom) p.bBottom = edge('thin');
            break;
          case 'topThickBottom':
            if (top) p.bTop = edge('thin');
            if (bottom) p.bBottom = edge('medium');
            break;
          case 'topDoubleBottom':
            if (top) p.bTop = edge('thin');
            if (bottom) p.bBottom = edge('double');
            break;
          case 'inside':
            if (!bottom) p.bBottom = edge(style);
            if (!right) p.bRight = edge(style);
            break;
          case 'insideH':
            if (!bottom) p.bBottom = edge(style);
            break;
          case 'insideV':
            if (!right) p.bRight = edge(style);
            break;
          case 'diagUp':
            p.bDiagUp = edge(style);
            break;
          case 'diagDown':
            p.bDiagDown = edge(style);
            break;
        }
        return Object.keys(p).length ? p : null;
      });
      // Clear adjacent opposite edges when removing borders so shared edges disappear
      if (preset === 'none') {
        const clear = (range: Range, key: 'bBottom' | 'bTop' | 'bRight' | 'bLeft') =>
          applyStyleFn(tx, sheet, [range], (b) => (b[key] ? { [key]: null } : null));
        if (rg.r1 > 0) clear({ r1: rg.r1 - 1, r2: rg.r1 - 1, c1: rg.c1, c2 }, 'bBottom');
        if (r2 < MAX_ROWS - 1) clear({ r1: r2 + 1, r2: r2 + 1, c1: rg.c1, c2 }, 'bTop');
        if (rg.c1 > 0) clear({ r1: rg.r1, r2, c1: rg.c1 - 1, c2: rg.c1 - 1 }, 'bRight');
        if (c2 < MAX_COLS - 1) clear({ r1: rg.r1, r2, c1: c2 + 1, c2: c2 + 1 }, 'bLeft');
      }
    }
  });
}

// ---------- alignment ----------

export function setHAlign(h: CellStyle['hAlign']): void {
  const cur = activeStyle().hAlign;
  applyStyle({ hAlign: cur === h ? null : h }, 'Align');
}

export function setVAlign(v: CellStyle['vAlign']): void {
  applyStyle({ vAlign: v === 'bottom' ? null : v }, 'Align');
}

export function toggleWrap(): void {
  applyStyle({ wrap: activeStyle().wrap ? null : true }, 'Wrap Text');
}

export function changeIndent(delta: 1 | -1): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  transact('Indent', (tx) => {
    applyStyleFn(tx, sheet, st.sel.ranges, (b) => {
      const n = Math.max(0, Math.min(250, (b.indent ?? 0) + delta));
      const p: StylePatch = { indent: n || null };
      if (n > 0 && (!b.hAlign || b.hAlign === 'general' || b.hAlign === 'center')) p.hAlign = 'left';
      return p;
    });
  });
}

export function setRotation(deg: number | 'vertical'): void {
  applyStyle({ rotation: deg === 'vertical' ? 255 : deg || null }, 'Orientation');
}

// ---------- merge ----------

export type MergeKind = 'center' | 'across' | 'merge' | 'unmerge';

export function mergeCells(kind: MergeKind, confirmed = false): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  if (!guardRanges(sheet, st.sel.ranges)) return;
  const rg0 = primaryRange(st.sel);
  if (kind === 'unmerge' || (kind === 'center' && sheet.merges.some((m) => m.r1 === rg0.r1 && m.c1 === rg0.c1 && m.r2 === rg0.r2 && m.c2 === rg0.c2))) {
    transact('Unmerge Cells', (tx) => {
      const merges = sheet.merges.filter((m) => !st.sel.ranges.some((rg) => m.r1 >= rg.r1 && m.r2 <= rg.r2 && m.c1 >= rg.c1 && m.c2 <= rg.c2));
      tx.setMeta(sheet, 'merges', merges);
      if (kind === 'center') applyStyleFn(tx, sheet, [rg0], () => ({ hAlign: null }));
    });
    return;
  }
  const targets: Range[] = [];
  for (const rg of st.sel.ranges) {
    if (isFullCols(rg) || isFullRows(rg)) continue;
    if (kind === 'across') for (let r = rg.r1; r <= rg.r2; r++) targets.push({ r1: r, r2: r, c1: rg.c1, c2: rg.c2 });
    else targets.push(rg);
  }
  const multiValue = targets.some((t) => {
    let n = 0;
    sheet.forEachInRange(t, (_r, _c, cell) => {
      if (cell.v !== undefined && cell.v !== null && cell.v !== '') n++;
      else if (cell.f) n++;
    });
    return n > 1;
  });
  if (multiValue && !confirmed) {
    openDialog('confirm', {
      message: 'Merging cells only keeps the upper-left value and discards other values.',
      onOk: () => mergeCells(kind, true),
    });
    return;
  }
  transact('Merge Cells', (tx) => {
    let merges = sheet.merges.filter((m) => !targets.some((t) => m.r1 <= t.r2 && t.r1 <= m.r2 && m.c1 <= t.c2 && t.c1 <= m.c2));
    for (const t of targets) {
      if (t.r1 === t.r2 && t.c1 === t.c2) continue;
      sheet.forEachInRange(t, (r, c, cell) => {
        if (r === t.r1 && c === t.c1) return;
        const n = { ...cell };
        delete n.v;
        delete n.f;
        delete n.e;
        tx.setCell(sheet, r, c, n);
      });
      merges = merges.concat([t]);
    }
    tx.setMeta(sheet, 'merges', merges);
    if (kind === 'center') applyStyleFn(tx, sheet, targets, () => ({ hAlign: 'center' }));
  });
}

// ---------- cell styles gallery ----------

export interface NamedStyle {
  name: string;
  group: 'Good, Bad and Neutral' | 'Data and Model' | 'Titles and Headings' | 'Themed Cell Styles' | 'Number Format';
  patch: StylePatch;
}

const accents = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];
const tint = (hex: string, t: number) => {
  const n = parseInt(hex.slice(1), 16);
  const ch = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => Math.round(v + (255 - v) * t));
  return '#' + ch.map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
};

export const CELL_STYLES: NamedStyle[] = [
  { name: 'Normal', group: 'Good, Bad and Neutral', patch: { fontColor: null, fillColor: null, bold: null, italic: null, bTop: null, bBottom: null, bLeft: null, bRight: null, fontSize: null, numFmt: null } },
  { name: 'Bad', group: 'Good, Bad and Neutral', patch: { fillColor: '#FFC7CE', fontColor: '#9C0006' } },
  { name: 'Good', group: 'Good, Bad and Neutral', patch: { fillColor: '#C6EFCE', fontColor: '#006100' } },
  { name: 'Neutral', group: 'Good, Bad and Neutral', patch: { fillColor: '#FFEB9C', fontColor: '#9C5700' } },
  { name: 'Calculation', group: 'Data and Model', patch: { fillColor: '#F2F2F2', fontColor: '#FA7D00', bold: true, bTop: { style: 'thin', color: '#7F7F7F' }, bBottom: { style: 'thin', color: '#7F7F7F' }, bLeft: { style: 'thin', color: '#7F7F7F' }, bRight: { style: 'thin', color: '#7F7F7F' } } },
  { name: 'Check Cell', group: 'Data and Model', patch: { fillColor: '#A5A5A5', fontColor: '#FFFFFF', bold: true, bTop: { style: 'double', color: '#3F3F3F' }, bBottom: { style: 'double', color: '#3F3F3F' }, bLeft: { style: 'double', color: '#3F3F3F' }, bRight: { style: 'double', color: '#3F3F3F' } } },
  { name: 'Explanatory Text', group: 'Data and Model', patch: { fontColor: '#7F7F7F', italic: true } },
  { name: 'Input', group: 'Data and Model', patch: { fillColor: '#FFCC99', fontColor: '#3F3F76', bTop: { style: 'thin', color: '#7F7F7F' }, bBottom: { style: 'thin', color: '#7F7F7F' }, bLeft: { style: 'thin', color: '#7F7F7F' }, bRight: { style: 'thin', color: '#7F7F7F' } } },
  { name: 'Linked Cell', group: 'Data and Model', patch: { fontColor: '#FA7D00', bBottom: { style: 'double', color: '#FF8001' } } },
  { name: 'Note', group: 'Data and Model', patch: { fillColor: '#FFFFCC', bTop: { style: 'thin', color: '#B2B2B2' }, bBottom: { style: 'thin', color: '#B2B2B2' }, bLeft: { style: 'thin', color: '#B2B2B2' }, bRight: { style: 'thin', color: '#B2B2B2' } } },
  { name: 'Output', group: 'Data and Model', patch: { fillColor: '#F2F2F2', fontColor: '#3F3F3F', bold: true, bTop: { style: 'thin', color: '#3F3F3F' }, bBottom: { style: 'thin', color: '#3F3F3F' }, bLeft: { style: 'thin', color: '#3F3F3F' }, bRight: { style: 'thin', color: '#3F3F3F' } } },
  { name: 'Warning Text', group: 'Data and Model', patch: { fontColor: '#FF0000' } },
  { name: 'Heading 1', group: 'Titles and Headings', patch: { bold: true, fontSize: 15, fontColor: '#44546A', bBottom: { style: 'thick', color: '#4472C4' } } },
  { name: 'Heading 2', group: 'Titles and Headings', patch: { bold: true, fontSize: 13, fontColor: '#44546A', bBottom: { style: 'thick', color: '#A2B8E1' } } },
  { name: 'Heading 3', group: 'Titles and Headings', patch: { bold: true, fontColor: '#44546A', bBottom: { style: 'medium', color: '#8EA9DB' } } },
  { name: 'Heading 4', group: 'Titles and Headings', patch: { bold: true, fontColor: '#44546A' } },
  { name: 'Title', group: 'Titles and Headings', patch: { fontSize: 18, fontColor: '#44546A', fontName: 'Calibri Light' } },
  { name: 'Total', group: 'Titles and Headings', patch: { bold: true, bTop: { style: 'thin', color: '#4472C4' }, bBottom: { style: 'double', color: '#4472C4' } } },
  ...accents.flatMap((a, i) => [
    { name: `20% - Accent${i + 1}`, group: 'Themed Cell Styles' as const, patch: { fillColor: tint(a, 0.8) } },
    { name: `40% - Accent${i + 1}`, group: 'Themed Cell Styles' as const, patch: { fillColor: tint(a, 0.6) } },
    { name: `60% - Accent${i + 1}`, group: 'Themed Cell Styles' as const, patch: { fillColor: tint(a, 0.4), fontColor: '#FFFFFF' } },
    { name: `Accent${i + 1}`, group: 'Themed Cell Styles' as const, patch: { fillColor: a, fontColor: '#FFFFFF' } },
  ]),
  { name: 'Comma', group: 'Number Format', patch: { numFmt: FMT.comma } },
  { name: 'Comma [0]', group: 'Number Format', patch: { numFmt: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)' } },
  { name: 'Currency', group: 'Number Format', patch: { numFmt: FMT.accounting } },
  { name: 'Currency [0]', group: 'Number Format', patch: { numFmt: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)' } },
  { name: 'Percent', group: 'Number Format', patch: { numFmt: '0%' } },
];

// ---------- format painter ----------

export function startFormatPainter(sticky: boolean): void {
  const st = S();
  if (st.painter) {
    setState({ painter: null });
    return;
  }
  setState({ painter: { sheetId: st.wb.activeSheetId, range: primaryRange(st.sel), sticky } });
}

/** Paint formats from the painter source onto `target` (tiling like Excel). */
export function applyFormatPainter(target: Range): void {
  const st = S();
  const p = st.painter;
  if (!p) return;
  const src = st.wb.sheetById(p.sheetId);
  const sheet = st.wb.activeSheet;
  if (!src) return;
  const sh = p.range.r2 - p.range.r1 + 1;
  const sw = p.range.c2 - p.range.c1 + 1;
  // A single-cell click paints the size of the source
  const tgt = target.r1 === target.r2 && target.c1 === target.c2 ? { r1: target.r1, c1: target.c1, r2: target.r1 + sh - 1, c2: target.c1 + sw - 1 } : target;
  const srcStyles: number[][] = [];
  for (let i = 0; i < sh; i++) {
    srcStyles.push([]);
    for (let j = 0; j < sw; j++) srcStyles[i].push(src.styleIdAt(p.range.r1 + i, p.range.c1 + j));
  }
  const srcMerges = src.merges.filter((m) => m.r1 >= p.range.r1 && m.r2 <= p.range.r2 && m.c1 >= p.range.c1 && m.c2 <= p.range.c2);
  transact('Format Painter', (tx) => {
    for (let r = tgt.r1; r <= Math.min(tgt.r2, tgt.r1 + 100000); r++) {
      for (let c = tgt.c1; c <= Math.min(tgt.c2, tgt.c1 + 2000); c++) {
        const s = srcStyles[(r - tgt.r1) % sh][(c - tgt.c1) % sw];
        const cell = sheet.getCell(r, c);
        if ((cell?.s ?? sheet.styleIdAt(r, c)) !== s || cell?.s === undefined) tx.setCell(sheet, r, c, { ...(cell ?? {}), s });
      }
    }
    if (srcMerges.length) {
      const add = srcMerges.map((m) => ({ r1: m.r1 - p.range.r1 + tgt.r1, r2: m.r2 - p.range.r1 + tgt.r1, c1: m.c1 - p.range.c1 + tgt.c1, c2: m.c2 - p.range.c1 + tgt.c1 }));
      tx.setMeta(sheet, 'merges', sheet.merges.filter((m) => !(m.r1 <= tgt.r2 && tgt.r1 <= m.r2 && m.c1 <= tgt.c2 && tgt.c1 <= m.c2)).concat(add));
    }
  });
  if (!p.sticky) setState({ painter: null });
}

// ---------- sizes ----------

export function setColumnWidth(px: number, cols?: number[]): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  if (!guardSheet('formatColumns')) return;
  const list = cols ?? selectedCols();
  transact('Column Width', (tx) => {
    const m = new Map(sheet.colWidths);
    const hidden = new Set(sheet.hiddenCols);
    for (const c of list) {
      if (px <= 0) hidden.add(c);
      else {
        hidden.delete(c);
        if (Math.round(px) === DEFAULT_COL_WIDTH) m.delete(c);
        else m.set(c, Math.round(px));
      }
    }
    tx.setMeta(sheet, 'colWidths', m);
    if (hidden.size !== sheet.hiddenCols.size || [...hidden].some((c) => !sheet.hiddenCols.has(c))) tx.setMeta(sheet, 'hiddenCols', hidden);
  });
}

export function setRowHeight(px: number, rows?: number[]): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  if (!guardSheet('formatRows')) return;
  const list = rows ?? selectedRows();
  transact('Row Height', (tx) => {
    const m = new Map(sheet.rowHeights);
    const hidden = new Set(sheet.hiddenRows);
    for (const r of list) {
      if (px <= 0) hidden.add(r);
      else {
        hidden.delete(r);
        m.set(r, Math.round(px));
      }
    }
    tx.setMeta(sheet, 'rowHeights', m);
    if (hidden.size !== sheet.hiddenRows.size || [...hidden].some((r) => !sheet.hiddenRows.has(r))) tx.setMeta(sheet, 'hiddenRows', hidden);
  });
}

export function selectedCols(): number[] {
  const st = S();
  const out = new Set<number>();
  for (const rg of st.sel.ranges) for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + MAX_COLS); c++) out.add(c);
  return [...out];
}

export function selectedRows(): number[] {
  const st = S();
  const out = new Set<number>();
  for (const rg of st.sel.ranges) {
    const end = isFullCols(rg) ? Math.min(rg.r2, (st.wb.activeSheet.usedRange()?.r2 ?? 0)) : rg.r2;
    for (let r = rg.r1; r <= end; r++) out.add(r);
  }
  return [...out];
}

/** Best-fit width for a column (Excel AutoFit). */
export function measureColumn(sheet: Sheet, c: number, rowFilter?: (r: number) => boolean): number {
  const wb = S().wb;
  let max = 0;
  for (const [r, row] of sheet.rows) {
    if (rowFilter && !rowFilter(r)) continue;
    const cell = row.get(c);
    if (!cell || (cell.v === undefined && cell.f === undefined)) continue;
    if (sheet.mergeAt(r, c)) continue;
    const style = wb.styles.get(cell.s ?? sheet.styleIdAt(r, c));
    const d = displayOf(sheet, r, c, style, 30);
    const text = d.left !== undefined ? (d.left ?? '') + (d.right ?? '') : d.text;
    if (style.wrap) {
      for (const line of text.split('\n')) max = Math.max(max, measureText(line, style));
    } else max = Math.max(max, measureText(text, style));
    max += 0;
    if (style.indent) max = Math.max(max, measureText(text, style) + style.indent * 9);
  }
  return max ? Math.ceil(max + 8) : DEFAULT_COL_WIDTH;
}

export function autoFitColumns(cols?: number[]): void {
  const sheet = S().wb.activeSheet;
  const list = cols ?? selectedCols();
  const sel = S().sel;
  const onlySelectedRows = sel.ranges.length === 1 && !isFullCols(sel.ranges[0]) && !cols;
  const rg = sel.ranges[0];
  transact('AutoFit Column Width', (tx) => {
    const m = new Map(sheet.colWidths);
    for (const c of list.slice(0, 2000)) {
      const w = measureColumn(sheet, c, onlySelectedRows ? (r) => r >= rg.r1 && r <= rg.r2 : undefined);
      if (w === DEFAULT_COL_WIDTH) m.delete(c);
      else m.set(c, w);
    }
    tx.setMeta(sheet, 'colWidths', m);
  });
}

export function measureRow(sheet: Sheet, r: number): number {
  const wb = S().wb;
  const row = sheet.rows.get(r);
  let h = DEFAULT_ROW_HEIGHT;
  const rs = sheet.rowStyles.get(r);
  if (rs !== undefined) h = Math.max(h, lineHeight(wb.styles.get(rs)) + 4);
  if (!row) return h;
  for (const [c, cell] of row) {
    const style = wb.styles.get(cell.s ?? sheet.styleIdAt(r, c));
    const lh = lineHeight(style);
    let lines = 1;
    if (cell.v !== undefined || cell.f !== undefined) {
      if (style.wrap && !sheet.mergeAt(r, c)) {
        const d = displayOf(sheet, r, c, style);
        lines = wrapText(d.text, style, sheet.colWidth(c) - 6).length;
      } else if (typeof cell.v === 'string' && cell.v.includes('\n') && style.wrap) lines = cell.v.split('\n').length;
    }
    h = Math.max(h, lines * lh + 5);
  }
  return Math.min(409 * (4 / 3), Math.ceil(h));
}

export function autoFitRows(rows?: number[]): void {
  const sheet = S().wb.activeSheet;
  const list = rows ?? selectedRows();
  transact('AutoFit Row Height', (tx) => {
    const m = new Map(sheet.rowHeights);
    for (const r of list.slice(0, 100000)) {
      const h = measureRow(sheet, r);
      if (h === DEFAULT_ROW_HEIGHT) m.delete(r);
      else m.set(r, h);
    }
    tx.setMeta(sheet, 'rowHeights', m);
  });
}

/** Excel auto-grows row heights when fonts get larger or text wraps (unless a custom height is set). */
export function autoRowHeights(tx: Tx, sheet: Sheet, ranges: Range[]): void {
  const m = new Map(sheet.rowHeights);
  let changed = false;
  for (const rg of ranges) {
    if (isFullCols(rg)) continue;
    for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 5000); r++) {
      const h = measureRow(sheet, r);
      const cur = m.get(r) ?? DEFAULT_ROW_HEIGHT;
      if (h !== cur) {
        if (h === DEFAULT_ROW_HEIGHT) m.delete(r);
        else m.set(r, h);
        changed = true;
      }
    }
  }
  if (changed) tx.setMeta(sheet, 'rowHeights', m);
}

export function hideRowsCols(axis: 'row' | 'col', hide: boolean): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  if (!guardSheet(axis === 'row' ? 'formatRows' : 'formatColumns')) return;
  transact(hide ? 'Hide' : 'Unhide', (tx) => {
    if (axis === 'row') {
      const s = new Set(sheet.hiddenRows);
      const m = new Map(sheet.rowHeights);
      for (const rg of st.sel.ranges) {
        const end = isFullCols(rg) ? Math.max(rg.r1, ...[...s]) : rg.r2;
        for (let r = rg.r1; r <= Math.min(end, rg.r1 + 1048576); r++) {
          if (hide) s.add(r);
          else {
            s.delete(r);
            if (m.get(r) === 0) m.delete(r);
          }
        }
      }
      tx.setMeta(sheet, 'hiddenRows', s);
      if (!hide) tx.setMeta(sheet, 'rowHeights', m);
    } else {
      const s = new Set(sheet.hiddenCols);
      for (const rg of st.sel.ranges) for (let c = rg.c1; c <= rg.c2; c++) hide ? s.add(c) : s.delete(c);
      tx.setMeta(sheet, 'hiddenCols', s);
    }
  });
  bump();
}

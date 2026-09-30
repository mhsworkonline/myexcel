// Pagination, page rendering, printing and PDF export.
import { createElement } from 'react';
import { MAX_COLS, MAX_ROWS, Range } from '../model/address';
import { sheetLayout } from '../model/layout';
import type { Sheet } from '../model/sheet';
import type { PageSetup } from '../model/types';
import { singleSel } from '../model/selection';
import type { Viewport, Pane } from '../ui/grid/geometry';
import { renderGrid } from '../ui/grid/render';
import { LIGHT } from '../ui/theme';
import { fileAdapter, FILTERS } from '../io/FileAdapter';
import { chartData } from './charts';
import { bump, S, setState, transact } from './store';

export const PAPER: Record<PageSetup['paperSize'], [number, number]> = {
  letter: [8.5, 11],
  legal: [8.5, 14],
  a4: [8.27, 11.69],
  a3: [11.69, 16.54],
  tabloid: [11, 17],
};

const DPI = 96;

export interface PageInfo {
  sheet: Sheet;
  rows: [number, number];
  cols: [number, number];
  scale: number;
  index: number;
}

export function paperSizePx(ps: PageSetup): { w: number; h: number } {
  let [w, h] = PAPER[ps.paperSize] ?? PAPER.letter;
  if (ps.orientation === 'landscape') [w, h] = [h, w];
  return { w: w * DPI, h: h * DPI };
}

function printRange(sheet: Sheet): Range | null {
  const ps = sheet.pageSetup;
  if (ps.printArea) return ps.printArea;
  const used = sheet.dataRange();
  let rg = used ? { ...used, r1: 0, c1: 0 } : null;
  // include charts
  const { rows, cols } = sheetLayout(sheet);
  for (const ch of sheet.charts) {
    const r2 = rows.indexAt(ch.anchor.y + ch.anchor.h);
    const c2 = cols.indexAt(ch.anchor.x + ch.anchor.w);
    rg = rg ? { ...rg, r2: Math.max(rg.r2, r2), c2: Math.max(rg.c2, c2) } : { r1: 0, c1: 0, r2, c2 };
  }
  return rg;
}

export function paginate(sheet: Sheet, override?: Partial<PageSetup>, only?: Range): PageInfo[] {
  const ps = { ...sheet.pageSetup, ...(override ?? {}) };
  const rg = only ?? printRange(sheet);
  if (!rg) return [{ sheet, rows: [0, 0], cols: [0, 0], scale: 1, index: 0 }];
  const { rows, cols } = sheetLayout(sheet);
  const paper = paperSizePx(ps);
  const availW = paper.w - (ps.margins.left + ps.margins.right) * DPI - (ps.headings ? 30 : 0);
  const availH = paper.h - (ps.margins.top + ps.margins.bottom) * DPI - (ps.headings ? 20 : 0);
  const titleRowsH = ps.printTitleRows ? rows.offset(ps.printTitleRows[1] + 1) - rows.offset(ps.printTitleRows[0]) : 0;
  const titleColsW = ps.printTitleCols ? cols.offset(ps.printTitleCols[1] + 1) - cols.offset(ps.printTitleCols[0]) : 0;
  const totalW = cols.offset(rg.c2 + 1) - cols.offset(rg.c1);
  const totalH = rows.offset(rg.r2 + 1) - rows.offset(rg.r1);
  let scale = ps.scale / 100;
  if (ps.fitToPage) {
    const sw = ps.fitToWidth > 0 ? (availW * ps.fitToWidth) / (totalW + titleColsW * (ps.fitToWidth - 1)) : Infinity;
    const sh = ps.fitToHeight > 0 ? (availH * ps.fitToHeight) / (totalH + titleRowsH * (ps.fitToHeight - 1)) : Infinity;
    scale = Math.min(1, sw, sh);
  }
  const split = (a: number, b: number, size: (i: number) => number, avail: number, manual: Set<number>): [number, number][] => {
    const out: [number, number][] = [];
    let start = a;
    let acc = 0;
    for (let i = a; i <= b; i++) {
      const s = size(i) * scale;
      if (i > start && (acc + s > avail || manual.has(i))) {
        out.push([start, i - 1]);
        start = i;
        acc = 0;
      }
      acc += s;
    }
    out.push([start, b]);
    return out;
  };
  const rowBlocks = split(rg.r1, rg.r2, (i) => sheet.rowHeight(i), availH - titleRowsH * scale, new Set(sheet.rowBreaks.map((b) => b + 1)));
  const colBlocks = split(rg.c1, rg.c2, (i) => sheet.colWidth(i), availW - titleColsW * scale, new Set(sheet.colBreaks.map((b) => b + 1)));
  const pages: PageInfo[] = [];
  if (ps.pageOrder === 'overThenDown') {
    for (const rb of rowBlocks) for (const cb of colBlocks) pages.push({ sheet, rows: rb, cols: cb, scale, index: pages.length });
  } else for (const cb of colBlocks) for (const rb of rowBlocks) pages.push({ sheet, rows: rb, cols: cb, scale, index: pages.length });
  return pages;
}

function hfText(tpl: string, page: number, pages: number, sheet: Sheet): string {
  const d = new Date();
  return tpl
    .replace(/&\[Page\]|&P/g, String(page))
    .replace(/&\[Pages\]|&N/g, String(pages))
    .replace(/&\[Date\]|&D/g, d.toLocaleDateString())
    .replace(/&\[Time\]|&T/g, d.toLocaleTimeString())
    .replace(/&\[File\]|&F/g, S().file.name)
    .replace(/&\[Tab\]|&A/g, sheet.name)
    .replace(/&[BIU"][^&]*?/g, '')
    .replace(/&&/g, '&');
}

async function drawCharts(ctx: CanvasRenderingContext2D, page: PageInfo, originX: number, originY: number, z: number): Promise<void> {
  const sheet = page.sheet;
  if (!sheet.charts.length) return;
  const { rows, cols } = sheetLayout(sheet);
  const x0 = cols.offset(page.cols[0]);
  const y0 = rows.offset(page.rows[0]);
  const x1 = cols.offset(page.cols[1] + 1);
  const y1 = rows.offset(page.rows[1] + 1);
  const { renderToStaticMarkup } = await import('react-dom/server');
  const { ChartSvg } = await import('../ui/charts/ChartSvg');
  for (const ch of sheet.charts) {
    const a = ch.anchor;
    if (a.x > x1 || a.y > y1 || a.x + a.w < x0 || a.y + a.h < y0) continue;
    const svg = renderToStaticMarkup(createElement(ChartSvg, { spec: ch, data: chartData(ch, sheet), width: a.w, height: a.h }));
    const img = new Image();
    await new Promise<void>((res) => {
      img.onload = () => res();
      img.onerror = () => res();
      img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
    });
    ctx.save();
    ctx.beginPath();
    ctx.rect(originX, originY, (x1 - x0) * z, (y1 - y0) * z);
    ctx.clip();
    ctx.drawImage(img, originX + (a.x - x0) * z, originY + (a.y - y0) * z, a.w * z, a.h * z);
    ctx.restore();
  }
}

/** Render one page to a canvas at `k` device pixels per CSS px. */
export async function renderPage(page: PageInfo, pages: number, k = 2, override?: Partial<PageSetup>): Promise<HTMLCanvasElement> {
  const sheet = page.sheet;
  const ps = { ...sheet.pageSetup, ...(override ?? {}) };
  const paper = paperSizePx(ps);
  const canvas = document.createElement('canvas');
  canvas.width = Math.round(paper.w * k);
  canvas.height = Math.round(paper.h * k);
  const ctx = canvas.getContext('2d')!;
  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  const z = page.scale;
  const { rows, cols } = sheetLayout(sheet);
  const hw = ps.headings ? Math.round(30 * z) : 0;
  const hh = ps.headings ? Math.round(20 * z) : 0;
  const tr = ps.printTitleRows && ps.printTitleRows[1] < page.rows[0] ? ps.printTitleRows : null;
  const tc = ps.printTitleCols && ps.printTitleCols[1] < page.cols[0] ? ps.printTitleCols : null;
  const trH = tr ? (rows.offset(tr[1] + 1) - rows.offset(tr[0])) * z : 0;
  const tcW = tc ? (cols.offset(tc[1] + 1) - cols.offset(tc[0])) * z : 0;
  const bodyW = (cols.offset(page.cols[1] + 1) - cols.offset(page.cols[0])) * z;
  const bodyH = (rows.offset(page.rows[1] + 1) - rows.offset(page.rows[0])) * z;
  const colPanes: Pane[] = [];
  const rowPanes: Pane[] = [];
  if (tc) colPanes.push({ start: hw, size: tcW, content: cols.offset(tc[0]) * z, first: tc[0], last: tc[1], scrollIdx: -1 });
  colPanes.push({ start: hw + tcW, size: bodyW, content: cols.offset(page.cols[0]) * z, first: page.cols[0], last: page.cols[1], scrollIdx: 1 });
  if (tr) rowPanes.push({ start: hh, size: trH, content: rows.offset(tr[0]) * z, first: tr[0], last: tr[1], scrollIdx: -1 });
  rowPanes.push({ start: hh + trH, size: bodyH, content: rows.offset(page.rows[0]) * z, first: page.rows[0], last: page.rows[1], scrollIdx: 1 });
  const vpW = hw + tcW + bodyW;
  const vpH = hh + trH + bodyH;
  let ox = ps.margins.left * DPI;
  let oy = ps.margins.top * DPI;
  const availW = paper.w - (ps.margins.left + ps.margins.right) * DPI;
  const availH = paper.h - (ps.margins.top + ps.margins.bottom) * DPI;
  if (ps.centerH) ox += Math.max(0, (availW - vpW) / 2);
  if (ps.centerV) oy += Math.max(0, (availH - vpH) / 2);
  const vp: Viewport = { z, hw, hh, width: vpW, height: vpH, colPanes, rowPanes, rows, cols, frozen: false, split: false };
  renderGrid(ctx, {
    sheet,
    wb: S().wb,
    vp,
    pal: LIGHT,
    dark: false,
    dpr: k,
    sel: singleSel(MAX_ROWS - 1, MAX_COLS - 1),
    editing: null,
    refs: [],
    clip: null,
    fillPreview: null,
    showFormulas: false,
    pageBreaks: null,
    invalid: null,
    filterButtons: false,
    listArrow: null,
    painterRange: null,
    print: { gridlines: ps.gridlines, origin: { x: ox, y: oy } },
  });
  ctx.setTransform(k, 0, 0, k, 0, 0);
  await drawCharts(ctx, page, ox + hw + tcW, oy + hh + trH, z);
  // header / footer
  ctx.fillStyle = '#000';
  ctx.font = '11px Calibri, Carlito, Arial, sans-serif';
  ctx.textBaseline = 'alphabetic';
  const n = page.index + 1;
  const put = (sec: { left: string; center: string; right: string }, y: number) => {
    const l = hfText(sec.left, n, pages, sheet);
    const c = hfText(sec.center, n, pages, sheet);
    const r = hfText(sec.right, n, pages, sheet);
    if (l) ctx.fillText(l, ps.margins.left * DPI, y);
    if (c) ctx.fillText(c, (paper.w - ctx.measureText(c).width) / 2, y);
    if (r) ctx.fillText(r, paper.w - ps.margins.right * DPI - ctx.measureText(r).width, y);
  };
  put(ps.header, ps.margins.header * DPI + 11);
  put(ps.footer, paper.h - ps.margins.footer * DPI);
  return canvas;
}

export type PrintWhat = 'active' | 'workbook' | 'selection';

export function collectPages(what: PrintWhat, override?: Partial<PageSetup>): PageInfo[] {
  const st = S();
  if (what === 'workbook') {
    const all: PageInfo[] = [];
    for (const s of st.wb.sheets.filter((x) => x.visibility === 'visible')) all.push(...paginate(s, override));
    return all.map((p, i) => ({ ...p, index: i }));
  }
  if (what === 'selection') {
    const rg = st.sel.ranges[st.sel.ranges.length - 1];
    return paginate(st.wb.activeSheet, override, rg);
  }
  return paginate(st.wb.activeSheet, override);
}

export async function printPages(pages: PageInfo[], override?: Partial<PageSetup>): Promise<void> {
  let root = document.getElementById('print-root');
  if (!root) {
    root = document.createElement('div');
    root.id = 'print-root';
    document.body.appendChild(root);
  }
  root.innerHTML = '';
  const style = document.createElement('style');
  const ps = { ...pages[0]?.sheet.pageSetup, ...(override ?? {}) } as PageSetup;
  const paper = PAPER[ps.paperSize] ?? PAPER.letter;
  const [pw, ph] = ps.orientation === 'landscape' ? [paper[1], paper[0]] : paper;
  style.textContent = `@page { size: ${pw}in ${ph}in; margin: 0 } #print-root img { width: ${pw}in; height: ${ph}in; display: block; page-break-after: always }`;
  root.appendChild(style);
  for (const p of pages) {
    const c = await renderPage(p, pages.length, 2, override);
    const img = document.createElement('img');
    img.src = c.toDataURL('image/png');
    root.appendChild(img);
  }
  await new Promise((r) => setTimeout(r, 50));
  window.print();
}

export async function exportPdf(pages?: PageInfo[], override?: Partial<PageSetup>): Promise<void> {
  const list = pages ?? collectPages('active', override);
  const { jsPDF } = await import('jspdf');
  const first = list[0];
  const ps = { ...first.sheet.pageSetup, ...(override ?? {}) };
  const paper = PAPER[ps.paperSize] ?? PAPER.letter;
  const [pw, ph] = ps.orientation === 'landscape' ? [paper[1], paper[0]] : paper;
  const doc = new jsPDF({ unit: 'in', format: [pw, ph], orientation: ps.orientation });
  for (let i = 0; i < list.length; i++) {
    const c = await renderPage(list[i], list.length, 2.5, override);
    if (i > 0) doc.addPage([pw, ph], ps.orientation);
    doc.addImage(c.toDataURL('image/jpeg', 0.92), 'JPEG', 0, 0, pw, ph);
  }
  const bytes = new Uint8Array(doc.output('arraybuffer'));
  await fileAdapter().saveAs(bytes, `${S().file.name}.pdf`, [FILTERS.pdf]);
}

// ---- View modes ----

const viewZoom = new WeakMap<Sheet, { normal: number; pageBreak: number }>();

/** Switch views. As in Excel, Page Break Preview keeps its own zoom (60% at first) so whole pages are visible. */
export function setViewMode(mode: 'normal' | 'pageBreak' | 'pageLayout'): void {
  const st = S();
  const prev = st.viewMode;
  if (prev === mode) return;
  const sheet = st.wb.activeSheet;
  const z = viewZoom.get(sheet) ?? { normal: sheet.zoom, pageBreak: 60 };
  if (mode === 'pageBreak') {
    z.normal = sheet.zoom;
    sheet.zoom = z.pageBreak;
  } else if (prev === 'pageBreak') {
    z.pageBreak = sheet.zoom;
    sheet.zoom = z.normal;
  }
  viewZoom.set(sheet, z);
  sheet.touch();
  setState({ viewMode: mode });
  bump();
}

// ---- Manual page breaks (stored as the last row/column of a page) ----

const sortUniq = (a: number[]) => [...new Set(a)].filter((x) => x >= 0).sort((x, y) => x - y);

/** Insert Page Break: above the active row and left of the active column (only one if it's in row 1 / column A). */
export function insertPageBreak(): void {
  const sheet = S().wb.activeSheet;
  const { r, c } = S().sel.active;
  transact('Insert Page Break', (tx) => {
    if (r > 0) tx.setMeta(sheet, 'rowBreaks', sortUniq([...sheet.rowBreaks, r - 1]));
    if (c > 0) tx.setMeta(sheet, 'colBreaks', sortUniq([...sheet.colBreaks, c - 1]));
  });
}

/** Remove Page Break: the manual breaks on the active cell's top and left edges. */
export function removePageBreak(): void {
  const sheet = S().wb.activeSheet;
  const { r, c } = S().sel.active;
  transact('Remove Page Break', (tx) => {
    if (sheet.rowBreaks.includes(r - 1)) tx.setMeta(sheet, 'rowBreaks', sheet.rowBreaks.filter((b) => b !== r - 1));
    if (sheet.colBreaks.includes(c - 1)) tx.setMeta(sheet, 'colBreaks', sheet.colBreaks.filter((b) => b !== c - 1));
  });
}

export function resetPageBreaks(): void {
  const sheet = S().wb.activeSheet;
  transact('Reset All Page Breaks', (tx) => {
    tx.setMeta(sheet, 'rowBreaks', []);
    tx.setMeta(sheet, 'colBreaks', []);
  });
}

/**
 * Page Break Preview drag: the break that started page `from` now starts page `to`. The moved
 * break becomes manual; dragging it off the printed area removes it.
 */
export function movePageBreak(axis: 'row' | 'col', from: number, to: number): void {
  if (from === to) return;
  const sheet = S().wb.activeSheet;
  const pages = paginate(sheet);
  const lo = Math.min(...pages.map((p) => (axis === 'row' ? p.rows[0] : p.cols[0])));
  const hi = Math.max(...pages.map((p) => (axis === 'row' ? p.rows[1] : p.cols[1])));
  const key = axis === 'row' ? 'rowBreaks' : 'colBreaks';
  const list = sheet[key].filter((b) => b !== from - 1);
  if (to > lo && to <= hi) list.push(to - 1);
  transact('Move Page Break', (tx) => tx.setMeta(sheet, key, sortUniq(list)));
}

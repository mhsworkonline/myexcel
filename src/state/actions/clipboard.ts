import { colToName, isFullCols, isFullRows, MAX_COLS, MAX_ROWS, Range } from '../../model/address';
import { RefInfo, refToRange, refToString, shiftFormula, tokenize } from '../../model/formula';
import { primaryRange, rangeSel } from '../../model/selection';
import type { Sheet } from '../../model/sheet';
import { BorderEdge, BorderStyleName, CellStyle, StylePatch } from '../../model/styles';
import type { Cell } from '../../model/types';
import { quoteSheetName } from '../../model/address';
import { alertBox, bump, S, setState, setStatus, transact } from '../store';
import { displayOf, getComputed } from '../values';
import { guardRanges, writeInput } from './edit';
import { isErrorVal } from '../../engine/Engine';

export interface ClipPayload {
  sheetId: string;
  sheetName: string;
  range: Range;
  cells: (Cell | undefined)[][];
  styleIds: number[][];
  merges: Range[]; // relative
  colWidths: number[];
  rowHeights: number[];
  text: string;
  cut: boolean;
}

let internal: ClipPayload | null = null;

export function getInternalClip(): ClipPayload | null {
  return internal;
}

function clipRange(sheet: Sheet, rg: Range): Range {
  if (isFullCols(rg) || isFullRows(rg)) {
    const used = sheet.usedRange();
    return {
      r1: rg.r1,
      c1: rg.c1,
      r2: isFullCols(rg) ? Math.max(rg.r1, Math.min(rg.r2, used?.r2 ?? 0)) : rg.r2,
      c2: isFullRows(rg) ? Math.max(rg.c1, Math.min(rg.c2, used?.c2 ?? 0)) : rg.c2,
    };
  }
  return rg;
}

function buildPayload(cut: boolean): ClipPayload | null {
  const st = S();
  const sheet = st.wb.activeSheet;
  if (st.sel.ranges.length > 1) {
    alertBox("This action won't work on multiple selections.");
    return null;
  }
  const rg = clipRange(sheet, primaryRange(st.sel));
  const cells: (Cell | undefined)[][] = [];
  const styleIds: number[][] = [];
  const lines: string[] = [];
  for (let r = rg.r1; r <= rg.r2; r++) {
    const row: (Cell | undefined)[] = [];
    const srow: number[] = [];
    const texts: string[] = [];
    for (let c = rg.c1; c <= rg.c2; c++) {
      const cell = sheet.getCell(r, c);
      row.push(cell ? { ...cell } : undefined);
      const sid = sheet.styleIdAt(r, c);
      srow.push(sid);
      if (sheet.isRowHidden(r) || sheet.isColHidden(c)) continue;
      const d = displayOf(sheet, r, c, st.wb.styles.get(sid), 30);
      let t = d.left !== undefined ? (d.left + d.right!).trim() : d.text;
      if (/[\t\n"]/.test(t)) t = '"' + t.replace(/"/g, '""') + '"';
      texts.push(t);
    }
    cells.push(row);
    styleIds.push(srow);
    if (!sheet.isRowHidden(r)) lines.push(texts.join('\t'));
  }
  const merges = sheet.merges
    .filter((m) => m.r1 >= rg.r1 && m.r2 <= rg.r2 && m.c1 >= rg.c1 && m.c2 <= rg.c2)
    .map((m) => ({ r1: m.r1 - rg.r1, r2: m.r2 - rg.r1, c1: m.c1 - rg.c1, c2: m.c2 - rg.c1 }));
  const colWidths: number[] = [];
  for (let c = rg.c1; c <= rg.c2; c++) colWidths.push(sheet.colWidth(c));
  const rowHeights: number[] = [];
  for (let r = rg.r1; r <= rg.r2; r++) rowHeights.push(sheet.rowHeight(r));
  return { sheetId: sheet.id, sheetName: sheet.name, range: rg, cells, styleIds, merges, colWidths, rowHeights, text: lines.join('\r\n') + '\r\n', cut };
}

// ---------- HTML generation (for pasting into real Excel / other apps) ----------

function edgeCss(e?: BorderEdge): string | undefined {
  if (!e) return undefined;
  const w = e.style === 'medium' || e.style.startsWith('medium') ? '1.0pt' : e.style === 'thick' ? '1.5pt' : '.5pt';
  const kind = e.style === 'double' ? 'double' : e.style.includes('dash') ? 'dashed' : e.style === 'dotted' || e.style === 'hair' ? 'dotted' : 'solid';
  return `${e.style === 'double' ? '2.0pt' : w} ${kind} ${e.color ?? 'windowtext'}`;
}

function styleCss(s: CellStyle): string {
  const p: string[] = [];
  if (s.fontName) p.push(`font-family:"${s.fontName}"`);
  if (s.fontSize) p.push(`font-size:${s.fontSize}pt`);
  if (s.bold) p.push('font-weight:700');
  if (s.italic) p.push('font-style:italic');
  const deco = [s.underline ? 'underline' : '', s.strike ? 'line-through' : ''].filter(Boolean).join(' ');
  if (deco) p.push(`text-decoration:${deco}`);
  if (s.fontColor) p.push(`color:${s.fontColor}`);
  if (s.fillColor) p.push(`background:${s.fillColor}`);
  if (s.hAlign && s.hAlign !== 'general') p.push(`text-align:${s.hAlign === 'centerContinuous' ? 'center' : s.hAlign}`);
  if (s.vAlign) p.push(`vertical-align:${s.vAlign}`);
  if (s.wrap) p.push('white-space:normal');
  if (s.numFmt) p.push(`mso-number-format:"${s.numFmt.replace(/"/g, '\\"')}"`);
  const bt = edgeCss(s.bTop);
  const bb = edgeCss(s.bBottom);
  const bl = edgeCss(s.bLeft);
  const br = edgeCss(s.bRight);
  if (bt) p.push(`border-top:${bt}`);
  if (bb) p.push(`border-bottom:${bb}`);
  if (bl) p.push(`border-left:${bl}`);
  if (br) p.push(`border-right:${br}`);
  return p.join(';');
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
}

function buildHtml(p: ClipPayload): string {
  const st = S();
  const sheet = st.wb.sheetById(p.sheetId)!;
  const rows: string[] = [];
  const covered = new Set<string>();
  for (let i = 0; i < p.cells.length; i++) {
    const r = p.range.r1 + i;
    if (sheet.isRowHidden(r)) continue;
    const tds: string[] = [];
    for (let j = 0; j < p.cells[i].length; j++) {
      const c = p.range.c1 + j;
      if (sheet.isColHidden(c) || covered.has(`${i},${j}`)) continue;
      const m = p.merges.find((mm) => mm.r1 === i && mm.c1 === j);
      let span = '';
      if (m) {
        span = ` rowspan=${m.r2 - m.r1 + 1} colspan=${m.c2 - m.c1 + 1}`;
        for (let a = m.r1; a <= m.r2; a++) for (let b = m.c1; b <= m.c2; b++) if (a !== i || b !== j) covered.add(`${a},${b}`);
      }
      const style = st.wb.styles.get(p.styleIds[i][j]);
      const d = displayOf(sheet, r, c, style, 30);
      const v = getComputed(sheet, r, c);
      const num = typeof v === 'number' ? ` x:num="${v}"` : '';
      const fm = p.cells[i][j]?.f ? ` x:fmla="${p.cells[i][j]!.f!.replace(/"/g, '&quot;')}"` : '';
      const text = d.left !== undefined ? (d.left + d.right!).trim() : d.text;
      const w = ` width=${p.colWidths[j]}`;
      tds.push(`<td${span}${w}${num}${fm} style='${styleCss(style)}'>${escapeHtml(text)}</td>`);
    }
    rows.push(`<tr height=${p.rowHeights[i]}>${tds.join('')}</tr>`);
  }
  return (
    `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">` +
    `<head><meta charset="utf-8"><meta name=Generator content="MyExcel"></head><body>` +
    `<table border=0 cellpadding=0 cellspacing=0 style='border-collapse:collapse;font-family:Calibri;font-size:11pt'>${rows.join('')}</table></body></html>`
  );
}

// ---------- copy / cut ----------

/** Handles the native copy/cut event (Ctrl+C / Ctrl+X). */
export function onCopyEvent(e: ClipboardEvent, cut: boolean): void {
  const p = doCopy(cut);
  if (!p || !e.clipboardData) return;
  e.clipboardData.setData('text/plain', p.text);
  e.clipboardData.setData('text/html', buildHtml(p));
  e.preventDefault();
}

export function doCopy(cut: boolean): ClipPayload | null {
  const st = S();
  const sheet = st.wb.activeSheet;
  if (cut && !guardRanges(sheet, st.sel.ranges)) return null;
  const p = buildPayload(cut);
  if (!p) return null;
  internal = p;
  setState({ clip: { sheetId: p.sheetId, range: p.range, cut } });
  setStatus('Select destination and press ENTER or choose Paste');
  return p;
}

/** Ribbon Copy button: uses the async clipboard API where available. */
export async function copyToSystem(cut: boolean): Promise<void> {
  const p = doCopy(cut);
  if (!p) return;
  try {
    if (typeof ClipboardItem !== 'undefined' && navigator.clipboard?.write) {
      await navigator.clipboard.write([
        new ClipboardItem({
          'text/plain': new Blob([p.text], { type: 'text/plain' }),
          'text/html': new Blob([buildHtml(p)], { type: 'text/html' }),
        }),
      ]);
    } else await navigator.clipboard?.writeText(p.text);
  } catch {
    /* permissions – internal clipboard still works */
  }
}

// ---------- paste ----------

export type PasteWhat =
  | 'all'
  | 'formulas'
  | 'values'
  | 'formats'
  | 'comments'
  | 'validation'
  | 'allExceptBorders'
  | 'columnWidths'
  | 'formulasAndNumberFormats'
  | 'valuesAndNumberFormats'
  | 'allMergingConditionalFormats';

export interface PasteOptions {
  what: PasteWhat;
  operation: 'none' | 'add' | 'subtract' | 'multiply' | 'divide';
  skipBlanks: boolean;
  transpose: boolean;
  link: boolean;
}

export const DEFAULT_PASTE: PasteOptions = { what: 'all', operation: 'none', skipBlanks: false, transpose: false, link: false };

/** Native paste event (Ctrl+V). */
export function onPasteEvent(e: ClipboardEvent): void {
  const html = e.clipboardData?.getData('text/html') ?? '';
  const text = e.clipboardData?.getData('text/plain') ?? '';
  e.preventDefault();
  pasteFrom({ html, text }, DEFAULT_PASTE);
}

/** Ribbon Paste button. */
export async function pasteFromSystem(opts: PasteOptions = DEFAULT_PASTE): Promise<void> {
  let html = '';
  let text = '';
  try {
    if (navigator.clipboard?.read) {
      const items = await navigator.clipboard.read();
      for (const it of items) {
        if (it.types.includes('text/html')) html = await (await it.getType('text/html')).text();
        if (it.types.includes('text/plain')) text = await (await it.getType('text/plain')).text();
      }
    } else if (navigator.clipboard?.readText) text = await navigator.clipboard.readText();
  } catch {
    /* permission denied – fall back to internal clipboard */
    if (internal) text = internal.text;
  }
  pasteFrom({ html, text }, opts);
}

function normText(s: string): string {
  return s.replace(/\r/g, '').replace(/\n+$/, '');
}

export function pasteFrom(data: { html: string; text: string }, opts: PasteOptions): void {
  const useInternal = internal && (normText(data.text) === normText(internal.text) || (!data.text && !data.html));
  if (useInternal && internal) {
    pasteInternal(internal, opts);
    return;
  }
  const grid = data.html ? parseHtmlTable(data.html) : null;
  if (grid && grid.cells.length) pasteExternal(grid, opts);
  else if (data.text) pasteExternal(parseTsv(data.text), opts);
}

function targetRange(srcH: number, srcW: number): Range | null {
  const st = S();
  const sel = primaryRange(st.sel);
  const selH = sel.r2 - sel.r1 + 1;
  const selW = sel.c2 - sel.c1 + 1;
  // tile if selection is a multiple of the source size
  if (selH % srcH === 0 && selW % srcW === 0 && (selH > srcH || selW > srcW) && !isFullCols(sel) && !isFullRows(sel)) return sel;
  const r2 = sel.r1 + srcH - 1;
  const c2 = sel.c1 + srcW - 1;
  if (r2 >= MAX_ROWS || c2 >= MAX_COLS) {
    alertBox("MyExcel can't paste the data because the copy area and paste area aren't the same size.");
    return null;
  }
  return { r1: sel.r1, c1: sel.c1, r2, c2 };
}

function applyOp(op: PasteOptions['operation'], target: unknown, src: unknown): unknown {
  if (op === 'none') return src;
  const a = typeof target === 'number' ? target : 0;
  if (typeof src !== 'number') return src;
  switch (op) {
    case 'add': return a + src;
    case 'subtract': return a - src;
    case 'multiply': return a * src;
    case 'divide': return src === 0 ? '#DIV/0!' : a / src;
  }
  return src;
}

function pasteInternal(p: ClipPayload, opts: PasteOptions): void {
  const st = S();
  const wb = st.wb;
  const sheet = wb.activeSheet;
  const src = wb.sheetById(p.sheetId);
  const H = opts.transpose ? p.cells[0]?.length ?? 0 : p.cells.length;
  const W = opts.transpose ? p.cells.length : p.cells[0]?.length ?? 0;
  const tgt = targetRange(H, W);
  if (!tgt) return;
  if (!guardRanges(sheet, [tgt])) return;
  const overlapsMerge = sheet.merges.some((m) => m.r1 <= tgt.r2 && tgt.r1 <= m.r2 && m.c1 <= tgt.c2 && tgt.c1 <= m.c2 && !(m.r1 >= tgt.r1 && m.r2 <= tgt.r2 && m.c1 >= tgt.c1 && m.c2 <= tgt.c2));
  if (overlapsMerge) {
    alertBox("We can't do that to a merged cell.");
    return;
  }
  const what = opts.what;
  const withValues = ['all', 'formulas', 'values', 'allExceptBorders', 'formulasAndNumberFormats', 'valuesAndNumberFormats', 'allMergingConditionalFormats'].includes(what);
  const withFormats = ['all', 'formats', 'allExceptBorders', 'allMergingConditionalFormats'].includes(what);
  const withNumFmt = what === 'formulasAndNumberFormats' || what === 'valuesAndNumberFormats';
  const valuesOnly = what === 'values' || what === 'valuesAndNumberFormats';
  transact(p.cut ? 'Cut' : 'Paste', (tx) => {
    const styles = wb.styles;
    for (let r = tgt.r1; r <= tgt.r2; r++) {
      for (let c: number = tgt.c1; c <= tgt.c2; c++) {
        const i0: number = (r - tgt.r1) % H;
        const j0: number = (c - tgt.c1) % W;
        const i = opts.transpose ? j0 : i0;
        const j = opts.transpose ? i0 : j0;
        const scell = p.cells[i]?.[j];
        const sr = p.range.r1 + i;
        const sc = p.range.c1 + j;
        const sid = p.styleIds[i][j];
        if (opts.skipBlanks && (!scell || (scell.v === undefined && scell.f === undefined))) continue;
        if (what === 'columnWidths') continue;
        const cur = sheet.getCell(r, c);
        const next: Cell = { ...(cur ?? {}) };
        if (opts.link) {
          next.f = `=${src && src.id !== sheet.id ? quoteSheetName(p.sheetName) + '!' : ''}${colToName(sc)}${sr + 1}`;
          delete next.v;
          delete next.e;
          tx.setCell(sheet, r, c, next);
          continue;
        }
        if (withValues) {
          delete next.v;
          delete next.f;
          delete next.e;
          if (scell) {
            if (scell.f && !valuesOnly && opts.operation === 'none') {
              next.f = p.cut ? scell.f : shiftFormula(scell.f, r - sr, c - sc);
            } else {
              let v: unknown = scell.f && src ? getComputed(src, sr, sc) : scell.v;
              if (isErrorVal(v)) {
                next.v = v.error;
                next.e = true;
              } else {
                v = applyOp(opts.operation, cur?.v, v);
                if (v !== undefined && v !== null) next.v = v as Cell['v'];
                if (typeof v === 'string' && v.startsWith('#') && opts.operation !== 'none') next.e = true;
                if (scell.e && opts.operation === 'none') next.e = true;
              }
            }
          }
        }
        if (withFormats) {
          if (what === 'allExceptBorders') {
            const s = styles.get(sid);
            next.s = styles.merge(sheet.styleIdAt(r, c), { ...s, bTop: styles.get(sheet.styleIdAt(r, c)).bTop ?? null, bBottom: styles.get(sheet.styleIdAt(r, c)).bBottom ?? null, bLeft: styles.get(sheet.styleIdAt(r, c)).bLeft ?? null, bRight: styles.get(sheet.styleIdAt(r, c)).bRight ?? null });
          } else next.s = sid;
        } else if (withNumFmt) {
          const nf = styles.get(sid).numFmt;
          next.s = styles.merge(sheet.styleIdAt(r, c), { numFmt: nf ?? null });
        }
        if (what === 'all' || what === 'comments' || what === 'allExceptBorders') {
          if (scell?.note) next.note = scell.note;
          else if (what !== 'comments') delete next.note;
          if (scell?.link) next.link = scell.link;
        }
        tx.setCell(sheet, r, c, next);
      }
    }
    if (what === 'columnWidths') {
      const m = new Map(sheet.colWidths);
      for (let c = tgt.c1; c <= tgt.c2; c++) m.set(c, p.colWidths[(c - tgt.c1) % p.colWidths.length]);
      tx.setMeta(sheet, 'colWidths', m);
    }
    if ((withFormats || what === 'all') && p.merges.length && !opts.transpose) {
      const add: Range[] = [];
      for (let r0 = tgt.r1; r0 <= tgt.r2; r0 += H)
        for (let c0 = tgt.c1; c0 <= tgt.c2; c0 += W)
          for (const m of p.merges) add.push({ r1: r0 + m.r1, r2: r0 + m.r2, c1: c0 + m.c1, c2: c0 + m.c2 });
      tx.setMeta(sheet, 'merges', sheet.merges.filter((m) => !(m.r1 <= tgt.r2 && tgt.r1 <= m.r2 && m.c1 <= tgt.c2 && tgt.c1 <= m.c2)).concat(add));
    }
    if (what === 'validation' || what === 'all') {
      if (src) {
        const dvs = src.validations.filter((dv) => dv.ranges.some((rg) => rg.r1 <= p.range.r2 && p.range.r1 <= rg.r2 && rg.c1 <= p.range.c2 && p.range.c1 <= rg.c2));
        if (dvs.length && src.id === sheet.id && !p.cut) {
          const dr = tgt.r1 - p.range.r1;
          const dc = tgt.c1 - p.range.c1;
          const added = dvs.map((dv) => ({
            ...structuredClone(dv),
            id: dv.id + 'p' + Date.now().toString(36),
            ranges: dv.ranges
              .map((rg) => ({ r1: Math.max(rg.r1, p.range.r1) + dr, r2: Math.min(rg.r2, p.range.r2) + dr, c1: Math.max(rg.c1, p.range.c1) + dc, c2: Math.min(rg.c2, p.range.c2) + dc }))
              .filter((rg) => rg.r1 <= rg.r2 && rg.c1 <= rg.c2),
          }));
          tx.setMeta(sheet, 'validations', sheet.validations.concat(added));
        }
      }
    }
    if (p.cut && src) {
      // clear source cells not overlapped by the target
      const dr = tgt.r1 - p.range.r1;
      const dc = tgt.c1 - p.range.c1;
      for (let r = p.range.r1; r <= p.range.r2; r++)
        for (let c = p.range.c1; c <= p.range.c2; c++) {
          if (src.id === sheet.id && r >= tgt.r1 && r <= tgt.r2 && c >= tgt.c1 && c <= tgt.c2) continue;
          if (src.getCell(r, c)) tx.setCell(src, r, c, undefined);
        }
      if (src.id === sheet.id && p.merges.length) {
        tx.setMeta(sheet, 'merges', sheet.merges.filter((m) => !(m.r1 >= p.range.r1 && m.r2 <= p.range.r2 && m.c1 >= p.range.c1 && m.c2 <= p.range.c2 && !(m.r1 >= tgt.r1 && m.r2 <= tgt.r2 && m.c1 >= tgt.c1 && m.c2 <= tgt.c2))));
      }
      retargetMovedRefs(tx, src, p.range, sheet, dr, dc);
    }
  }, rangeSel(tgt));
  if (p.cut) {
    internal = null;
    setState({ clip: null });
  }
  setStatus(null);
}

/** After cut-paste, formulas referencing the moved block follow it (Excel semantics). */
function retargetMovedRefs(tx: import('../../model/commands').Tx, src: Sheet, rg: Range, dst: Sheet, dr: number, dc: number): void {
  const wb = S().wb;
  for (const s of wb.sheets) {
    for (const [r, row] of [...s.rows]) {
      for (const [c, cell] of [...row]) {
        if (!cell.f) continue;
        const toks = tokenize(cell.f);
        let changed = false;
        const out = toks
          .map((t) => {
            if (t.type !== 'ref' || !t.ref) return t.text;
            const refSheet = t.ref.sheet ? wb.sheetByName(t.ref.sheet) : s;
            if (refSheet !== src) return t.text;
            const rr = refToRange(t.ref);
            if (rr.r1 < rg.r1 || rr.r2 > rg.r2 || rr.c1 < rg.c1 || rr.c2 > rg.c2) return t.text;
            const mv = (p: RefInfo['a']) => ({ ...p, r: p.r !== undefined ? p.r + dr : undefined, c: p.c !== undefined ? p.c + dc : undefined });
            changed = true;
            const sheetName = dst !== s || t.ref.sheet ? dst.name : undefined;
            return refToString({ sheet: sheetName, a: mv(t.ref.a), b: t.ref.b ? mv(t.ref.b) : undefined });
          })
          .join('');
        if (changed) tx.setCell(s, r, c, { ...cell, f: out });
      }
    }
  }
}

// ---------- external data ----------

interface ExtGrid {
  cells: { text: string; num?: number; formula?: string; style?: StylePatch; rowspan?: number; colspan?: number }[][];
  colWidths?: number[];
}

function parseTsv(text: string): ExtGrid {
  const rows: string[][] = [];
  let row: string[] = [];
  let cur = '';
  let q = false;
  const t = text.replace(/\r\n/g, '\n').replace(/\n$/, '');
  for (let i = 0; i < t.length; i++) {
    const ch = t[i];
    if (q) {
      if (ch === '"') {
        if (t[i + 1] === '"') {
          cur += '"';
          i++;
        } else q = false;
      } else cur += ch;
    } else if (ch === '"' && cur === '') q = true;
    else if (ch === '\t') {
      row.push(cur);
      cur = '';
    } else if (ch === '\n') {
      row.push(cur);
      rows.push(row);
      row = [];
      cur = '';
    } else cur += ch;
  }
  row.push(cur);
  rows.push(row);
  return { cells: rows.map((r) => r.map((text) => ({ text }))) };
}

function parseCssBlock(css: string): Map<string, Record<string, string>> {
  const m = new Map<string, Record<string, string>>();
  const re = /([^{}]+)\{([^}]*)\}/g;
  let x: RegExpExecArray | null;
  while ((x = re.exec(css))) {
    const props = parseDecl(x[2]);
    for (const sel of x[1].split(',')) {
      const s = sel.trim();
      const cls = /\.([\w-]+)/.exec(s);
      if (cls) m.set(cls[1], { ...(m.get(cls[1]) ?? {}), ...props });
      else if (/^td$/i.test(s)) m.set('__td', { ...(m.get('__td') ?? {}), ...props });
    }
  }
  return m;
}

function parseDecl(s: string): Record<string, string> {
  const o: Record<string, string> = {};
  for (const part of s.split(';')) {
    const i = part.indexOf(':');
    if (i < 0) continue;
    o[part.slice(0, i).trim().toLowerCase()] = part.slice(i + 1).trim();
  }
  return o;
}

function cssColor(v: string | undefined): string | undefined {
  if (!v) return undefined;
  const t = v.trim().toLowerCase();
  if (t === 'windowtext' || t === 'black') return '#000000';
  if (t === 'white') return '#FFFFFF';
  if (/^#[0-9a-f]{6}$/.test(t)) return t.toUpperCase();
  if (/^#[0-9a-f]{3}$/.test(t)) return ('#' + t[1] + t[1] + t[2] + t[2] + t[3] + t[3]).toUpperCase();
  const m = /rgb\((\d+),\s*(\d+),\s*(\d+)/.exec(t);
  if (m) return '#' + [m[1], m[2], m[3]].map((n) => (+n).toString(16).padStart(2, '0')).join('').toUpperCase();
  if (t === 'red') return '#FF0000';
  if (t === 'blue') return '#0000FF';
  if (t === 'yellow') return '#FFFF00';
  if (t === 'green') return '#008000';
  return undefined;
}

function cssBorder(v: string | undefined): BorderEdge | undefined {
  if (!v || /none/i.test(v)) return undefined;
  const w = parseFloat(v) || 0.5;
  const kind = /double/i.test(v) ? 'double' : /dashed/i.test(v) ? 'dashed' : /dotted/i.test(v) ? 'dotted' : 'solid';
  let style: BorderStyleName = 'thin';
  if (kind === 'double') style = 'double';
  else if (kind === 'dashed') style = w >= 1 ? 'mediumDashed' : 'dashed';
  else if (kind === 'dotted') style = 'dotted';
  else style = w >= 1.5 ? 'thick' : w >= 1 ? 'medium' : 'thin';
  const colorMatch = /(#[0-9a-f]{3,6}|rgb\([^)]*\)|windowtext|black|[a-z]+)\s*$/i.exec(v.trim());
  return { style, color: cssColor(colorMatch?.[1]) ?? '#000000' };
}

function propsToStyle(p: Record<string, string>): StylePatch {
  const s: StylePatch = {};
  if (p['font-weight'] && (p['font-weight'] === 'bold' || parseInt(p['font-weight'], 10) >= 600)) s.bold = true;
  if (p['font-style'] === 'italic') s.italic = true;
  const deco = p['text-decoration'] ?? p['text-underline-style'] ?? '';
  if (/underline|single/.test(deco)) s.underline = 'single';
  if (/line-through/.test(deco)) s.strike = true;
  const col = cssColor(p['color']);
  if (col && col !== '#000000') s.fontColor = col;
  const bg = cssColor(p['background'] ?? p['background-color']);
  if (bg && bg !== '#FFFFFF') s.fillColor = bg;
  if (p['font-size']) {
    const n = parseFloat(p['font-size']);
    if (n) s.fontSize = /px/.test(p['font-size']) ? Math.round((n * 72) / 96) : n;
  }
  if (p['font-family']) s.fontName = p['font-family'].split(',')[0].replace(/["']/g, '').trim();
  const ta = p['text-align'];
  if (ta === 'center' || ta === 'right' || ta === 'left') s.hAlign = ta;
  const va = p['vertical-align'];
  if (va === 'top') s.vAlign = 'top';
  else if (va === 'middle') s.vAlign = 'middle';
  if (p['white-space'] === 'normal') s.wrap = true;
  const nf = p['mso-number-format'];
  if (nf) {
    let f = nf.replace(/^["']|["']$/g, '').replace(/\\(.)/g, '$1');
    if (f === 'Standard' || f === 'General') f = '';
    if (f === 'Percent') f = '0.00%';
    if (f === 'Fixed') f = '0.00';
    if (f === 'Short Date') f = 'm/d/yyyy';
    if (f) s.numFmt = f;
  }
  const all = cssBorder(p['border']);
  const bt = cssBorder(p['border-top']) ?? all;
  const bb = cssBorder(p['border-bottom']) ?? all;
  const bl = cssBorder(p['border-left']) ?? all;
  const br = cssBorder(p['border-right']) ?? all;
  if (bt) s.bTop = bt;
  if (bb) s.bBottom = bb;
  if (bl) s.bLeft = bl;
  if (br) s.bRight = br;
  return s;
}

function parseHtmlTable(html: string): ExtGrid | null {
  if (typeof DOMParser === 'undefined') return null;
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const table = doc.querySelector('table');
  if (!table) return null;
  let css = new Map<string, Record<string, string>>();
  doc.querySelectorAll('style').forEach((st) => {
    const parsed = parseCssBlock((st.textContent ?? '').replace(/<!--|-->/g, ''));
    css = new Map([...css, ...parsed]);
  });
  const grid: ExtGrid['cells'] = [];
  const occupied: boolean[][] = [];
  const trs = table.querySelectorAll('tr');
  trs.forEach((tr, ri) => {
    grid[ri] = grid[ri] ?? [];
    occupied[ri] = occupied[ri] ?? [];
    let ci = 0;
    tr.querySelectorAll('td,th').forEach((td) => {
      while (occupied[ri][ci]) ci++;
      const props: Record<string, string> = { ...(css.get('__td') ?? {}) };
      td.classList.forEach((cl) => Object.assign(props, css.get(cl) ?? {}));
      Object.assign(props, parseDecl(td.getAttribute('style') ?? ''));
      if (td.tagName === 'TH') props['font-weight'] = props['font-weight'] ?? 'bold';
      const inner: Record<string, string> = td.querySelector('b,strong') ? { 'font-weight': 'bold' } : {};
      const style = propsToStyle({ ...props, ...inner });
      const rs = parseInt(td.getAttribute('rowspan') ?? '1', 10) || 1;
      const cs = parseInt(td.getAttribute('colspan') ?? '1', 10) || 1;
      const text = (td as HTMLElement).innerText ?? td.textContent ?? '';
      const numAttr = td.getAttribute('x:num');
      const fm = td.getAttribute('x:fmla') ?? undefined;
      const num = numAttr !== null && numAttr !== '' ? parseFloat(numAttr) : undefined;
      grid[ri][ci] = { text: text.replace(/ /g, ' ').replace(/\n$/, ''), num: num !== undefined && !isNaN(num) ? num : undefined, formula: fm, style, rowspan: rs, colspan: cs };
      for (let a = 0; a < rs; a++)
        for (let b = 0; b < cs; b++) {
          occupied[ri + a] = occupied[ri + a] ?? [];
          occupied[ri + a][ci + b] = true;
          if (a || b) {
            grid[ri + a] = grid[ri + a] ?? [];
            grid[ri + a][ci + b] = grid[ri + a][ci + b] ?? { text: '' };
          }
        }
      ci += cs;
    });
  });
  const w = Math.max(0, ...grid.map((r) => r.length));
  for (const r of grid) for (let j = 0; j < w; j++) r[j] = r[j] ?? { text: '' };
  return { cells: grid };
}

function pasteExternal(g: ExtGrid, opts: PasteOptions): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  const H0 = g.cells.length;
  const W0 = Math.max(0, ...g.cells.map((r) => r.length));
  if (!H0 || !W0) return;
  const H = opts.transpose ? W0 : H0;
  const W = opts.transpose ? H0 : W0;
  const tgt = targetRange(H, W);
  if (!tgt) return;
  if (!guardRanges(sheet, [tgt])) return;
  const valuesOnly = opts.what === 'values' || opts.what === 'valuesAndNumberFormats';
  const withFormats = opts.what === 'all' || opts.what === 'formats' || opts.what === 'allExceptBorders';
  transact('Paste', (tx) => {
    const merges: Range[] = [];
    for (let r = tgt.r1; r <= tgt.r2; r++) {
      for (let c: number = tgt.c1; c <= tgt.c2; c++) {
        const i0: number = (r - tgt.r1) % H;
        const j0: number = (c - tgt.c1) % W;
        const i = opts.transpose ? j0 : i0;
        const j = opts.transpose ? i0 : j0;
        const x = g.cells[i]?.[j];
        if (!x) continue;
        if (opts.skipBlanks && x.text === '') continue;
        if (opts.what !== 'formats') {
          if (x.formula && !valuesOnly) writeInput(tx, sheet, r, c, x.formula);
          else if (x.num !== undefined) {
            const cur = sheet.getCell(r, c) ?? {};
            const v = applyOp(opts.operation, cur.v, x.num);
            const next: Cell = { ...cur, v: v as number };
            delete next.f;
            delete next.e;
            tx.setCell(sheet, r, c, next);
          } else if (opts.operation !== 'none') {
            const cur = sheet.getCell(r, c) ?? {};
            const n = Number(x.text);
            if (!isNaN(n) && x.text !== '') tx.setCell(sheet, r, c, { ...cur, v: applyOp(opts.operation, cur.v, n) as number });
          } else writeInput(tx, sheet, r, c, x.text);
        }
        if (x.style && (withFormats || opts.what === 'valuesAndNumberFormats') && Object.keys(x.style).length) {
          const cur = sheet.getCell(r, c) ?? {};
          const patch = opts.what === 'valuesAndNumberFormats' ? { numFmt: x.style.numFmt ?? null } : x.style;
          tx.setCell(sheet, r, c, { ...cur, s: st.wb.styles.merge(cur.s ?? sheet.styleIdAt(r, c), patch) });
        }
        if ((x.rowspan ?? 1) > 1 || (x.colspan ?? 1) > 1) {
          if (!opts.transpose) merges.push({ r1: r, c1: c, r2: r + (x.rowspan ?? 1) - 1, c2: c + (x.colspan ?? 1) - 1 });
        }
      }
    }
    if (merges.length && withFormats) tx.setMeta(sheet, 'merges', sheet.merges.filter((m) => !(m.r1 <= tgt.r2 && tgt.r1 <= m.r2 && m.c1 <= tgt.c2 && tgt.c1 <= m.c2)).concat(merges));
  }, rangeSel(tgt));
  setState({ clip: null });
  setStatus(null);
  bump();
}

export function clearClip(): void {
  if (S().clip) {
    setState({ clip: null });
    setStatus(null);
  }
}

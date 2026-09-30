// Canvas renderer for the grid.
import { isErrorVal } from '../../engine/Engine';
import { colToName, isFullCols, isFullRows, Range } from '../../model/address';
import { formatGeneral } from '../../model/numfmt';
import type { Selection } from '../../model/selection';
import type { Sheet } from '../../model/sheet';
import { BorderEdge, CellStyle, DEFAULT_FONT_SIZE } from '../../model/styles';
import type { Workbook } from '../../model/workbook';
import { cfAt, CFResult } from '../../state/cf';
import { cssFont, ptToPx } from '../../state/measure';
import { sparklineValues } from '../../state/sparkline';
import { tableAt, tableLook } from '../../state/tableStyles';
import { displayOf, Display } from '../../state/values';
import { adaptColor, GridPalette } from '../theme';
import { colX, Pane, paneRange, rowY, SPLIT_BAR, Viewport } from './geometry';

export interface RefHighlight {
  range: Range;
  color: string;
}

export interface RenderState {
  sheet: Sheet;
  wb: Workbook;
  vp: Viewport;
  pal: GridPalette;
  dark: boolean;
  dpr: number;
  sel: Selection;
  editing: { r: number; c: number } | null;
  refs: RefHighlight[];
  clip: { range: Range; phase: number } | null;
  fillPreview: Range | null;
  showFormulas: boolean;
  pageBreaks: { rows: number[]; cols: number[] } | null;
  invalid: Set<string> | null;
  filterButtons: boolean;
  listArrow: { r: number; c: number } | null;
  painterRange: Range | null;
  /** Print rendering: no selection/UI overlays, gridlines per page setup, drawn at an origin offset. */
  print?: { gridlines: boolean; origin: { x: number; y: number } };
}

let dpr = 1;
const snap = (v: number) => Math.round(v * dpr) / dpr;

let currentFont = '';
function setFont(ctx: CanvasRenderingContext2D, f: string): void {
  if (f !== currentFont) {
    ctx.font = f;
    currentFont = f;
  }
}

interface CellInfo {
  style: CellStyle;
  fill?: string;
  cf: CFResult | null;
  tableBorders?: { bTop?: BorderEdge; bBottom?: BorderEdge; bLeft?: BorderEdge; bRight?: BorderEdge };
}

function resolveCell(rs: RenderState, r: number, c: number): CellInfo {
  const { sheet, wb } = rs;
  let style = wb.styles.get(sheet.styleIdAt(r, c));
  let fill = style.fillColor;
  let tableBorders: CellInfo['tableBorders'];
  if (sheet.tables.length) {
    const t = tableAt(sheet, r, c);
    if (t) {
      const look = tableLook(t, r, c);
      // explicit formatting wins over table style
      if (!fill && look.fill) fill = look.fill;
      if (look.fontColor && !style.fontColor) style = { ...style, fontColor: look.fontColor };
      if (look.bold && style.bold === undefined) style = { ...style, bold: true };
      tableBorders = { bTop: look.bTop, bBottom: look.bBottom, bLeft: look.bLeft, bRight: look.bRight };
    }
  }
  const cf = sheet.conditionalFormats.length ? cfAt(sheet, r, c) : null;
  if (cf) {
    if (cf.fill) fill = cf.fill;
    if (cf.style) {
      const s = cf.style;
      if (s.fillColor) fill = s.fillColor;
      style = {
        ...style,
        ...(s.fontColor ? { fontColor: s.fontColor } : {}),
        ...(s.bold ? { bold: true } : {}),
        ...(s.italic ? { italic: true } : {}),
        ...(s.underline ? { underline: 'single' as const } : {}),
        ...(s.strike ? { strike: true } : {}),
        ...(s.numFmt ? { numFmt: s.numFmt } : {}),
      };
      if (s.borderColor) {
        const e: BorderEdge = { style: 'thin', color: s.borderColor };
        style = { ...style, bTop: e, bBottom: e, bLeft: e, bRight: e };
      }
    }
  }
  return { style, fill, cf, tableBorders };
}

function hasValue(sheet: Sheet, r: number, c: number): boolean {
  const cell = sheet.getCell(r, c);
  return !!cell && (cell.f !== undefined || (cell.v !== undefined && cell.v !== null && cell.v !== ''));
}

// ---------- borders ----------

function borderWidth(e: BorderEdge): number {
  switch (e.style) {
    case 'medium':
    case 'mediumDashed':
    case 'mediumDashDot':
    case 'mediumDashDotDot':
    case 'slantDashDot':
      return 2;
    case 'thick':
      return 3;
    case 'double':
      return 3;
    default:
      return 1;
  }
}

function dashFor(e: BorderEdge): number[] {
  switch (e.style) {
    case 'dashed':
    case 'mediumDashed':
      return [4, 2];
    case 'dotted':
      return [1, 1];
    case 'hair':
      return [1, 1];
    case 'dashDot':
    case 'mediumDashDot':
    case 'slantDashDot':
      return [6, 2, 2, 2];
    case 'dashDotDot':
    case 'mediumDashDotDot':
      return [6, 2, 2, 2, 2, 2];
    default:
      return [];
  }
}

/** Draw a horizontal (dir='h') or vertical border segment centred on the gridline at `pos`. */
function drawEdge(ctx: CanvasRenderingContext2D, e: BorderEdge, dir: 'h' | 'v', pos: number, a: number, b: number, z: number, dark: boolean): void {
  const color = adaptColor(e.color ?? '#000000', 'text', dark)!;
  ctx.fillStyle = color;
  ctx.strokeStyle = color;
  const w = Math.max(1, Math.round(borderWidth(e) * Math.max(1, z * 0.9)));
  if (e.style === 'double') {
    const g = 1 / dpr >= 1 ? 1 : 1;
    if (dir === 'h') {
      ctx.fillRect(snap(a), snap(pos - 2), b - a, g);
      ctx.fillRect(snap(a), snap(pos), b - a, g);
    } else {
      ctx.fillRect(snap(pos - 2), snap(a), g, b - a);
      ctx.fillRect(snap(pos), snap(a), g, b - a);
    }
    return;
  }
  const dash = dashFor(e);
  const off = Math.floor((w - 1) / 2);
  if (!dash.length) {
    if (dir === 'h') ctx.fillRect(snap(a), snap(pos - off), b - a, w);
    else ctx.fillRect(snap(pos - off), snap(a), w, b - a);
    return;
  }
  ctx.save();
  ctx.lineWidth = w;
  ctx.setLineDash(dash.map((d) => d * Math.max(1, w * 0.8)));
  ctx.beginPath();
  if (dir === 'h') {
    const y = snap(pos - off) + w / 2;
    ctx.moveTo(a, y);
    ctx.lineTo(b, y);
  } else {
    const x = snap(pos - off) + w / 2;
    ctx.moveTo(x, a);
    ctx.lineTo(x, b);
  }
  ctx.stroke();
  ctx.restore();
}

// ---------- icons ----------

const IC = { green: '#00A650', yellow: '#FFC000', red: '#E53935', gray: '#808080', black: '#262626', blue: '#4472C4' };

export function drawIcon(ctx: CanvasRenderingContext2D, set: string, idx: number, count: number, x: number, y: number, s: number): void {
  const lvl = count <= 1 ? 1 : idx / (count - 1); // 0 = worst, 1 = best
  const col3 = lvl >= 0.99 ? IC.green : lvl <= 0.01 ? IC.red : IC.yellow;
  const cx = x + s / 2;
  const cy = y + s / 2;
  const rr = s * 0.42;
  ctx.save();
  if (/Arrows/.test(set)) {
    const gray = /Gray/.test(set);
    ctx.fillStyle = gray ? IC.gray : lvl >= 0.99 ? IC.green : lvl <= 0.01 ? IC.red : IC.yellow;
    ctx.translate(cx, cy);
    const ang = lvl >= 0.99 ? -Math.PI / 2 : lvl <= 0.01 ? Math.PI / 2 : lvl > 0.5 ? -Math.PI / 4 : lvl < 0.5 ? Math.PI / 4 : 0;
    ctx.rotate(ang);
    ctx.beginPath();
    ctx.moveTo(rr, 0);
    ctx.lineTo(0, -rr);
    ctx.lineTo(0, -rr * 0.45);
    ctx.lineTo(-rr, -rr * 0.45);
    ctx.lineTo(-rr, rr * 0.45);
    ctx.lineTo(0, rr * 0.45);
    ctx.lineTo(0, rr);
    ctx.closePath();
    ctx.fill();
  } else if (/Traffic|Signs|Symbols|Quarters/.test(set)) {
    const n = count;
    let color = col3;
    if (n === 4) color = [IC.black, IC.red, IC.yellow, IC.green][idx];
    ctx.fillStyle = color;
    if (/Signs/.test(set) && lvl <= 0.01) {
      ctx.beginPath();
      ctx.moveTo(cx, cy - rr);
      ctx.lineTo(cx + rr, cy);
      ctx.lineTo(cx, cy + rr);
      ctx.lineTo(cx - rr, cy);
      ctx.fill();
    } else if (/Signs/.test(set) && lvl < 0.99) {
      ctx.beginPath();
      ctx.moveTo(cx, cy - rr);
      ctx.lineTo(cx + rr, cy + rr * 0.8);
      ctx.lineTo(cx - rr, cy + rr * 0.8);
      ctx.fill();
    } else if (/Quarters/.test(set)) {
      ctx.strokeStyle = IC.black;
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.arc(cx, cy, rr, 0, Math.PI * 2);
      ctx.stroke();
      ctx.fillStyle = IC.black;
      ctx.beginPath();
      ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, rr, -Math.PI / 2, -Math.PI / 2 + Math.PI * 2 * lvl);
      ctx.fill();
    } else {
      ctx.beginPath();
      ctx.arc(cx, cy, rr, 0, Math.PI * 2);
      ctx.fill();
      if (/Symbols/.test(set)) {
        ctx.strokeStyle = '#FFFFFF';
        ctx.lineWidth = Math.max(1.5, s * 0.12);
        ctx.beginPath();
        if (lvl >= 0.99) {
          ctx.moveTo(cx - rr * 0.5, cy);
          ctx.lineTo(cx - rr * 0.1, cy + rr * 0.4);
          ctx.lineTo(cx + rr * 0.5, cy - rr * 0.4);
        } else if (lvl <= 0.01) {
          ctx.moveTo(cx - rr * 0.4, cy - rr * 0.4);
          ctx.lineTo(cx + rr * 0.4, cy + rr * 0.4);
          ctx.moveTo(cx + rr * 0.4, cy - rr * 0.4);
          ctx.lineTo(cx - rr * 0.4, cy + rr * 0.4);
        } else {
          ctx.moveTo(cx, cy - rr * 0.5);
          ctx.lineTo(cx, cy + rr * 0.1);
          ctx.moveTo(cx, cy + rr * 0.35);
          ctx.lineTo(cx, cy + rr * 0.45);
        }
        ctx.stroke();
      }
    }
  } else if (/Flags/.test(set)) {
    ctx.fillStyle = col3;
    ctx.fillRect(x + s * 0.2, y + s * 0.15, s * 0.1, s * 0.75);
    ctx.beginPath();
    ctx.moveTo(x + s * 0.3, y + s * 0.15);
    ctx.lineTo(x + s * 0.85, y + s * 0.32);
    ctx.lineTo(x + s * 0.3, y + s * 0.5);
    ctx.fill();
  } else if (/Stars/.test(set)) {
    ctx.fillStyle = IC.yellow;
    ctx.strokeStyle = IC.yellow;
    const star = () => {
      ctx.beginPath();
      for (let i = 0; i < 10; i++) {
        const a = -Math.PI / 2 + (i * Math.PI) / 5;
        const rad = i % 2 === 0 ? rr : rr * 0.45;
        ctx.lineTo(cx + Math.cos(a) * rad, cy + Math.sin(a) * rad);
      }
      ctx.closePath();
    };
    star();
    ctx.stroke();
    if (lvl > 0) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(x, y, s * (lvl >= 0.99 ? 1 : 0.5), s);
      ctx.clip();
      star();
      ctx.fill();
      ctx.restore();
    }
  } else if (/Triangles/.test(set)) {
    ctx.fillStyle = lvl >= 0.99 ? IC.green : lvl <= 0.01 ? IC.red : IC.yellow;
    ctx.beginPath();
    if (lvl >= 0.99) {
      ctx.moveTo(cx, cy - rr);
      ctx.lineTo(cx + rr, cy + rr * 0.6);
      ctx.lineTo(cx - rr, cy + rr * 0.6);
    } else if (lvl <= 0.01) {
      ctx.moveTo(cx, cy + rr);
      ctx.lineTo(cx + rr, cy - rr * 0.6);
      ctx.lineTo(cx - rr, cy - rr * 0.6);
    } else ctx.rect(cx - rr, cy - rr * 0.25, rr * 2, rr * 0.5);
    ctx.fill();
  } else {
    // Rating / Boxes: vertical bars
    const bars = 4;
    const filled = Math.round(lvl * bars);
    for (let i = 0; i < bars; i++) {
      ctx.fillStyle = i < filled ? IC.blue : '#C8C8C8';
      const bh = (s * 0.8 * (i + 1)) / bars;
      ctx.fillRect(x + s * 0.1 + (i * s * 0.8) / bars, y + s * 0.9 - bh, (s * 0.8) / bars - 1, bh);
    }
  }
  ctx.restore();
}

// ---------- text ----------

function textColor(rs: RenderState, style: CellStyle, d: Display, cell: { link?: string } | undefined): string {
  if (d.color) return d.color;
  if (style.fontColor) return adaptColor(style.fontColor, 'text', rs.dark)!;
  if (cell?.link) return rs.pal.hyperlink;
  return rs.pal.text;
}

function drawDecorations(ctx: CanvasRenderingContext2D, style: CellStyle, x: number, baseline: number, w: number, fontPx: number, color: string): void {
  if (!style.underline && !style.strike) return;
  ctx.fillStyle = color;
  const t = Math.max(1, Math.round(fontPx / 14));
  if (style.underline) {
    const y = Math.round(baseline + Math.max(1, fontPx * 0.1));
    ctx.fillRect(x, y, w, t);
    if (style.underline === 'double' || style.underline === 'doubleAccounting') ctx.fillRect(x, y + t + 1, w, t);
  }
  if (style.strike) ctx.fillRect(x, Math.round(baseline - fontPx * 0.3), w, t);
}

interface TextJob {
  r: number;
  c: number;
  x: number;
  y: number;
  w: number;
  h: number;
  info: CellInfo;
  merged: boolean;
}

function drawCellText(ctx: CanvasRenderingContext2D, rs: RenderState, job: TextJob, clipL: number, clipR: number): void {
  const { sheet, vp } = rs;
  const { r, c, x, y, w, h, info } = job;
  const style = info.style;
  const z = vp.z;
  const cell = sheet.getCell(r, c);
  if (rs.editing && rs.editing.r === r && rs.editing.c === c) return;
  const fontPx = ptToPx(style.fontSize ?? DEFAULT_FONT_SIZE, z);
  setFont(ctx, cssFont(style, z));
  const padL = Math.round(2 * z) + (style.indent ? Math.round(style.indent * 9 * z) : 0);
  const padR = Math.round(3 * z);
  const avail = w - padL - padR;
  let avail2 = avail;
  let d = displayOf(sheet, r, c, style, 11, rs.showFormulas);
  const numeric = typeof d.value === 'number' && !rs.showFormulas;
  // General numbers shrink precision to fit, other numbers show ####
  if (numeric && !style.wrap) {
    let tw = ctx.measureText(d.text).width;
    if (tw > avail) {
      if (!style.numFmt && typeof d.value === 'number') {
        const num = d.value;
        for (let n = 10; n >= 1 && tw > avail; n--) {
          d = { ...d, text: formatGeneral(num, n) };
          tw = ctx.measureText(d.text).width;
        }
      }
      // Excel is lenient by a couple of pixels before falling back to ####
      if (tw > avail && tw <= w - 2) avail2 = w - 2;
      else if (tw > avail && !job.merged) {
        const hw = ctx.measureText('#').width || 6;
        d = { ...d, text: '#'.repeat(Math.max(1, Math.floor(avail / hw))), left: undefined, right: undefined };
      }
    }
  }
  const icon = info.cf?.icon;
  const bar = info.cf?.bar;
  if ((icon && !icon.showValue) || (bar && !bar.showValue)) return;
  const color = textColor(rs, style, d, cell);
  ctx.fillStyle = color;
  const vAlign = style.vAlign ?? 'bottom';
  const lineH = fontPx * 1.2;
  const ascent = fontPx * 0.78;
  const descent = fontPx * 0.22;
  const baselineFor = (lines: number, i: number) => {
    const total = lines * lineH;
    if (vAlign === 'top') return y + Math.round(2 * z) + ascent + i * lineH;
    if (vAlign === 'middle' || vAlign === 'justify' || vAlign === 'distributed') return y + (h - total) / 2 + ascent + (lineH - fontPx) / 2 + i * lineH;
    return y + h - Math.round(3 * z) - descent - (lines - 1 - i) * lineH;
  };
  const iconSize = icon ? Math.min(h - 2, 16 * z) : 0;
  ctx.save();
  ctx.beginPath();
  ctx.rect(clipL, y, clipR - clipL, h);
  ctx.clip();
  if (icon) drawIcon(ctx, icon.set, icon.index, icon.count, x + padL, y + (h - iconSize) / 2, iconSize);
  const rot = style.rotation ?? 0;
  if (rot && rot !== 255) {
    const ang = ((rot > 90 ? -(rot - 90) : rot) * Math.PI) / 180;
    const tw = ctx.measureText(d.text).width;
    ctx.translate(x + w / 2, y + h / 2);
    ctx.rotate(-ang);
    ctx.textAlign = 'center';
    ctx.fillText(d.text, 0, fontPx * 0.35);
    drawDecorations(ctx, style, -tw / 2, fontPx * 0.35, tw, fontPx, color);
    ctx.restore();
    ctx.textAlign = 'left';
    return;
  }
  if (rot === 255) {
    // stacked vertical text
    ctx.textAlign = 'center';
    const chars = d.text.split('');
    const total = chars.length * lineH;
    let yy = vAlign === 'top' ? y + ascent + 2 : vAlign === 'middle' ? y + (h - total) / 2 + ascent : y + h - total + ascent - 2;
    for (const ch of chars) {
      ctx.fillText(ch, x + w / 2, yy);
      yy += lineH;
    }
    ctx.textAlign = 'left';
    ctx.restore();
    return;
  }
  if (d.left !== undefined && d.right !== undefined && !style.wrap) {
    // accounting: symbol left, number right
    const bl = baselineFor(1, 0);
    ctx.fillText(d.left, x + padL + iconSize, bl);
    const rw = ctx.measureText(d.right).width;
    ctx.fillText(d.right, x + w - padR - rw, bl);
    ctx.restore();
    return;
  }
  const hAlign = style.hAlign;
  const fill = hAlign === 'fill';
  if (style.wrap || hAlign === 'justify' || vAlign === 'justify' || d.text.includes('\n') && style.wrap) {
    const lines = wrapCanvas(ctx, d.text, Math.max(4, avail));
    for (let i = 0; i < lines.length; i++) {
      const lw = ctx.measureText(lines[i]).width;
      let tx = x + padL;
      if (d.hAlign === 'center') tx = x + (w - lw) / 2;
      else if (d.hAlign === 'right') tx = x + w - padR - lw;
      const bl = baselineFor(lines.length, i);
      ctx.fillText(lines[i], tx, bl);
      drawDecorations(ctx, style, tx, bl, lw, fontPx, color);
    }
    ctx.restore();
    return;
  }
  let text = d.text.replace(/\n/g, ' ');
  if (fill && text) {
    const one = ctx.measureText(text).width || 1;
    text = text.repeat(Math.max(1, Math.floor(avail / one)));
  }
  const tw = ctx.measureText(text).width;
  let tx: number;
  if (d.hAlign === 'center') tx = x + (w - tw) / 2;
  else if (d.hAlign === 'right') tx = x + w - (avail2 > avail ? 1 : padR) - tw;
  else tx = x + padL + (icon ? iconSize + 2 : 0);
  if (d.hAlign === 'right' && icon) tx = Math.max(tx, x + padL + iconSize + 2);
  const bl = baselineFor(1, 0);
  ctx.fillText(text, tx, bl);
  drawDecorations(ctx, style, tx, bl, tw, fontPx, color);
  ctx.restore();
}

function wrapCanvas(ctx: CanvasRenderingContext2D, text: string, width: number): string[] {
  const out: string[] = [];
  for (const para of text.split('\n')) {
    const words = para.split(/(\s+)/);
    let line = '';
    for (const w of words) {
      const cand = line + w;
      if (line && ctx.measureText(cand.trimEnd()).width > width) {
        out.push(line.trimEnd());
        line = w.trimStart();
        while (line.length > 1 && ctx.measureText(line).width > width) {
          let k = line.length - 1;
          while (k > 1 && ctx.measureText(line.slice(0, k)).width > width) k--;
          out.push(line.slice(0, k));
          line = line.slice(k);
        }
      } else line = cand;
    }
    out.push(line);
  }
  return out;
}

// ---------- panes ----------

function drawPane(ctx: CanvasRenderingContext2D, rs: RenderState, cp: Pane, rp: Pane): void {
  const { sheet, vp, pal } = rs;
  const z = vp.z;
  const [r0, r1] = paneRange(vp, rp, 'row');
  const [c0, c1] = paneRange(vp, cp, 'col');
  const px = cp.start;
  const py = rp.start;
  const pw = cp.size;
  const ph = rp.size;
  if (pw <= 0 || ph <= 0) return;
  ctx.save();
  currentFont = ''; // save/restore resets ctx.font behind the cache's back
  ctx.beginPath();
  ctx.rect(px, py, pw, ph);
  ctx.clip();
  ctx.fillStyle = pal.bg;
  ctx.fillRect(px, py, pw, ph);

  const xs: number[] = [];
  const ys: number[] = [];
  for (let c = c0; c <= c1 + 1; c++) xs[c - c0] = snap(colX(vp, cp, c));
  for (let r = r0; r <= r1 + 1; r++) ys[r - r0] = snap(rowY(vp, rp, r));
  const X = (c: number) => xs[c - c0];
  const Y = (r: number) => ys[r - r0];
  const g = 1 / dpr; // one device pixel

  // gridlines
  const showGrid = rs.print ? rs.print.gridlines : sheet.showGridlines;
  if (showGrid) {
    ctx.fillStyle = pal.grid;
    for (let c = c0; c <= c1; c++) {
      if (vp.cols.size(c) === 0) continue;
      ctx.fillRect(X(c + 1) - g, py, g, ph);
    }
    for (let r = r0; r <= r1; r++) {
      if (vp.rows.size(r) === 0) continue;
      ctx.fillRect(px, Y(r + 1) - g, pw, g);
    }
  }

  // merged areas intersecting this pane
  const merges = sheet.merges.filter((m) => m.r2 >= r0 && m.r1 <= r1 && m.c2 >= c0 && m.c1 <= c1);
  const mergeRect = (m: Range) => {
    const x = snap(colX(vp, cp, m.c1));
    const y = snap(rowY(vp, rp, m.r1));
    const x2 = snap(colX(vp, cp, m.c2 + 1));
    const y2 = snap(rowY(vp, rp, m.r2 + 1));
    return { x, y, w: x2 - x, h: y2 - y };
  };
  const covered = (r: number, c: number) => {
    for (const m of merges) if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) return m;
    return undefined;
  };

  // pass 1: fills + collect jobs
  const infos = new Map<number, CellInfo>();
  const key = (r: number, c: number) => r * 16384 + c;
  const infoAt = (r: number, c: number) => {
    const k = key(r, c);
    let i = infos.get(k);
    if (!i) infos.set(k, (i = resolveCell(rs, r, c)));
    return i;
  };
  const rowHasStyle = (r: number) => sheet.rows.has(r) || sheet.rowStyles.size > 0 || sheet.colStyles.size > 0 || sheet.tables.length > 0 || sheet.conditionalFormats.length > 0;
  for (let r = r0; r <= r1; r++) {
    if (vp.rows.size(r) === 0 || !rowHasStyle(r)) continue;
    for (let c = c0; c <= c1; c++) {
      if (vp.cols.size(c) === 0) continue;
      if (merges.length && covered(r, c)) continue;
      const cell = sheet.getCell(r, c);
      if (!cell && !sheet.rowStyles.has(r) && !sheet.colStyles.has(c) && !sheet.tables.length && !sheet.conditionalFormats.length) continue;
      const info = infoAt(r, c);
      if (info.fill) {
        ctx.fillStyle = adaptColor(info.fill, 'fill', rs.dark)!;
        ctx.fillRect(X(c) - g, Y(r) - g, X(c + 1) - X(c) + g, Y(r + 1) - Y(r) + g);
      }
      if (info.cf?.bar) drawBar(ctx, info.cf.bar, X(c), Y(r), X(c + 1) - X(c) - g, Y(r + 1) - Y(r) - g, z);
    }
  }
  for (const m of merges) {
    const rc = mergeRect(m);
    const info = infoAt(m.r1, m.c1);
    ctx.fillStyle = info.fill ? adaptColor(info.fill, 'fill', rs.dark)! : pal.bg;
    ctx.fillRect(rc.x, rc.y, rc.w - g, rc.h - g);
    if (showGrid && !info.fill) {
      // keep outer gridlines
      ctx.fillStyle = pal.grid;
      ctx.fillRect(rc.x + rc.w - g, rc.y, g, rc.h);
      ctx.fillRect(rc.x, rc.y + rc.h - g, rc.w, g);
    }
  }

  // pass 2: borders
  for (let r = r0; r <= r1; r++) {
    if (vp.rows.size(r) === 0 || !rowHasStyle(r)) continue;
    for (let c = c0; c <= c1; c++) {
      if (vp.cols.size(c) === 0) continue;
      const k = key(r, c);
      const info = infos.get(k) ?? (sheet.getCell(r, c) || sheet.rowStyles.has(r) || sheet.colStyles.has(c) || sheet.tables.length ? infoAt(r, c) : undefined);
      if (!info) continue;
      const s = info.style;
      const tb = info.tableBorders;
      const top = s.bTop ?? tb?.bTop;
      const bottom = s.bBottom ?? tb?.bBottom;
      const left = s.bLeft ?? tb?.bLeft;
      const right = s.bRight ?? tb?.bRight;
      if (!top && !bottom && !left && !right && !s.bDiagUp && !s.bDiagDown) continue;
      const m = merges.length ? covered(r, c) : undefined;
      const x1 = X(c) - g;
      const x2 = X(c + 1) - g;
      const y1 = Y(r) - g;
      const y2 = Y(r + 1) - g;
      if (top && (!m || r === m.r1)) drawEdge(ctx, top, 'h', y1, x1, x2 + g, z, rs.dark);
      if (bottom && (!m || r === m.r2)) drawEdge(ctx, bottom, 'h', y2, x1, x2 + g, z, rs.dark);
      if (left && (!m || c === m.c1)) drawEdge(ctx, left, 'v', x1, y1, y2 + g, z, rs.dark);
      if (right && (!m || c === m.c2)) drawEdge(ctx, right, 'v', x2, y1, y2 + g, z, rs.dark);
      if ((s.bDiagDown || s.bDiagUp) && !m) {
        ctx.save();
        ctx.lineWidth = 1;
        if (s.bDiagDown) {
          ctx.strokeStyle = adaptColor(s.bDiagDown.color ?? '#000000', 'text', rs.dark)!;
          ctx.beginPath();
          ctx.moveTo(x1, y1);
          ctx.lineTo(x2, y2);
          ctx.stroke();
        }
        if (s.bDiagUp) {
          ctx.strokeStyle = adaptColor(s.bDiagUp.color ?? '#000000', 'text', rs.dark)!;
          ctx.beginPath();
          ctx.moveTo(x1, y2);
          ctx.lineTo(x2, y1);
          ctx.stroke();
        }
        ctx.restore();
      }
    }
  }

  // pass 3: text (with overflow into neighbouring empty cells)
  ctx.textBaseline = 'alphabetic';
  const extraL = 30;
  const extraR = 30;
  const cStart = Math.max(cp.first, c0 - extraL);
  const cEnd = Math.min(cp.last, c1 + extraR);
  for (let r = r0; r <= r1; r++) {
    if (vp.rows.size(r) === 0) continue;
    const row = sheet.rows.get(r);
    if (!row) continue;
    const cols: number[] = [];
    if (row.size < (cEnd - cStart) * 2) {
      for (const c of row.keys()) if (c >= cStart && c <= cEnd) cols.push(c);
      cols.sort((a, b) => a - b);
    } else for (let c = cStart; c <= cEnd; c++) if (row.has(c)) cols.push(c);
    for (const c of cols) {
      if (vp.cols.size(c) === 0) continue;
      const cell = row.get(c)!;
      if (cell.v === undefined && cell.f === undefined) continue;
      if (cell.v === null || cell.v === '') if (cell.f === undefined) continue;
      if (merges.length && covered(r, c) && !(covered(r, c)!.r1 === r && covered(r, c)!.c1 === c)) continue;
      const m = sheet.merges.length ? sheet.mergeAt(r, c) : undefined;
      if (m && (m.r1 !== r || m.c1 !== c)) continue;
      const info = infoAt(r, c);
      let x: number;
      let y: number;
      let w: number;
      let h: number;
      if (m) {
        x = snap(colX(vp, cp, m.c1));
        y = snap(rowY(vp, rp, m.r1));
        w = snap(colX(vp, cp, m.c2 + 1)) - x - g;
        h = snap(rowY(vp, rp, m.r2 + 1)) - y - g;
      } else {
        x = snap(colX(vp, cp, c));
        y = Y(r);
        w = snap(colX(vp, cp, c + 1)) - x - g;
        h = Y(r + 1) - y - g;
      }
      if (m && (m.c1 > c1 || m.c2 < c0)) continue;
      // text overflow bounds
      let clipL = x;
      let clipR = x + w;
      const s = info.style;
      const isText = !rs.showFormulas ? typeof (cell.f ? undefined : cell.v) === 'string' || (cell.f !== undefined && !s.wrap) : true;
      if (!m && !s.wrap && isText && (!s.rotation || s.rotation === 0) && s.hAlign !== 'fill') {
        const d = displayOf(sheet, r, c, s, 11, rs.showFormulas);
        if (typeof d.value !== 'number' || rs.showFormulas) {
          setFont(ctx, cssFont(s, z));
          const tw = ctx.measureText(d.text).width + 4 * z;
          if (tw > w) {
            const al = d.hAlign;
            if (al === 'left' || al === 'center') {
              let cc = c + 1;
              let right = x + w;
              const need = al === 'center' ? x + w / 2 + tw / 2 : x + tw;
              while (right < need && cc <= cp.last && !hasValue(sheet, r, cc) && !sheet.mergeAt(r, cc)) {
                right += vp.cols.size(cc) * z;
                cc++;
              }
              clipR = Math.max(clipR, right);
            }
            if (al === 'right' || al === 'center') {
              let cc = c - 1;
              let left = x;
              const need = al === 'center' ? x + w / 2 - tw / 2 : x + w - tw;
              while (left > need && cc >= cp.first && !hasValue(sheet, r, cc) && !sheet.mergeAt(r, cc)) {
                left -= vp.cols.size(cc) * z;
                cc--;
              }
              clipL = Math.min(clipL, left);
            }
          }
        }
      }
      if (clipR < px || clipL > px + pw) continue;
      drawCellText(ctx, rs, { r, c, x, y, w, h, info, merged: !!m }, clipL, clipR);
    }
  }

  // sparklines
  if (sheet.sparklines.length) {
    for (const g2 of sheet.sparklines)
      for (const it of g2.items) {
        if (it.r < r0 || it.r > r1 || it.c < c0 || it.c > c1) continue;
        drawSparkline(ctx, rs, g2, it, X(it.c), Y(it.r), X(it.c + 1) - X(it.c) - g, Y(it.r + 1) - Y(it.r) - g);
      }
  }

  // note indicators
  for (let r = r0; r <= r1; r++) {
    const row = sheet.rows.get(r);
    if (!row) continue;
    for (const [c, cell] of row) {
      if (!cell.note || c < c0 || c > c1) continue;
      const m = sheet.mergeAt(r, c);
      const xr = m ? snap(colX(vp, cp, m.c2 + 1)) : X(c + 1);
      const yt = Y(r);
      const s = Math.round(6 * Math.max(0.75, z));
      ctx.fillStyle = cell.note.threaded ? '#7B61FF' : pal.noteMark;
      ctx.beginPath();
      ctx.moveTo(xr - g - s, yt);
      ctx.lineTo(xr - g, yt);
      ctx.lineTo(xr - g, yt + s);
      ctx.fill();
    }
  }

  // invalid data circles
  if (rs.invalid) {
    ctx.strokeStyle = '#FF0000';
    ctx.lineWidth = 1.5;
    for (const k of rs.invalid) {
      const [r, c] = k.split(',').map(Number);
      if (r < r0 || r > r1 || c < c0 || c > c1) continue;
      ctx.beginPath();
      ctx.ellipse((X(c) + X(c + 1)) / 2, (Y(r) + Y(r + 1)) / 2, (X(c + 1) - X(c)) / 2 + 3, (Y(r + 1) - Y(r)) / 2 + 2, 0, 0, Math.PI * 2);
      ctx.stroke();
    }
  }

  // page breaks
  if (rs.pageBreaks) {
    ctx.save();
    ctx.strokeStyle = pal.pageBreak;
    ctx.lineWidth = 1;
    ctx.setLineDash([4, 3]);
    ctx.beginPath();
    for (const r of rs.pageBreaks.rows) {
      if (r < r0 || r > r1 + 1) continue;
      const y = Y(r) - 0.5;
      ctx.moveTo(px, y);
      ctx.lineTo(px + pw, y);
    }
    for (const c of rs.pageBreaks.cols) {
      if (c < c0 || c > c1 + 1) continue;
      const x = X(c) - 0.5;
      ctx.moveTo(x, py);
      ctx.lineTo(x, py + ph);
    }
    ctx.stroke();
    ctx.restore();
  }

  if (!rs.print) drawOverlays(ctx, rs, cp, rp, r0, r1, c0, c1);
  ctx.restore();
  currentFont = '';
}

function drawBar(ctx: CanvasRenderingContext2D, bar: NonNullable<CFResult['bar']>, x: number, y: number, w: number, h: number, z: number): void {
  const pad = Math.round(2 * z);
  const bx = x + pad + (w - 2 * pad) * bar.from;
  const bw = (w - 2 * pad) * (bar.to - bar.from);
  const by = y + pad;
  const bh = h - 2 * pad;
  if (bw <= 0) return;
  if (bar.gradient) {
    const grad = ctx.createLinearGradient(bx, 0, bx + bw, 0);
    if (bar.negative) {
      grad.addColorStop(0, '#FFFFFF');
      grad.addColorStop(1, bar.color);
    } else {
      grad.addColorStop(0, bar.color);
      grad.addColorStop(1, '#FFFFFF');
    }
    ctx.fillStyle = grad;
  } else ctx.fillStyle = bar.color;
  ctx.fillRect(bx, by, bw, bh);
  ctx.strokeStyle = bar.color;
  ctx.lineWidth = 1;
  if (bar.gradient) ctx.strokeRect(bx + 0.5, by + 0.5, bw - 1, bh - 1);
  if (bar.axis !== undefined) {
    ctx.save();
    ctx.setLineDash([2, 2]);
    ctx.strokeStyle = '#000000';
    const ax = x + pad + (w - 2 * pad) * bar.axis;
    ctx.beginPath();
    ctx.moveTo(ax + 0.5, y);
    ctx.lineTo(ax + 0.5, y + h);
    ctx.stroke();
    ctx.restore();
  }
}

function drawSparkline(ctx: CanvasRenderingContext2D, rs: RenderState, g: Sheet['sparklines'][0], it: { r: number; c: number; ref: string }, x: number, y: number, w: number, h: number): void {
  const vals = sparklineValues(rs.wb, rs.sheet, it.ref);
  if (!vals.length) return;
  const nums = vals.filter((v): v is number => typeof v === 'number');
  if (!nums.length) return;
  const pad = 3;
  const min = Math.min(...nums, g.type === 'column' ? 0 : Infinity);
  const max = Math.max(...nums, g.type === 'column' ? 0 : -Infinity);
  const span = max - min || 1;
  const ix = (i: number) => x + pad + ((w - 2 * pad) * (vals.length === 1 ? 0.5 : i / (vals.length - 1)));
  const iy = (v: number) => y + h - pad - ((v - min) / span) * (h - 2 * pad);
  ctx.save();
  ctx.beginPath();
  ctx.rect(x, y, w, h);
  ctx.clip();
  if (g.type === 'line') {
    ctx.strokeStyle = g.color;
    ctx.lineWidth = 1.25;
    ctx.beginPath();
    let started = false;
    vals.forEach((v, i) => {
      if (typeof v !== 'number') return;
      if (!started) ctx.moveTo(ix(i), iy(v));
      else ctx.lineTo(ix(i), iy(v));
      started = true;
    });
    ctx.stroke();
    const hi = Math.max(...nums);
    const lo = Math.min(...nums);
    vals.forEach((v, i) => {
      if (typeof v !== 'number') return;
      const isHi = g.highPoint && v === hi;
      const isLo = g.lowPoint && v === lo;
      if (g.markers || isHi || isLo) {
        ctx.fillStyle = isHi ? '#00A650' : isLo ? '#E53935' : g.color;
        ctx.beginPath();
        ctx.arc(ix(i), iy(v), 2, 0, Math.PI * 2);
        ctx.fill();
      }
    });
  } else {
    const bw = Math.max(1, (w - 2 * pad) / vals.length - 1);
    vals.forEach((v, i) => {
      if (typeof v !== 'number') return;
      const bx = x + pad + i * ((w - 2 * pad) / vals.length);
      if (g.type === 'winloss') {
        ctx.fillStyle = v < 0 ? g.negColor ?? '#E53935' : g.color;
        const mid = y + h / 2;
        if (v > 0) ctx.fillRect(bx, y + pad, bw, mid - y - pad - 1);
        else if (v < 0) ctx.fillRect(bx, mid + 1, bw, y + h - pad - mid - 1);
      } else {
        ctx.fillStyle = v < 0 ? g.negColor ?? '#E53935' : g.color;
        const zero = iy(0);
        const top = iy(v);
        ctx.fillRect(bx, Math.min(zero, top), bw, Math.max(1, Math.abs(zero - top)));
      }
    });
  }
  ctx.restore();
}

// ---------- overlays ----------

function drawOverlays(ctx: CanvasRenderingContext2D, rs: RenderState, cp: Pane, rp: Pane, r0: number, r1: number, c0: number, c1: number): void {
  const { vp, pal, sel, sheet } = rs;
  const z = vp.z;
  const rect = (rg: Range) => {
    const x = snap(colX(vp, cp, rg.c1));
    const y = snap(rowY(vp, rp, rg.r1));
    const x2 = snap(colX(vp, cp, Math.min(rg.c2, vp.cols.count - 1) + 1));
    const y2 = snap(rowY(vp, rp, Math.min(rg.r2, vp.rows.count - 1) + 1));
    return { x, y, w: x2 - x, h: y2 - y };
  };
  const visible = (rg: Range) => rg.r2 >= r0 && rg.r1 <= r1 && rg.c2 >= c0 && rg.c1 <= c1;

  // AutoFilter buttons
  const af = sheet.autoFilter;
  if (af && rs.filterButtons && af.range.r1 >= r0 && af.range.r1 <= r1) {
    for (let c = Math.max(c0, af.range.c1); c <= Math.min(c1, af.range.c2); c++) {
      if (vp.cols.size(c) === 0) continue;
      const r = rect({ r1: af.range.r1, r2: af.range.r1, c1: c, c2: c });
      drawFilterButton(ctx, rs, r.x + r.w - Math.round(17 * z), r.y + r.h - Math.round(17 * z), Math.round(16 * z), !!af.filters[c], af.sort?.col === c ? (af.sort.desc ? 'desc' : 'asc') : null);
    }
  }
  // table header filter buttons (tables without a sheet-level filter)
  for (const t of sheet.tables) {
    if (!t.headerRow || !t.showFilterButton || (af && af.range.r1 === t.range.r1 && af.range.c1 === t.range.c1)) continue;
    if (t.range.r1 < r0 || t.range.r1 > r1) continue;
    for (let c = Math.max(c0, t.range.c1); c <= Math.min(c1, t.range.c2); c++) {
      const r = rect({ r1: t.range.r1, r2: t.range.r1, c1: c, c2: c });
      drawFilterButton(ctx, rs, r.x + r.w - Math.round(17 * z), r.y + r.h - Math.round(17 * z), Math.round(16 * z), false, null);
    }
  }

  // formula reference highlights
  for (const ref of rs.refs) {
    if (!visible(ref.range)) continue;
    const rg = rect(ref.range);
    ctx.fillStyle = ref.color + '1F';
    ctx.fillRect(rg.x, rg.y, rg.w, rg.h);
    ctx.strokeStyle = ref.color;
    ctx.lineWidth = 2;
    ctx.strokeRect(rg.x, rg.y, rg.w - 1, rg.h - 1);
    ctx.fillStyle = ref.color;
    const s = 5;
    for (const [hx, hy] of [
      [rg.x, rg.y],
      [rg.x + rg.w - 1, rg.y],
      [rg.x, rg.y + rg.h - 1],
      [rg.x + rg.w - 1, rg.y + rg.h - 1],
    ])
      ctx.fillRect(hx - s / 2, hy - s / 2, s, s);
  }

  // selection
  const multi = sel.ranges.length > 1;
  const active = sel.active;
  const activeMerge = sheet.mergeAt(active.r, active.c);
  const activeRange: Range = activeMerge ?? { r1: active.r, c1: active.c, r2: active.r, c2: active.c };
  ctx.fillStyle = pal.selFill;
  for (const rg of sel.ranges) {
    if (!visible(rg)) continue;
    const R = rect(rg);
    const single = (rg.r1 === rg.r2 && rg.c1 === rg.c2) || (activeMerge && rg.r1 === activeMerge.r1 && rg.c1 === activeMerge.c1 && rg.r2 === activeMerge.r2 && rg.c2 === activeMerge.c2);
    if (single) continue;
    // shade all but active cell
    const A = rect(activeRange);
    ctx.save();
    ctx.beginPath();
    ctx.rect(R.x, R.y, R.w, R.h);
    if (active.r >= rg.r1 && active.r <= rg.r2 && active.c >= rg.c1 && active.c <= rg.c2) ctx.rect(A.x + A.w, A.y, -A.w, A.h);
    ctx.fill('evenodd');
    ctx.restore();
  }
  const bw = Math.max(2, Math.round(2 * Math.min(1.5, z)));
  if (!multi) {
    const rg = sel.ranges[0];
    if (visible(rg)) {
      const R = rect(rg);
      ctx.fillStyle = pal.accent;
      // 2px border drawn inside/outside the gridline like Excel
      ctx.fillRect(R.x - 1 - 1 / dpr, R.y - 1 - 1 / dpr, R.w + bw - 1 / dpr, bw);
      ctx.fillRect(R.x - 1 - 1 / dpr, R.y + R.h - 1 - 1 / dpr, R.w + bw - 1 / dpr, bw);
      ctx.fillRect(R.x - 1 - 1 / dpr, R.y - 1 - 1 / dpr, bw, R.h + bw - 1 / dpr);
      ctx.fillRect(R.x + R.w - 1 - 1 / dpr, R.y - 1 - 1 / dpr, bw, R.h + bw - 1 / dpr);
      // fill handle
      if (!isFullCols(rg) || rg.c1 === rg.c2 || true) {
        const hs = Math.round(6 * Math.max(0.85, Math.min(1.3, z)));
        const hx = R.x + R.w - 1 - hs / 2;
        const hy = R.y + R.h - 1 - hs / 2;
        if (!(isFullCols(rg) && isFullRows(rg))) {
          ctx.fillStyle = pal.bg;
          ctx.fillRect(hx - 1, hy - 1, hs + 2, hs + 2);
          ctx.fillStyle = pal.accent;
          ctx.fillRect(hx, hy, hs, hs);
        }
      }
    }
  } else {
    // multi-range: thin border around active cell only
    if (visible(activeRange)) {
      const A = rect(activeRange);
      ctx.strokeStyle = pal.accent;
      ctx.lineWidth = 1;
      ctx.strokeRect(A.x - 0.5, A.y - 0.5, A.w, A.h);
    }
  }

  // fill-handle drag preview
  if (rs.fillPreview && visible(rs.fillPreview)) {
    const R = rect(rs.fillPreview);
    ctx.save();
    ctx.strokeStyle = '#7F7F7F';
    ctx.lineWidth = 1;
    ctx.setLineDash([3, 2]);
    ctx.strokeRect(R.x - 0.5, R.y - 0.5, R.w, R.h);
    ctx.restore();
  }

  // format painter source
  if (rs.painterRange && visible(rs.painterRange)) marchingAnts(ctx, rect(rs.painterRange), 0, pal.accent);

  // copy marquee
  if (rs.clip && visible(rs.clip.range)) marchingAnts(ctx, rect(rs.clip.range), rs.clip.phase, pal.accent);

  // validation list arrow
  if (rs.listArrow && rs.listArrow.r >= r0 && rs.listArrow.r <= r1 && rs.listArrow.c + 1 >= c0 && rs.listArrow.c <= c1) {
    const m = sheet.mergeAt(rs.listArrow.r, rs.listArrow.c);
    const R = rect(m ?? { r1: rs.listArrow.r, r2: rs.listArrow.r, c1: rs.listArrow.c, c2: rs.listArrow.c });
    const s = Math.round(Math.min(R.h, 17 * z));
    const x = R.x + R.w + 1;
    const y = R.y + R.h - s;
    ctx.fillStyle = pal.filterBtnBg;
    ctx.fillRect(x, y, s, s);
    ctx.strokeStyle = pal.filterBtnBorder;
    ctx.lineWidth = 1;
    ctx.strokeRect(x + 0.5, y + 0.5, s - 1, s - 1);
    drawChevron(ctx, x + s / 2, y + s / 2, pal.headerText);
  }
}

function drawChevron(ctx: CanvasRenderingContext2D, cx: number, cy: number, color: string): void {
  ctx.fillStyle = color;
  ctx.beginPath();
  ctx.moveTo(cx - 3.5, cy - 1.5);
  ctx.lineTo(cx + 3.5, cy - 1.5);
  ctx.lineTo(cx, cy + 2.5);
  ctx.closePath();
  ctx.fill();
}

function drawFilterButton(ctx: CanvasRenderingContext2D, rs: RenderState, x: number, y: number, s: number, active: boolean, sort: 'asc' | 'desc' | null): void {
  const { pal } = rs;
  ctx.fillStyle = pal.filterBtnBg;
  ctx.fillRect(x, y, s, s);
  ctx.strokeStyle = pal.filterBtnBorder;
  ctx.lineWidth = 1;
  ctx.strokeRect(x + 0.5, y + 0.5, s - 1, s - 1);
  const col = pal.headerText;
  if (active) {
    // funnel
    ctx.fillStyle = col;
    ctx.beginPath();
    ctx.moveTo(x + s * 0.2, y + s * 0.25);
    ctx.lineTo(x + s * 0.65, y + s * 0.25);
    ctx.lineTo(x + s * 0.47, y + s * 0.5);
    ctx.lineTo(x + s * 0.47, y + s * 0.78);
    ctx.lineTo(x + s * 0.38, y + s * 0.7);
    ctx.lineTo(x + s * 0.38, y + s * 0.5);
    ctx.closePath();
    ctx.fill();
    drawChevronSmall(ctx, x + s * 0.78, y + s * 0.62, col);
  } else if (sort) {
    ctx.fillStyle = col;
    ctx.fillRect(x + s * 0.3, y + s * 0.2, 1, s * 0.6);
    ctx.beginPath();
    if (sort === 'asc') {
      ctx.moveTo(x + s * 0.3 - 2.5, y + s * 0.35);
      ctx.lineTo(x + s * 0.3 + 3, y + s * 0.35);
      ctx.lineTo(x + s * 0.3 + 0.5, y + s * 0.2);
    } else {
      ctx.moveTo(x + s * 0.3 - 2.5, y + s * 0.65);
      ctx.lineTo(x + s * 0.3 + 3, y + s * 0.65);
      ctx.lineTo(x + s * 0.3 + 0.5, y + s * 0.8);
    }
    ctx.fill();
    drawChevronSmall(ctx, x + s * 0.68, y + s * 0.5, col);
  } else drawChevron(ctx, x + s / 2, y + s / 2, col);
}

function drawChevronSmall(ctx: CanvasRenderingContext2D, cx: number, cy: number, color: string): void {
  ctx.fillStyle = color;
  ctx.beginPath();
  ctx.moveTo(cx - 2.5, cy - 1);
  ctx.lineTo(cx + 2.5, cy - 1);
  ctx.lineTo(cx, cy + 2);
  ctx.closePath();
  ctx.fill();
}

function marchingAnts(ctx: CanvasRenderingContext2D, R: { x: number; y: number; w: number; h: number }, phase: number, color: string): void {
  ctx.save();
  ctx.lineWidth = 2;
  ctx.strokeStyle = '#FFFFFF';
  ctx.strokeRect(R.x - 1, R.y - 1, R.w, R.h);
  ctx.strokeStyle = color;
  ctx.setLineDash([4, 4]);
  ctx.lineDashOffset = -phase;
  ctx.strokeRect(R.x - 1, R.y - 1, R.w, R.h);
  ctx.restore();
}

// ---------- headers ----------

function selectedCols(sel: Selection): { any: Set<number>; full: Set<number> } {
  const any = new Set<number>();
  const full = new Set<number>();
  for (const rg of sel.ranges) {
    const n = Math.min(rg.c2, rg.c1 + 16384);
    for (let c = rg.c1; c <= n; c++) {
      any.add(c);
      if (isFullCols(rg)) full.add(c);
    }
  }
  return { any, full };
}

function drawColHeaders(ctx: CanvasRenderingContext2D, rs: RenderState, cp: Pane, sc: { any: Set<number>; full: Set<number> }): void {
  const { vp, pal } = rs;
  const z = vp.z;
  const [c0, c1] = paneRange(vp, cp, 'col');
  const g = 1 / dpr;
  ctx.save();
  ctx.beginPath();
  ctx.rect(cp.start, 0, cp.size, vp.hh);
  ctx.clip();
  ctx.fillStyle = pal.headerBg;
  ctx.fillRect(cp.start, 0, cp.size, vp.hh);
  setFont(ctx, `${Math.round(11 * 96 / 72 * z * 0.97)}px "Segoe UI", "Calibri", "Carlito", Arial, sans-serif`);
  ctx.textBaseline = 'middle';
  ctx.textAlign = 'center';
  for (let c = c0; c <= c1; c++) {
    const w = vp.cols.size(c) * z;
    const x = snap(colX(vp, cp, c));
    const x2 = snap(colX(vp, cp, c + 1));
    if (w === 0) {
      // hidden column: double line marker
      ctx.fillStyle = pal.headerLine;
      ctx.fillRect(x - 2 * g, 0, g, vp.hh);
      continue;
    }
    const isSel = sc.any.has(c);
    const isFull = sc.full.has(c);
    if (isSel) {
      ctx.fillStyle = isFull ? pal.headerFullBg : pal.headerSelBg;
      ctx.fillRect(x, 0, x2 - x, vp.hh);
    }
    ctx.fillStyle = pal.headerLine;
    ctx.fillRect(x2 - g, 0, g, vp.hh);
    if (isSel && !isFull) {
      ctx.fillStyle = pal.accent;
      ctx.fillRect(x - g, vp.hh - Math.round(2 * Math.max(1, z)), x2 - x + g, Math.round(2 * Math.max(1, z)));
    }
    ctx.fillStyle = isFull ? pal.headerFullText : isSel ? pal.headerSelText : pal.headerText;
    if (x2 - x > 6) ctx.fillText(colToName(c), (x + x2) / 2, vp.hh / 2 + 0.5);
  }
  ctx.fillStyle = pal.headerLine;
  ctx.fillRect(cp.start, vp.hh - g, cp.size, g);
  ctx.restore();
  currentFont = '';
  ctx.textAlign = 'left';
}

function drawRowHeaders(ctx: CanvasRenderingContext2D, rs: RenderState, rp: Pane, sel: Selection): void {
  const { vp, pal } = rs;
  const z = vp.z;
  const [r0, r1] = paneRange(vp, rp, 'row');
  const g = 1 / dpr;
  const selRows = (r: number) => {
    let any = false;
    let full = false;
    for (const rg of sel.ranges)
      if (r >= rg.r1 && r <= rg.r2) {
        any = true;
        if (isFullRows(rg)) full = true;
      }
    return { any, full };
  };
  ctx.save();
  ctx.beginPath();
  ctx.rect(0, rp.start, vp.hw, rp.size);
  ctx.clip();
  ctx.fillStyle = pal.headerBg;
  ctx.fillRect(0, rp.start, vp.hw, rp.size);
  setFont(ctx, `${Math.round(11 * 96 / 72 * z * 0.97)}px "Segoe UI", "Calibri", "Carlito", Arial, sans-serif`);
  ctx.textBaseline = 'middle';
  ctx.textAlign = 'center';
  for (let r = r0; r <= r1; r++) {
    const h = vp.rows.size(r) * z;
    const y = snap(rowY(vp, rp, r));
    const y2 = snap(rowY(vp, rp, r + 1));
    if (h === 0) {
      ctx.fillStyle = pal.headerLine;
      ctx.fillRect(0, y - 2 * g, vp.hw, g);
      continue;
    }
    const { any, full } = selRows(r);
    if (any) {
      ctx.fillStyle = full ? pal.headerFullBg : pal.headerSelBg;
      ctx.fillRect(0, y, vp.hw, y2 - y);
    }
    ctx.fillStyle = pal.headerLine;
    ctx.fillRect(0, y2 - g, vp.hw, g);
    if (any && !full) {
      ctx.fillStyle = pal.accent;
      const bw = Math.round(2 * Math.max(1, z));
      ctx.fillRect(vp.hw - bw, y - g, bw, y2 - y + g);
    }
    ctx.fillStyle = full ? pal.headerFullText : any ? pal.headerSelText : pal.headerText;
    if (y2 - y > 6) ctx.fillText(String(r + 1), vp.hw / 2, (y + y2) / 2 + 0.5);
  }
  ctx.fillStyle = pal.headerLine;
  ctx.fillRect(vp.hw - g, rp.start, g, rp.size);
  ctx.restore();
  currentFont = '';
  ctx.textAlign = 'left';
}

// ---------- entry ----------

export function renderGrid(ctx: CanvasRenderingContext2D, rs: RenderState): void {
  dpr = rs.dpr;
  currentFont = '';
  const { vp, pal, sheet } = rs;
  const ox = rs.print ? rs.print.origin.x : 0;
  const oy = rs.print ? rs.print.origin.y : 0;
  ctx.setTransform(dpr, 0, 0, dpr, ox * dpr, oy * dpr);
  ctx.fillStyle = pal.bg;
  ctx.fillRect(0, 0, vp.width, vp.height);
  for (const rp of vp.rowPanes) for (const cp of vp.colPanes) drawPane(ctx, rs, cp, rp);
  if (rs.print ? vp.hw > 0 : sheet.showHeaders) {
    const sc = selectedCols(rs.sel);
    for (const cp of vp.colPanes) drawColHeaders(ctx, rs, cp, sc);
    for (const rp of vp.rowPanes) drawRowHeaders(ctx, rs, rp, rs.sel);
    // corner
    ctx.fillStyle = pal.headerBg;
    ctx.fillRect(0, 0, vp.hw, vp.hh);
    ctx.fillStyle = pal.headerLine;
    ctx.fillRect(vp.hw - 1 / dpr, 0, 1 / dpr, vp.hh);
    ctx.fillRect(0, vp.hh - 1 / dpr, vp.hw, 1 / dpr);
    ctx.fillStyle = pal.cornerTriangle;
    ctx.beginPath();
    const s = Math.round(Math.min(vp.hw, vp.hh) * 0.55);
    ctx.moveTo(vp.hw - 4, vp.hh - 4);
    ctx.lineTo(vp.hw - 4, vp.hh - 4 - s);
    ctx.lineTo(vp.hw - 4 - s, vp.hh - 4);
    ctx.fill();
  }
  // frozen pane dividers
  if (vp.frozen && !rs.print) {
    ctx.fillStyle = pal.frozenLine;
    if (vp.colPanes.length > 1) ctx.fillRect(snap(vp.colPanes[1].start) - 1 / dpr, 0, 1 / dpr, vp.height);
    if (vp.rowPanes.length > 1) ctx.fillRect(0, snap(vp.rowPanes[1].start) - 1 / dpr, vp.width, 1 / dpr);
  }
  if (vp.split) {
    ctx.fillStyle = pal.splitBar;
    if (vp.colPanes.length > 1) ctx.fillRect(vp.colPanes[0].start + vp.colPanes[0].size, 0, SPLIT_BAR, vp.height);
    if (vp.rowPanes.length > 1) ctx.fillRect(0, vp.rowPanes[0].start + vp.rowPanes[0].size, vp.width, SPLIT_BAR);
  }
  void isErrorVal;
}

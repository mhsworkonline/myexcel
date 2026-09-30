// Canvas text measurement shared by the grid renderer and autofit.
import { CellStyle, DEFAULT_FONT_SIZE, docDefaultFont } from '../model/styles';

let ctx: CanvasRenderingContext2D | null = null;
function getCtx(): CanvasRenderingContext2D | null {
  if (ctx) return ctx;
  if (typeof document === 'undefined') return null;
  const c = document.createElement('canvas');
  ctx = c.getContext('2d');
  return ctx;
}

export function fontFamily(name: string | undefined): string {
  const n = name || docDefaultFont();
  const fallbacks = '"Carlito", "Calibri", "Segoe UI", Arial, sans-serif';
  return `"${n}", ${fallbacks}`;
}

/** Font size in px for a point size at a given zoom. */
export function ptToPx(pt: number, zoom = 1): number {
  return (pt * 96) / 72 * zoom;
}

export function cssFont(style: CellStyle, zoom = 1): string {
  const size = ptToPx(style.fontSize ?? DEFAULT_FONT_SIZE, zoom) * (style.vertAlign ? 0.7 : 1);
  return `${style.italic ? 'italic ' : ''}${style.bold ? 'bold ' : ''}${size.toFixed(2)}px ${fontFamily(style.fontName)}`;
}

const widthCache = new Map<string, number>();

export function measureText(text: string, style: CellStyle): number {
  const font = cssFont(style);
  const key = font + '|' + text;
  const hit = widthCache.get(key);
  if (hit !== undefined) return hit;
  const c = getCtx();
  let w: number;
  if (!c) w = text.length * 7;
  else {
    c.font = font;
    w = c.measureText(text).width;
  }
  if (widthCache.size > 20000) widthCache.clear();
  widthCache.set(key, w);
  return w;
}

/**
 * Excel-style word wrap shared by drawing and row auto-fit: breaks at spaces and after hyphens,
 * keeps explicit line breaks, and splits a word that is wider than the cell on its own.
 */
export function wrapLines(text: string, width: number, measure: (s: string) => number): string[] {
  const out: string[] = [];
  for (const para of text.split('\n')) {
    const words = para.split(/(\s+)|(?<=-)/).filter(Boolean);
    let line = '';
    for (const w of words) {
      const cand = line + w;
      if (measure(cand.trimEnd()) <= width) {
        line = cand;
        continue;
      }
      if (line.trim()) {
        out.push(line.trimEnd());
        line = w.trimStart();
      } else line = cand;
      // a single word wider than the cell is split where it overflows
      while (line.length > 1 && measure(line) > width) {
        let k = line.length - 1;
        while (k > 1 && measure(line.slice(0, k)) > width) k--;
        out.push(line.slice(0, k));
        line = line.slice(k);
      }
    }
    out.push(line.trimEnd());
  }
  return out;
}

/** Word-wrap text into lines that fit `width` px (100% zoom). */
export function wrapText(text: string, style: CellStyle, width: number): string[] {
  return wrapLines(text, width, (s) => measureText(s, style));
}

/** Row height one line of text needs, px at 100%, close to Excel's (Calibri 11 → 20, 14 → 25). */
export function rowLineHeight(style: CellStyle): number {
  return Math.round(ptToPx(style.fontSize ?? DEFAULT_FONT_SIZE) * 1.33 + 0.5);
}

export function lineHeight(style: CellStyle): number {
  return Math.round(ptToPx(style.fontSize ?? DEFAULT_FONT_SIZE) * 1.2);
}

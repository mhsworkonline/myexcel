// Canvas text measurement shared by the grid renderer and autofit.
import { CellStyle, DEFAULT_FONT, DEFAULT_FONT_SIZE } from '../model/styles';

let ctx: CanvasRenderingContext2D | null = null;
function getCtx(): CanvasRenderingContext2D | null {
  if (ctx) return ctx;
  if (typeof document === 'undefined') return null;
  const c = document.createElement('canvas');
  ctx = c.getContext('2d');
  return ctx;
}

export function fontFamily(name: string | undefined): string {
  const n = name || DEFAULT_FONT;
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

/** Word-wrap text into lines that fit `width` px (100% zoom). */
export function wrapText(text: string, style: CellStyle, width: number): string[] {
  const out: string[] = [];
  for (const para of text.split('\n')) {
    const words = para.split(/(\s+)/);
    let line = '';
    for (const w of words) {
      const cand = line + w;
      if (line && measureText(cand.trimEnd(), style) > width) {
        out.push(line.trimEnd());
        line = w.trimStart();
        // break very long words
        while (measureText(line, style) > width && line.length > 1) {
          let k = line.length - 1;
          while (k > 1 && measureText(line.slice(0, k), style) > width) k--;
          out.push(line.slice(0, k));
          line = line.slice(k);
        }
      } else line = cand;
    }
    out.push(line);
  }
  return out;
}

export function lineHeight(style: CellStyle): number {
  return Math.round(ptToPx(style.fontSize ?? DEFAULT_FONT_SIZE) * 1.2);
}

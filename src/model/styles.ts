// Flat, deduplicated cell styles referenced by numeric id.

export type BorderStyleName =
  | 'thin'
  | 'medium'
  | 'thick'
  | 'dashed'
  | 'dotted'
  | 'double'
  | 'hair'
  | 'mediumDashed'
  | 'dashDot'
  | 'mediumDashDot'
  | 'dashDotDot'
  | 'mediumDashDotDot'
  | 'slantDashDot';

export interface BorderEdge {
  style: BorderStyleName;
  color?: string;
}

export type HAlign = 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
export type VAlign = 'top' | 'middle' | 'bottom' | 'justify' | 'distributed';
export type Underline = 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';

export interface CellStyle {
  fontName?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: Underline;
  strike?: boolean;
  fontColor?: string;
  vertAlign?: 'superscript' | 'subscript';
  fillColor?: string;
  patternType?: string;
  patternColor?: string;
  bTop?: BorderEdge;
  bRight?: BorderEdge;
  bBottom?: BorderEdge;
  bLeft?: BorderEdge;
  bDiagUp?: BorderEdge;
  bDiagDown?: BorderEdge;
  hAlign?: HAlign;
  vAlign?: VAlign;
  wrap?: boolean;
  shrink?: boolean;
  indent?: number;
  rotation?: number;
  numFmt?: string;
  locked?: boolean; // default true
  hidden?: boolean; // formula hidden when protected
}

export const DEFAULT_FONT = 'Calibri';
export const DEFAULT_FONT_SIZE = 11;

export type StylePatch = { [K in keyof CellStyle]?: CellStyle[K] | null };

function canonicalKey(s: CellStyle): string {
  const keys = Object.keys(s).sort() as (keyof CellStyle)[];
  const o: Record<string, unknown> = {};
  for (const k of keys) {
    const v = s[k];
    if (v === undefined || v === null) continue;
    o[k] = v;
  }
  return JSON.stringify(o);
}

export function cleanStyle(s: CellStyle): CellStyle {
  const o: CellStyle = {};
  for (const k of Object.keys(s) as (keyof CellStyle)[]) {
    const v = s[k];
    if (v === undefined || v === null || v === false && k !== 'locked') continue;
    if (k === 'locked' && v === true) continue;
    if (k === 'numFmt' && v === 'General') continue;
    if (k === 'hAlign' && v === 'general') continue;
    if (k === 'vAlign' && v === 'bottom') continue;
    if (k === 'indent' && v === 0) continue;
    if (k === 'rotation' && v === 0) continue;
    if (k === 'fontName' && v === DEFAULT_FONT) continue;
    if (k === 'fontSize' && v === DEFAULT_FONT_SIZE) continue;
    (o as Record<string, unknown>)[k] = v;
  }
  return o;
}

export class StyleTable {
  styles: CellStyle[] = [{}];
  private index = new Map<string, number>([['{}', 0]]);

  get(id: number | undefined): CellStyle {
    return (id !== undefined && this.styles[id]) || this.styles[0];
  }

  intern(s: CellStyle): number {
    const clean = cleanStyle(s);
    const key = canonicalKey(clean);
    const hit = this.index.get(key);
    if (hit !== undefined) return hit;
    const id = this.styles.length;
    this.styles.push(Object.freeze(clean) as CellStyle);
    this.index.set(key, id);
    return id;
  }

  /** Returns id of (base style + patch). null in patch removes a key. */
  merge(baseId: number | undefined, patch: StylePatch): number {
    const base = { ...this.get(baseId) } as Record<string, unknown>;
    for (const [k, v] of Object.entries(patch)) {
      if (v === null || v === undefined) delete base[k];
      else base[k] = v;
    }
    return this.intern(base as CellStyle);
  }

  toJSON(): CellStyle[] {
    return this.styles;
  }

  static fromJSON(arr: CellStyle[]): StyleTable {
    const t = new StyleTable();
    // Preserve ids exactly.
    t.styles = [{}];
    t.index = new Map([['{}', 0]]);
    for (let i = 1; i < arr.length; i++) {
      const clean = cleanStyle(arr[i]);
      t.styles.push(Object.freeze(clean) as CellStyle);
      const key = canonicalKey(clean);
      if (!t.index.has(key)) t.index.set(key, i);
    }
    return t;
  }
}

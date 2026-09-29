// Lightweight Excel formula tokenizer + reference utilities.
import { MAX_COLS, MAX_ROWS, colToName, nameToCol, Range } from './address';

export type TokType = 'ws' | 'op' | 'lparen' | 'rparen' | 'sep' | 'num' | 'str' | 'bool' | 'err' | 'ref' | 'func' | 'name' | 'other';

export interface FTok {
  type: TokType;
  text: string;
  start: number;
  end: number;
  /** For refs: parsed reference. */
  ref?: RefInfo;
}

export interface RefPart {
  r?: number; // undefined for whole-column
  c?: number; // undefined for whole-row
  absR: boolean;
  absC: boolean;
}

export interface RefInfo {
  sheet?: string;
  a: RefPart;
  b?: RefPart;
}

const SHEET_UNQUOTED = /^([A-Za-z_À-￿][\w.À-￿]*)!/;
const CELL = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![\w(])/;
const COLONLY = /^(\$?)([A-Za-z]{1,3})(?![\w(])/;
const ROWONLY = /^(\$?)(\d{1,7})(?![\w.(])/;

function parseCellPart(s: string): { part: RefPart; len: number } | null {
  const m = CELL.exec(s);
  if (!m) return null;
  const c = nameToCol(m[2]);
  const r = parseInt(m[4], 10) - 1;
  if (c >= MAX_COLS || r >= MAX_ROWS || r < 0) return null;
  return { part: { r, c, absC: !!m[1], absR: !!m[3] }, len: m[0].length };
}

/** Try to parse a reference (optionally sheet-qualified) at the start of s. */
function parseRefAt(s: string): { info: RefInfo; len: number } | null {
  let sheet: string | undefined;
  let off = 0;
  if (s[0] === "'") {
    let i = 1;
    let name = '';
    while (i < s.length) {
      if (s[i] === "'") {
        if (s[i + 1] === "'") {
          name += "'";
          i += 2;
          continue;
        }
        break;
      }
      name += s[i++];
    }
    if (s[i] !== "'" || s[i + 1] !== '!') return null;
    sheet = name;
    off = i + 2;
  } else {
    const m = SHEET_UNQUOTED.exec(s);
    if (m) {
      sheet = m[1];
      off = m[0].length;
    }
  }
  const rest = s.slice(off);
  const a = parseCellPart(rest);
  if (a) {
    let len = off + a.len;
    const after = s.slice(len);
    if (after[0] === ':') {
      const b = parseCellPart(after.slice(1));
      if (b) {
        len += 1 + b.len;
        return { info: { sheet, a: a.part, b: b.part }, len };
      }
    }
    return { info: { sheet, a: a.part }, len };
  }
  // column range A:C
  const ca = COLONLY.exec(rest);
  if (ca && rest[ca[0].length] === ':') {
    const cb = COLONLY.exec(rest.slice(ca[0].length + 1));
    if (cb) {
      return {
        info: {
          sheet,
          a: { c: nameToCol(ca[2]), absC: !!ca[1], absR: false },
          b: { c: nameToCol(cb[2]), absC: !!cb[1], absR: false },
        },
        len: off + ca[0].length + 1 + cb[0].length,
      };
    }
  }
  const ra = ROWONLY.exec(rest);
  if (ra && rest[ra[0].length] === ':') {
    const rb = ROWONLY.exec(rest.slice(ra[0].length + 1));
    if (rb) {
      return {
        info: {
          sheet,
          a: { r: parseInt(ra[2], 10) - 1, absR: !!ra[1], absC: false },
          b: { r: parseInt(rb[2], 10) - 1, absR: !!rb[1], absC: false },
        },
        len: off + ra[0].length + 1 + rb[0].length,
      };
    }
  }
  return null;
}

export function tokenize(f: string): FTok[] {
  const toks: FTok[] = [];
  let i = f.startsWith('=') ? 1 : 0;
  if (i === 1) toks.push({ type: 'op', text: '=', start: 0, end: 1 });
  const prevSignificant = () => {
    for (let k = toks.length - 1; k >= 0; k--) if (toks[k].type !== 'ws') return toks[k];
    return undefined;
  };
  while (i < f.length) {
    const ch = f[i];
    const rest = f.slice(i);
    if (/\s/.test(ch)) {
      let j = i;
      while (j < f.length && /\s/.test(f[j])) j++;
      toks.push({ type: 'ws', text: f.slice(i, j), start: i, end: j });
      i = j;
      continue;
    }
    if (ch === '"') {
      let j = i + 1;
      while (j < f.length) {
        if (f[j] === '"') {
          if (f[j + 1] === '"') {
            j += 2;
            continue;
          }
          break;
        }
        j++;
      }
      toks.push({ type: 'str', text: f.slice(i, j + 1), start: i, end: Math.min(j + 1, f.length) });
      i = j + 1;
      continue;
    }
    if (ch === '#') {
      const m = /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!)/i.exec(rest);
      const len = m ? m[0].length : 1;
      toks.push({ type: 'err', text: f.slice(i, i + len), start: i, end: i + len });
      i += len;
      continue;
    }
    if (ch === '(') {
      toks.push({ type: 'lparen', text: ch, start: i, end: i + 1 });
      i++;
      continue;
    }
    if (ch === ')') {
      toks.push({ type: 'rparen', text: ch, start: i, end: i + 1 });
      i++;
      continue;
    }
    if (ch === ',' || ch === ';') {
      toks.push({ type: 'sep', text: ch, start: i, end: i + 1 });
      i++;
      continue;
    }
    if ('+-*/^&=<>%:'.includes(ch)) {
      const two = f.slice(i, i + 2);
      const len = two === '<=' || two === '>=' || two === '<>' ? 2 : 1;
      toks.push({ type: 'op', text: f.slice(i, i + len), start: i, end: i + len });
      i += len;
      continue;
    }
    if (ch === '{' || ch === '}') {
      toks.push({ type: 'other', text: ch, start: i, end: i + 1 });
      i++;
      continue;
    }
    // Numbers (but 1:3 row ranges are refs)
    const prev = prevSignificant();
    const refAllowed = !prev || prev.type === 'op' || prev.type === 'lparen' || prev.type === 'sep';
    if (refAllowed || ch === "'") {
      const ref = parseRefAt(rest);
      if (ref) {
        toks.push({ type: 'ref', text: f.slice(i, i + ref.len), start: i, end: i + ref.len, ref: ref.info });
        i += ref.len;
        continue;
      }
    }
    const num = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(rest);
    if (num) {
      toks.push({ type: 'num', text: num[0], start: i, end: i + num[0].length });
      i += num[0].length;
      continue;
    }
    const id = /^[A-Za-z_\\À-￿][\w.À-￿]*(\[[^\]]*(\[[^\]]*\][^\]]*)*\])?/.exec(rest);
    if (id) {
      const text = id[0];
      let j = i + text.length;
      let k = j;
      while (k < f.length && f[k] === ' ') k++;
      const up = text.toUpperCase();
      if (f[k] === '(' && !text.includes('[')) toks.push({ type: 'func', text, start: i, end: j });
      else if (up === 'TRUE' || up === 'FALSE') toks.push({ type: 'bool', text, start: i, end: j });
      else toks.push({ type: 'name', text, start: i, end: j });
      i = j;
      continue;
    }
    toks.push({ type: 'other', text: ch, start: i, end: i + 1 });
    i++;
  }
  return toks;
}

export function refPartToString(p: RefPart): string {
  let s = '';
  if (p.c !== undefined) s += (p.absC ? '$' : '') + colToName(p.c);
  if (p.r !== undefined) s += (p.absR ? '$' : '') + (p.r + 1);
  return s;
}

function quoteSheet(name: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}

export function refToString(ref: RefInfo): string {
  const pre = ref.sheet ? quoteSheet(ref.sheet) + '!' : '';
  return pre + refPartToString(ref.a) + (ref.b ? ':' + refPartToString(ref.b) : '');
}

export function refToRange(ref: RefInfo): Range {
  const a = ref.a;
  const b = ref.b ?? ref.a;
  const r1 = a.r ?? 0;
  const r2 = b.r ?? MAX_ROWS - 1;
  const c1 = a.c ?? 0;
  const c2 = b.c ?? MAX_COLS - 1;
  return { r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: Math.min(c1, c2), c2: Math.max(c1, c2) };
}

/** Shift relative parts of every reference by (dr, dc). Out-of-bounds references become #REF!. */
export function shiftFormula(f: string, dr: number, dc: number): string {
  if (!f.startsWith('=') || (dr === 0 && dc === 0)) return f;
  const toks = tokenize(f);
  let out = '';
  for (const t of toks) {
    if (t.type !== 'ref' || !t.ref) {
      out += t.text;
      continue;
    }
    const shiftPart = (p: RefPart): RefPart | null => {
      const n: RefPart = { ...p };
      if (n.r !== undefined && !n.absR) n.r += dr;
      if (n.c !== undefined && !n.absC) n.c += dc;
      if ((n.r !== undefined && (n.r < 0 || n.r >= MAX_ROWS)) || (n.c !== undefined && (n.c < 0 || n.c >= MAX_COLS))) return null;
      return n;
    };
    const a = shiftPart(t.ref.a);
    const b = t.ref.b ? shiftPart(t.ref.b) : undefined;
    if (!a || b === null) {
      out += '#REF!';
      continue;
    }
    out += refToString({ sheet: t.ref.sheet, a, b });
  }
  return out;
}

/** Cycle absolute/relative on the reference at or before caret (F4). */
export function toggleRefAt(f: string, caret: number): { text: string; caret: number } {
  const toks = tokenize(f);
  const t = toks.find((x) => x.type === 'ref' && caret >= x.start && caret <= x.end);
  if (!t || !t.ref) return { text: f, caret };
  const cycle = (p: RefPart): RefPart => {
    // A1 → $A$1 → A$1 → $A1 → A1
    if (!p.absR && !p.absC) return { ...p, absR: true, absC: true };
    if (p.absR && p.absC) return { ...p, absR: true, absC: false };
    if (p.absR && !p.absC) return { ...p, absR: false, absC: true };
    return { ...p, absR: false, absC: false };
  };
  const nref: RefInfo = { sheet: t.ref.sheet, a: cycle(t.ref.a), b: t.ref.b ? cycle(t.ref.b) : undefined };
  const s = refToString(nref);
  const text = f.slice(0, t.start) + s + f.slice(t.end);
  return { text, caret: t.start + s.length };
}

export interface FormulaRef {
  sheet?: string;
  range: Range;
  start: number;
  end: number;
  colorIndex: number;
}

export const REF_COLORS = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E', '#636363', '#997300'];
export const REF_COLORS_EXCEL = ['#0000FF', '#008000', '#9900CC', '#800000', '#00CC33', '#CC6600', '#CC0099'];

export function extractRefs(f: string): FormulaRef[] {
  if (!f.startsWith('=')) return [];
  const toks = tokenize(f);
  const out: FormulaRef[] = [];
  const colorByText = new Map<string, number>();
  for (const t of toks) {
    if (t.type !== 'ref' || !t.ref) continue;
    const key = t.text.replace(/\$/g, '').toUpperCase();
    let ci = colorByText.get(key);
    if (ci === undefined) {
      ci = colorByText.size % REF_COLORS_EXCEL.length;
      colorByText.set(key, ci);
    }
    out.push({ sheet: t.ref.sheet, range: refToRange(t.ref), start: t.start, end: t.end, colorIndex: ci });
  }
  return out;
}

/** Whether inserting a reference at caret makes sense (after operator, '(' or ','). */
export function canInsertRefAt(f: string, caret: number): boolean {
  if (!f.startsWith('=')) return false;
  const before = f.slice(0, caret).trimEnd();
  if (before.length === 0) return false;
  const last = before[before.length - 1];
  return '=+-*/^&(,<>:;%'.includes(last);
}

/** Returns the function call context at the caret: name and argument index. */
export function functionContextAt(f: string, caret: number): { name: string; argIndex: number } | null {
  const toks = tokenize(f.slice(0, caret));
  const stack: { name: string; arg: number }[] = [];
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.type === 'lparen') {
      let j = i - 1;
      while (j >= 0 && toks[j].type === 'ws') j--;
      stack.push({ name: j >= 0 && toks[j].type === 'func' ? toks[j].text.toUpperCase() : '', arg: 0 });
    } else if (t.type === 'rparen') stack.pop();
    else if (t.type === 'sep' && stack.length) stack[stack.length - 1].arg++;
  }
  for (let k = stack.length - 1; k >= 0; k--) if (stack[k].name) return { name: stack[k].name, argIndex: stack[k].arg };
  return null;
}

/** The identifier being typed at the caret (for function autocomplete). */
export function partialIdentAt(f: string, caret: number): { text: string; start: number } | null {
  if (!f.startsWith('=')) return null;
  const before = f.slice(0, caret);
  const m = /([A-Za-z_][A-Za-z0-9_.]*)$/.exec(before);
  if (!m) return null;
  const start = caret - m[1].length;
  const prev = before.slice(0, start).trimEnd();
  const last = prev[prev.length - 1];
  if (last && !'=+-*/^&(,<>;: '.includes(last)) return null;
  return { text: m[1], start };
}

/** Rename sheet references inside a formula. */
export function renameSheetInFormula(f: string, oldName: string, newName: string): string {
  if (!f.startsWith('=')) return f;
  const toks = tokenize(f);
  let changed = false;
  const out = toks
    .map((t) => {
      if (t.type === 'ref' && t.ref?.sheet && t.ref.sheet.toLowerCase() === oldName.toLowerCase()) {
        changed = true;
        return refToString({ ...t.ref, sheet: newName });
      }
      return t.text;
    })
    .join('');
  return changed ? out : f;
}

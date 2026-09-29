// Reference adjustment for row/column insertion and deletion.
import type { Range } from './address';
import { RefInfo, RefPart, refToString, tokenize } from './formula';

export type Axis = 'row' | 'col';

export interface StructOp {
  axis: Axis;
  at: number;
  /** positive = insert count, negative = delete count */
  delta: number;
}

function adjIndex(i: number, op: StructOp): number | null {
  if (op.delta > 0) return i >= op.at ? i + op.delta : i;
  const n = -op.delta;
  if (i < op.at) return i;
  if (i < op.at + n) return null;
  return i - n;
}

/** Adjust a [a,b] span. Returns null when fully deleted. */
export function adjSpan(a: number, b: number, op: StructOp): [number, number] | null {
  if (op.delta > 0) {
    return [a >= op.at ? a + op.delta : a, b >= op.at ? b + op.delta : b];
  }
  const n = -op.delta;
  const end = op.at + n - 1;
  if (a >= op.at && b <= end) return null;
  const na = a < op.at ? a : a <= end ? op.at : a - n;
  const nb = b < op.at ? b : b <= end ? op.at - 1 : b - n;
  if (nb < na) return null;
  return [na, nb];
}

export function adjustRange(rg: Range, op: StructOp): Range | null {
  if (op.axis === 'row') {
    const s = adjSpan(rg.r1, rg.r2, op);
    return s ? { ...rg, r1: s[0], r2: s[1] } : null;
  }
  const s = adjSpan(rg.c1, rg.c2, op);
  return s ? { ...rg, c1: s[0], c2: s[1] } : null;
}

function adjustRef(ref: RefInfo, op: StructOp): RefInfo | null {
  const key = op.axis === 'row' ? 'r' : 'c';
  const a = ref.a;
  const b = ref.b;
  if (a[key] === undefined) return ref; // whole row/col ref unaffected on this axis
  if (!b) {
    const n = adjIndex(a[key]!, op);
    if (n === null) return null;
    return { ...ref, a: { ...a, [key]: n } };
  }
  const lo = Math.min(a[key]!, b[key]!);
  const hi = Math.max(a[key]!, b[key]!);
  const s = adjSpan(lo, hi, op);
  if (!s) return null;
  const na: RefPart = { ...a, [key]: a[key]! <= b[key]! ? s[0] : s[1] };
  const nb: RefPart = { ...b, [key]: a[key]! <= b[key]! ? s[1] : s[0] };
  return { ...ref, a: na, b: nb };
}

/**
 * Adjust references in `formula` (living on sheet `hostSheet`) for a structural op on `targetSheet`.
 * Returns the same string instance when nothing changed.
 */
export function adjustFormula(formula: string, hostSheet: string, targetSheet: string, op: StructOp): string {
  if (!formula.startsWith('=')) return formula;
  const toks = tokenize(formula);
  let changed = false;
  let out = '';
  const target = targetSheet.toLowerCase();
  for (const t of toks) {
    if (t.type !== 'ref' || !t.ref) {
      out += t.text;
      continue;
    }
    const refSheet = (t.ref.sheet ?? hostSheet).toLowerCase();
    if (refSheet !== target) {
      out += t.text;
      continue;
    }
    const n = adjustRef(t.ref, op);
    if (!n) {
      out += t.ref.sheet ? `${t.text.slice(0, t.text.indexOf('!') + 1)}#REF!` : '#REF!';
      changed = true;
      continue;
    }
    const s = refToString(n);
    if (s !== t.text) changed = true;
    out += s;
  }
  return changed ? out : formula;
}

/** Adjust a defined-name / chart reference string like "Sheet1!$A$1:$B$5" (no leading '='). */
export function adjustRefString(ref: string, hostSheet: string, targetSheet: string, op: StructOp): string {
  const r = adjustFormula('=' + ref, hostSheet, targetSheet, op);
  return r.slice(1);
}

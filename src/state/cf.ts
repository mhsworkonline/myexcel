// Conditional formatting evaluation with per-revision caches.
import { isErrorVal } from '../engine/Engine';
import { shiftFormula } from '../model/formula';
import { serialToParts } from '../model/numfmt';
import type { Sheet } from '../model/sheet';
import type { CFRule, CFStyle, CFVO } from '../model/types';
import { S } from './store';
import { getComputed } from './values';

export interface CFResult {
  style?: CFStyle;
  fill?: string;
  bar?: { from: number; to: number; color: string; gradient: boolean; negative: boolean; axis?: number; showValue: boolean };
  icon?: { set: string; index: number; count: number; showValue: boolean };
}

interface RuleStats {
  nums: number[]; // sorted ascending
  sum: number;
  counts?: Map<string, number>;
  thresh?: number;
  avg?: number;
  sd?: number;
}

let cacheKey = '';
const statsCache = new Map<string, RuleStats>();
const evalCache = new Map<string, boolean>();

function key(): string {
  const st = S();
  return `${st.wb.rev}:${st.engine.rev}`;
}

function checkCache(): void {
  const k = key();
  if (k !== cacheKey) {
    cacheKey = k;
    statsCache.clear();
    evalCache.clear();
  }
}

function inRule(rule: CFRule, r: number, c: number): boolean {
  for (const rg of rule.ranges) if (r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2) return true;
  return false;
}

function statsFor(sheet: Sheet, rule: CFRule): RuleStats {
  let s = statsCache.get(rule.id);
  if (s) return s;
  const nums: number[] = [];
  const counts = new Map<string, number>();
  const needCounts = rule.type === 'duplicateValues' || rule.type === 'uniqueValues';
  for (const rg of rule.ranges) {
    sheet.forEachInRange(rg, (r, c) => {
      const v = getComputed(sheet, r, c);
      if (typeof v === 'number') nums.push(v);
      if (needCounts && v !== null && v !== '' && !isErrorVal(v)) {
        const k = String(v).toLowerCase();
        counts.set(k, (counts.get(k) ?? 0) + 1);
      }
    });
  }
  nums.sort((a, b) => a - b);
  const sum = nums.reduce((a, b) => a + b, 0);
  s = { nums, sum, counts };
  if (nums.length) {
    s.avg = sum / nums.length;
    s.sd = Math.sqrt(nums.reduce((a, b) => a + (b - s!.avg!) * (b - s!.avg!), 0) / Math.max(1, nums.length - 1));
  }
  if (rule.type === 'top10' && nums.length) {
    const n = rule.percent ? Math.max(1, Math.floor((nums.length * (rule.rank ?? 10)) / 100)) : Math.min(rule.rank ?? 10, nums.length);
    s.thresh = rule.bottom ? nums[n - 1] : nums[nums.length - n];
  }
  statsCache.set(rule.id, s);
  return s;
}

function topLeft(rule: CFRule): { r: number; c: number } {
  let r = Infinity;
  let c = Infinity;
  for (const rg of rule.ranges) {
    r = Math.min(r, rg.r1);
    c = Math.min(c, rg.c1);
  }
  return { r, c };
}

function evalFormula(sheet: Sheet, rule: CFRule, f: string, r: number, c: number): unknown {
  const t = f.trim();
  if (!t.startsWith('=')) {
    const n = Number(t);
    if (!isNaN(n) && t !== '') return n;
    return t.replace(/^"|"$/g, '');
  }
  const tl = topLeft(rule);
  const shifted = shiftFormula(t, r - tl.r, c - tl.c);
  const v = S().engine.evaluate(shifted, sheet);
  if (isErrorVal(v)) return null;
  return v;
}

function percentile(sorted: number[], p: number): number {
  if (!sorted.length) return 0;
  const idx = (sorted.length - 1) * p;
  const lo = Math.floor(idx);
  const hi = Math.ceil(idx);
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (idx - lo);
}

function cfvoValue(sheet: Sheet, rule: CFRule, v: CFVO, st: RuleStats, isMin: boolean): number {
  const lo = st.nums[0] ?? 0;
  const hi = st.nums[st.nums.length - 1] ?? 0;
  switch (v.type) {
    case 'min':
    case 'autoMin':
      return v.type === 'autoMin' ? Math.min(0, lo) : lo;
    case 'max':
    case 'autoMax':
      return v.type === 'autoMax' ? Math.max(0, hi) : hi;
    case 'num':
      return Number(v.value ?? 0);
    case 'percent':
      return lo + ((hi - lo) * Number(v.value ?? 0)) / 100;
    case 'percentile':
      return percentile(st.nums, Number(v.value ?? 0) / 100);
    case 'formula': {
      const x = evalFormula(sheet, rule, String(v.value ?? '0'), topLeft(rule).r, topLeft(rule).c);
      return typeof x === 'number' ? x : isMin ? lo : hi;
    }
  }
  return 0;
}

function hexToRgb(h: string): [number, number, number] {
  const n = parseInt(h.replace('#', ''), 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

function mix(a: string, b: string, t: number): string {
  const x = hexToRgb(a);
  const y = hexToRgb(b);
  return '#' + x.map((v, i) => Math.round(v + (y[i] - v) * t).toString(16).padStart(2, '0')).join('');
}

const ICON_COUNTS: Record<string, number> = {};
export function iconCount(set: string): number {
  if (ICON_COUNTS[set]) return ICON_COUNTS[set];
  const n = parseInt(set, 10) || 3;
  ICON_COUNTS[set] = n;
  return n;
}

function textOf(v: unknown): string {
  if (v === null || v === undefined) return '';
  if (isErrorVal(v)) return v.error;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return String(v);
}

function compare(op: CFRule['operator'], v: unknown, a: unknown, b: unknown): boolean {
  const num = (x: unknown) => (typeof x === 'number' ? x : typeof x === 'boolean' ? (x ? 1 : 0) : NaN);
  let x: number | string = num(v);
  let ya: number | string = num(a);
  let yb: number | string = num(b);
  if (isNaN(x as number) || isNaN(ya as number)) {
    // string comparison (case-insensitive) when both are text
    if (typeof v !== 'string' || typeof a !== 'string') {
      if (op === 'notEqual') return textOf(v).toLowerCase() !== textOf(a).toLowerCase();
      if (op === 'equal') return textOf(v).toLowerCase() === textOf(a).toLowerCase();
      if (v === null || v === '' || typeof v !== 'number') return false;
    }
    x = textOf(v).toLowerCase();
    ya = textOf(a).toLowerCase();
    yb = textOf(b).toLowerCase();
  }
  switch (op) {
    case 'between': return x >= (ya < yb ? ya : yb) && x <= (ya < yb ? yb : ya);
    case 'notBetween': return !(x >= (ya < yb ? ya : yb) && x <= (ya < yb ? yb : ya));
    case 'equal': return x === ya;
    case 'notEqual': return x !== ya;
    case 'greaterThan': return x > ya;
    case 'lessThan': return x < ya;
    case 'greaterThanOrEqual': return x >= ya;
    case 'lessThanOrEqual': return x <= ya;
  }
  return false;
}

function inPeriod(serial: number, period: CFRule['timePeriod']): boolean {
  const now = new Date();
  const today = Math.floor((Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()) - Date.UTC(1899, 11, 30)) / 86400000);
  const d = Math.floor(serial);
  const tp = serialToParts(today);
  const dp = serialToParts(d);
  const weekStart = today - tp.dow;
  switch (period) {
    case 'today': return d === today;
    case 'yesterday': return d === today - 1;
    case 'tomorrow': return d === today + 1;
    case 'last7Days': return d > today - 7 && d <= today;
    case 'thisWeek': return d >= weekStart && d < weekStart + 7;
    case 'lastWeek': return d >= weekStart - 7 && d < weekStart;
    case 'nextWeek': return d >= weekStart + 7 && d < weekStart + 14;
    case 'thisMonth': return dp.y === tp.y && dp.m === tp.m;
    case 'lastMonth': return (tp.m === 1 ? dp.y === tp.y - 1 && dp.m === 12 : dp.y === tp.y && dp.m === tp.m - 1);
    case 'nextMonth': return (tp.m === 12 ? dp.y === tp.y + 1 && dp.m === 1 : dp.y === tp.y && dp.m === tp.m + 1);
  }
  return false;
}

function ruleMatches(sheet: Sheet, rule: CFRule, r: number, c: number, v: unknown): boolean {
  const ck = `${rule.id}:${r}:${c}`;
  const hit = evalCache.get(ck);
  if (hit !== undefined) return hit;
  let res = false;
  switch (rule.type) {
    case 'cellIs': {
      if (v === null || v === '') {
        res = false;
        break;
      }
      const a = evalFormula(sheet, rule, rule.formulas?.[0] ?? '', r, c);
      const b = rule.formulas?.[1] !== undefined ? evalFormula(sheet, rule, rule.formulas[1], r, c) : undefined;
      res = compare(rule.operator ?? 'equal', v, a, b);
      break;
    }
    case 'expression': {
      const x = evalFormula(sheet, rule, rule.formulas?.[0] ?? '', r, c);
      res = x === true || (typeof x === 'number' && x !== 0);
      break;
    }
    case 'containsText':
      res = textOf(v).toLowerCase().includes((rule.text ?? '').toLowerCase());
      break;
    case 'notContainsText':
      res = !textOf(v).toLowerCase().includes((rule.text ?? '').toLowerCase());
      break;
    case 'beginsWith':
      res = textOf(v).toLowerCase().startsWith((rule.text ?? '').toLowerCase());
      break;
    case 'endsWith':
      res = textOf(v).toLowerCase().endsWith((rule.text ?? '').toLowerCase());
      break;
    case 'containsBlanks':
      res = v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
      break;
    case 'notContainsBlanks':
      res = !(v === null || v === undefined || (typeof v === 'string' && v.trim() === ''));
      break;
    case 'containsErrors':
      res = isErrorVal(v);
      break;
    case 'notContainsErrors':
      res = !isErrorVal(v);
      break;
    case 'timePeriod':
      res = typeof v === 'number' && inPeriod(v, rule.timePeriod);
      break;
    case 'top10': {
      const st = statsFor(sheet, rule);
      res = typeof v === 'number' && st.thresh !== undefined && (rule.bottom ? v <= st.thresh : v >= st.thresh);
      break;
    }
    case 'aboveAverage': {
      const st = statsFor(sheet, rule);
      if (typeof v !== 'number' || st.avg === undefined) break;
      const above = rule.aboveAverage !== false;
      const lim = st.avg + (rule.stdDev ? (above ? 1 : -1) * rule.stdDev * (st.sd ?? 0) : 0);
      res = above ? (rule.equalAverage ? v >= lim : v > lim) : rule.equalAverage ? v <= lim : v < lim;
      break;
    }
    case 'duplicateValues':
    case 'uniqueValues': {
      if (v === null || v === '' || isErrorVal(v)) break;
      const st = statsFor(sheet, rule);
      const n = st.counts?.get(String(v).toLowerCase()) ?? 0;
      res = rule.type === 'duplicateValues' ? n > 1 : n === 1;
      break;
    }
    default:
      res = false;
  }
  evalCache.set(ck, res);
  return res;
}

/** Conditional formatting outcome for a cell, or null. */
export function cfAt(sheet: Sheet, r: number, c: number): CFResult | null {
  const rules = sheet.conditionalFormats;
  if (!rules.length) return null;
  let applicable: CFRule[] | null = null;
  for (const rule of rules) {
    if (inRule(rule, r, c)) (applicable ??= []).push(rule);
  }
  if (!applicable) return null;
  checkCache();
  applicable.sort((a, b) => a.priority - b.priority);
  const v = getComputed(sheet, r, c);
  const out: CFResult = {};
  let style: CFStyle | undefined;
  for (const rule of applicable) {
    if (rule.type === 'dataBar' && rule.dataBar) {
      if (out.bar || typeof v !== 'number') continue;
      const st = statsFor(sheet, rule);
      const lo = cfvoValue(sheet, rule, rule.dataBar.min, st, true);
      const hi = cfvoValue(sheet, rule, rule.dataBar.max, st, false);
      if (hi === lo) {
        out.bar = { from: 0, to: 0.5, color: rule.dataBar.color, gradient: rule.dataBar.gradient, negative: false, showValue: rule.dataBar.showValue };
        continue;
      }
      if (lo < 0 && hi > 0) {
        const axis = -lo / (hi - lo);
        const pos = (Math.max(lo, Math.min(hi, v)) - lo) / (hi - lo);
        out.bar = { from: Math.min(axis, pos), to: Math.max(axis, pos), color: v < 0 ? rule.dataBar.negColor ?? '#FF0000' : rule.dataBar.color, gradient: rule.dataBar.gradient, negative: v < 0, axis, showValue: rule.dataBar.showValue };
      } else {
        const frac = Math.max(0, Math.min(1, (v - lo) / (hi - lo)));
        out.bar = { from: 0, to: 0.1 + 0.9 * frac, color: rule.dataBar.color, gradient: rule.dataBar.gradient, negative: false, showValue: rule.dataBar.showValue };
      }
      continue;
    }
    if (rule.type === 'colorScale' && rule.colorScale) {
      if (out.fill || typeof v !== 'number') continue;
      const st = statsFor(sheet, rule);
      const cs = rule.colorScale;
      const pts = cs.cfvos.map((cv, i) => cfvoValue(sheet, rule, cv, st, i === 0));
      if (pts.length === 2 || cs.colors.length === 2) {
        const t = pts[1] === pts[0] ? 0 : (v - pts[0]) / (pts[pts.length - 1] - pts[0]);
        out.fill = mix(cs.colors[0], cs.colors[cs.colors.length - 1], Math.max(0, Math.min(1, t)));
      } else {
        if (v <= pts[1]) {
          const t = pts[1] === pts[0] ? 0 : (v - pts[0]) / (pts[1] - pts[0]);
          out.fill = mix(cs.colors[0], cs.colors[1], Math.max(0, Math.min(1, t)));
        } else {
          const t = pts[2] === pts[1] ? 1 : (v - pts[1]) / (pts[2] - pts[1]);
          out.fill = mix(cs.colors[1], cs.colors[2], Math.max(0, Math.min(1, t)));
        }
      }
      continue;
    }
    if (rule.type === 'iconSet' && rule.iconSet) {
      if (out.icon || typeof v !== 'number') continue;
      const st = statsFor(sheet, rule);
      const n = iconCount(rule.iconSet.set);
      const cfvos = rule.iconSet.cfvos.length === n ? rule.iconSet.cfvos : Array.from({ length: n }, (_, i) => ({ type: 'percent' as const, value: Math.round((i * 100) / n) }));
      let idx = 0;
      for (let i = 1; i < n; i++) {
        const t = cfvoValue(sheet, rule, cfvos[i], st, false);
        if ((cfvos[i] as CFVO).gte === false ? v > t : v >= t) idx = i;
      }
      if (rule.iconSet.reverse) idx = n - 1 - idx;
      out.icon = { set: rule.iconSet.set, index: idx, count: n, showValue: rule.iconSet.showValue !== false };
      continue;
    }
    if (ruleMatches(sheet, rule, r, c, v)) {
      if (rule.style) style = { ...rule.style, ...(style ?? {}) };
      if (rule.stopIfTrue) break;
    }
  }
  if (style) out.style = style;
  return out.style || out.fill || out.bar || out.icon ? out : null;
}

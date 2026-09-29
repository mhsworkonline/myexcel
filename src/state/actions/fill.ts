import { Range } from '../../model/address';
import { shiftFormula } from '../../model/formula';
import { isDateFormat, partsToSerial, serialToParts } from '../../model/numfmt';
import { primaryRange, rangeSel } from '../../model/selection';
import type { Sheet } from '../../model/sheet';
import type { Cell } from '../../model/types';
import { alertBox, S, setStatus, transact } from '../store';
import { guardRanges, writeInput } from './edit';

export type FillMode = 'auto' | 'copy' | 'series' | 'formats' | 'noFormats' | 'days' | 'weekdays' | 'months' | 'years';

const LISTS = [
  ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
  ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  ['Q1', 'Q2', 'Q3', 'Q4'],
];

function findList(s: string): { list: string[]; idx: number; caseFn: (x: string) => string } | null {
  for (const list of LISTS) {
    const idx = list.findIndex((x) => x.toLowerCase() === s.toLowerCase());
    if (idx >= 0) {
      const caseFn = s === s.toUpperCase() ? (x: string) => x.toUpperCase() : s === s.toLowerCase() ? (x: string) => x.toLowerCase() : (x: string) => x;
      return { list, idx, caseFn };
    }
  }
  return null;
}

interface Src {
  cell: Cell | undefined;
  r: number;
  c: number;
}

function addMonths(serial: number, n: number): number {
  const p = serialToParts(serial);
  const m0 = p.m - 1 + n;
  const y = p.y + Math.floor(m0 / 12);
  const m = ((m0 % 12) + 12) % 12;
  const dim = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
  return partsToSerial(y, m + 1, Math.min(p.d, dim)) + (serial - Math.floor(serial));
}

function addWeekdays(serial: number, n: number): number {
  let s = serial;
  let k = Math.abs(n);
  const d = n >= 0 ? 1 : -1;
  while (k > 0) {
    s += d;
    const dow = serialToParts(s).dow;
    if (dow !== 0 && dow !== 6) k--;
  }
  return s;
}

/** Produce the value for position k (0-based beyond the source end) in one fill line. */
function seriesGen(src: Src[], sheet: Sheet, mode: FillMode): (k: number) => { cell: Cell | undefined; from: Src } {
  const n = src.length;
  const wb = S().wb;
  const vals = src.map((s) => s.cell?.v);
  const allNums = src.every((s) => typeof s.cell?.v === 'number' && !s.cell?.f);
  const isDate = allNums && isDateFormat(wb.styles.get(src[0].cell?.s ?? sheet.styleIdAt(src[0].r, src[0].c)).numFmt);
  const copyMode = mode === 'copy' || mode === 'formats';
  if (!copyMode && allNums && n > 0) {
    const nums = vals as number[];
    if (isDate && mode !== 'series') {
      let unit: 'd' | 'm' | 'y' | 'w' = mode === 'months' ? 'm' : mode === 'years' ? 'y' : mode === 'weekdays' ? 'w' : 'd';
      let step = 1;
      if (n >= 2 && mode === 'auto') {
        const a = serialToParts(nums[0]);
        const b = serialToParts(nums[1]);
        if (a.d === b.d && (b.m !== a.m || b.y !== a.y)) {
          unit = 'm';
          step = (b.y - a.y) * 12 + (b.m - a.m);
        } else step = nums[1] - nums[0];
      }
      return (k) => {
        const base = nums[n - 1];
        const t = k + 1;
        let v: number;
        if (unit === 'm') v = addMonths(base, step * t);
        else if (unit === 'y') v = addMonths(base, 12 * t);
        else if (unit === 'w') v = addWeekdays(base, t);
        else v = base + step * t;
        return { cell: { ...src[n - 1].cell, v }, from: src[n - 1] };
      };
    }
    if (n === 1 && mode !== 'series') return () => ({ cell: src[0].cell, from: src[0] });
    // linear trend (least squares)
    let step: number;
    let intercept: number;
    if (n === 1) {
      step = 1;
      intercept = nums[0];
    } else {
      const xm = (n - 1) / 2;
      const ym = nums.reduce((a, b) => a + b, 0) / n;
      let num = 0;
      let den = 0;
      nums.forEach((y, x) => {
        num += (x - xm) * (y - ym);
        den += (x - xm) * (x - xm);
      });
      step = den ? num / den : 0;
      intercept = ym - step * xm;
    }
    return (k) => {
      const x = n + k;
      const v = +(intercept + step * x).toPrecision(15);
      return { cell: { ...src[(n - 1 + k + 1) % n].cell, v }, from: src[(n + k) % n] };
    };
  }
  if (!copyMode) {
    // Uniform text sequences: list items (Jan, Feb / Mon, Tue / Q1) or "Item 1", "Item 2"
    const strs = src.map((x) => (x.cell && !x.cell.f && typeof x.cell.v === 'string' ? x.cell.v : null));
    if (n > 0 && strs.every((x) => x !== null)) {
      const lists = strs.map((x) => findList(x!));
      const l0 = lists[0];
      if (l0 && lists.every((l) => l && l.list === l0.list)) {
        const L = l0.list.length;
        const idxs = lists.map((l) => l!.idx);
        let step = 1;
        if (n > 1) {
          const d = (idxs[1] - idxs[0] + L) % L;
          step = idxs.every((v, i) => i === 0 || (v - idxs[i - 1] + L) % L === d) ? d : 1;
        }
        return (k) => {
          const j = (((idxs[0] + step * (n + k)) % L) + L) % L;
          const from = src[(n + k) % n];
          return { cell: { ...from.cell, v: l0.caseFn(l0.list[j]) }, from };
        };
      }
      const tn = strs.map((x) => /^(.*?)(d+)(D*)$/.exec(x!));
      const t0 = tn[0];
      if (t0 && tn.every((m) => m && m[1] === t0[1] && m[3] === t0[3])) {
        const nums = tn.map((m) => parseInt(m![2], 10));
        let step = 1;
        if (n > 1) {
          const d = nums[1] - nums[0];
          step = nums.every((v, i) => i === 0 || v - nums[i - 1] === d) ? d : 1;
        }
        return (k) => {
          const num = nums[0] + step * (n + k);
          const digits = t0[2].startsWith('0') ? String(Math.abs(num)).padStart(t0[2].length, '0') : String(Math.abs(num));
          const from = src[(n + k) % n];
          return { cell: { ...from.cell, v: t0[1] + (num < 0 ? '-' : '') + digits + t0[3] }, from };
        };
      }
    }
  }
  // Mixed pattern: repeat each source cell; list items / trailing numbers advance per cycle
  return (k) => {
    const idx = k % n;
    const cycle = Math.floor(k / n) + 1;
    const from = src[idx];
    const cell = from.cell;
    if (!cell || copyMode || cell.f || typeof cell.v !== 'string') return { cell, from };
    const lst = findList(cell.v);
    if (lst) return { cell: { ...cell, v: lst.caseFn(lst.list[(lst.idx + cycle) % lst.list.length]) }, from };
    const m = /^(.*?)(d+)(D*)$/.exec(cell.v);
    if (m) {
      const num = parseInt(m[2], 10) + cycle;
      const digits = m[2].startsWith('0') ? String(num).padStart(m[2].length, '0') : String(num);
      return { cell: { ...cell, v: m[1] + digits + m[3] }, from };
    }
    return { cell, from };
  };
}

/** Fill from `src` into `target` (target includes src). Direction inferred. */
export function autoFill(src: Range, target: Range, mode: FillMode = 'auto', ctrl = false): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  if (!guardRanges(sheet, [target])) return;
  const down = target.r2 > src.r2;
  const up = target.r1 < src.r1;
  const right = target.c2 > src.c2;
  const left = target.c1 < src.c1;
  if (!(down || up || right || left)) {
    // shrinking the selection with the fill handle clears cells
    if (target.r2 < src.r2 || target.c2 < src.c2) {
      transact('Clear', (tx) => {
        const clr: Range = target.r2 < src.r2 ? { r1: target.r2 + 1, r2: src.r2, c1: src.c1, c2: src.c2 } : { r1: src.r1, r2: src.r2, c1: target.c2 + 1, c2: src.c2 };
        sheet.forEachInRange(clr, (r, c, cell) => {
          const n = { ...cell };
          delete n.v;
          delete n.f;
          delete n.e;
          tx.setCell(sheet, r, c, n);
        });
      }, rangeSel(target));
    }
    return;
  }
  const vertical = down || up;
  transact('Auto Fill', (tx) => {
    const lines = vertical ? src.c2 - src.c1 + 1 : src.r2 - src.r1 + 1;
    for (let li = 0; li < lines; li++) {
      const srcCells: Src[] = [];
      if (vertical) for (let r = src.r1; r <= src.r2; r++) srcCells.push({ cell: sheet.getCell(r, src.c1 + li), r, c: src.c1 + li });
      else for (let c = src.c1; c <= src.c2; c++) srcCells.push({ cell: sheet.getCell(src.r1 + li, c), r: src.r1 + li, c });
      const reverse = up || left;
      const ordered = reverse ? [...srcCells].reverse() : srcCells;
      const single = ordered.length === 1 && typeof ordered[0].cell?.v === 'number' && !ordered[0].cell?.f;
      const eff: FillMode = ctrl && mode === 'auto' ? (single ? 'series' : 'copy') : mode;
      const gen = seriesGen(ordered, sheet, eff);
      const count = vertical ? (down ? target.r2 - src.r2 : src.r1 - target.r1) : right ? target.c2 - src.c2 : src.c1 - target.c1;
      for (let k = 0; k < count; k++) {
        const r = vertical ? (down ? src.r2 + 1 + k : src.r1 - 1 - k) : src.r1 + li;
        const c = vertical ? src.c1 + li : right ? src.c2 + 1 + k : src.c1 - 1 - k;
        const { cell: genCell, from } = gen(k);
        let next: Cell | undefined = genCell ? { ...genCell } : undefined;
        if (next?.f) next.f = shiftFormula(next.f, r - from.r, c - from.c);
        const cur = sheet.getCell(r, c);
        if (mode === 'formats') {
          next = { ...(cur ?? {}), s: from.cell?.s ?? sheet.styleIdAt(from.r, from.c) };
        } else if (mode === 'noFormats') {
          const keep = cur?.s;
          next = next ? { ...next } : {};
          if (keep !== undefined) next.s = keep;
          else delete next.s;
        } else if (next && next.s === undefined) {
          const sid = sheet.styleIdAt(from.r, from.c);
          if (sid) next.s = sid;
        }
        if (next) {
          delete next.note;
          if (cur?.note) next.note = cur.note;
        }
        tx.setCell(sheet, r, c, next && Object.keys(next).length ? next : undefined);
      }
    }
    // Merges in source repeat across the target
    const srcMerges = sheet.merges.filter((m) => m.r1 >= src.r1 && m.r2 <= src.r2 && m.c1 >= src.c1 && m.c2 <= src.c2);
    if (srcMerges.length) {
      const h = src.r2 - src.r1 + 1;
      const w = src.c2 - src.c1 + 1;
      const add: Range[] = [];
      for (let r0 = target.r1; r0 + h - 1 <= target.r2; r0 += vertical ? h : target.r2 + 1)
        for (let c0 = target.c1; c0 + w - 1 <= target.c2; c0 += vertical ? target.c2 + 1 : w) {
          if (r0 === src.r1 && c0 === src.c1) continue;
          for (const m of srcMerges) add.push({ r1: m.r1 - src.r1 + r0, r2: m.r2 - src.r1 + r0, c1: m.c1 - src.c1 + c0, c2: m.c2 - src.c1 + c0 });
        }
      tx.setMeta(sheet, 'merges', sheet.merges.concat(add));
    }
  }, rangeSel(target));
}

export function fillDirection(dir: 'down' | 'right' | 'up' | 'left'): void {
  const st = S();
  const rg = primaryRange(st.sel);
  const sheet = st.wb.activeSheet;
  let src: Range;
  if (dir === 'down') {
    if (rg.r1 === rg.r2) {
      if (rg.r1 === 0) return;
      src = { ...rg, r1: rg.r1 - 1, r2: rg.r1 - 1 };
      autoFill(src, { ...rg, r1: rg.r1 - 1 }, 'copy');
      setSel(rg);
      return;
    }
    src = { ...rg, r2: rg.r1 };
  } else if (dir === 'right') {
    if (rg.c1 === rg.c2) {
      if (rg.c1 === 0) return;
      src = { ...rg, c1: rg.c1 - 1, c2: rg.c1 - 1 };
      autoFill(src, { ...rg, c1: rg.c1 - 1 }, 'copy');
      setSel(rg);
      return;
    }
    src = { ...rg, c2: rg.c1 };
  } else if (dir === 'up') src = { ...rg, r1: rg.r2 };
  else src = { ...rg, c1: rg.c2 };
  void sheet;
  autoFill(src, rg, 'copy');
}

function setSel(rg: Range): void {
  import('../store').then(({ setState }) => setState({ sel: rangeSel(rg) }));
}

export interface SeriesOptions {
  rowsOrCols: 'rows' | 'columns';
  type: 'linear' | 'growth' | 'date' | 'autofill';
  dateUnit: 'day' | 'weekday' | 'month' | 'year';
  step: number;
  stop?: number;
  trend: boolean;
}

export function fillSeries(o: SeriesOptions): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  const rg = primaryRange(st.sel);
  if (o.type === 'autofill') {
    const src = o.rowsOrCols === 'columns' ? { ...rg, r2: rg.r1 } : { ...rg, c2: rg.c1 };
    autoFill(src, rg, 'auto');
    return;
  }
  transact('Series', (tx) => {
    const lines = o.rowsOrCols === 'columns' ? rg.c2 - rg.c1 + 1 : rg.r2 - rg.r1 + 1;
    const len = o.rowsOrCols === 'columns' ? rg.r2 - rg.r1 + 1 : rg.c2 - rg.c1 + 1;
    for (let li = 0; li < lines; li++) {
      const r0 = o.rowsOrCols === 'columns' ? rg.r1 : rg.r1 + li;
      const c0 = o.rowsOrCols === 'columns' ? rg.c1 + li : rg.c1;
      const start = sheet.getCell(r0, c0)?.v;
      if (typeof start !== 'number') continue;
      let v = start;
      for (let k = 1; k < len; k++) {
        if (o.type === 'linear') v = v + o.step;
        else if (o.type === 'growth') v = v * o.step;
        else if (o.dateUnit === 'day') v = v + o.step;
        else if (o.dateUnit === 'weekday') v = addWeekdays(v, o.step);
        else if (o.dateUnit === 'month') v = addMonths(v, o.step);
        else v = addMonths(v, 12 * o.step);
        if (o.stop !== undefined && (o.step >= 0 ? v > o.stop : v < o.stop)) break;
        const r = o.rowsOrCols === 'columns' ? r0 + k : r0;
        const c = o.rowsOrCols === 'columns' ? c0 : c0 + k;
        const cur = sheet.getCell(r, c) ?? {};
        const sid = sheet.getCell(r0, c0)?.s;
        tx.setCell(sheet, r, c, { ...cur, v: +v.toPrecision(15), f: undefined, s: cur.s ?? sid });
        const cc = sheet.getCell(r, c);
        if (cc && cc.f === undefined) delete cc.f;
      }
    }
  });
}

// ---------- Flash Fill (lite) ----------

type Piece = { kind: 'lit'; text: string } | { kind: 'tok'; col: number; idx: number; fromEnd: boolean; tcase: 'same' | 'upper' | 'lower' | 'proper'; first: boolean };

function tokensOf(s: string): string[] {
  return s.match(/[A-Za-z0-9À-￿]+/g) ?? [];
}

function applyCase(t: string, c: Piece & { kind: 'tok' }): string {
  let x = c.first ? t.slice(0, 1) : t;
  if (c.tcase === 'upper') x = x.toUpperCase();
  else if (c.tcase === 'lower') x = x.toLowerCase();
  else if (c.tcase === 'proper') x = x.slice(0, 1).toUpperCase() + x.slice(1).toLowerCase();
  return x;
}

function runProgram(prog: Piece[], srcs: string[]): string | null {
  let out = '';
  for (const p of prog) {
    if (p.kind === 'lit') out += p.text;
    else {
      const toks = tokensOf(srcs[p.col] ?? '');
      const t = p.fromEnd ? toks[toks.length - 1 - p.idx] : toks[p.idx];
      if (t === undefined) return null;
      out += applyCase(t, p);
    }
  }
  return out;
}

function inferPrograms(srcs: string[], target: string): Piece[][] {
  // depth-first greedy search with limited branching
  const results: Piece[][] = [];
  const toks = srcs.map(tokensOf);
  const rec = (pos: number, acc: Piece[]) => {
    if (results.length > 40) return;
    if (pos === target.length) {
      results.push(acc);
      return;
    }
    const rest = target.slice(pos);
    let matched = false;
    toks.forEach((list, col) => {
      list.forEach((t, idx) => {
        for (const tcase of ['same', 'upper', 'lower', 'proper'] as const) {
          for (const first of [false, true]) {
            const piece = { kind: 'tok' as const, col, idx, fromEnd: false, tcase, first };
            const s = applyCase(t, piece);
            if (s && rest.startsWith(s) && (first || s.length > 0)) {
              matched = true;
              rec(pos + s.length, acc.concat([piece]));
              if (idx === list.length - 1) rec(pos + s.length, acc.concat([{ ...piece, idx: 0, fromEnd: true }]));
              break;
            }
          }
        }
      });
    });
    // literal character
    if (!matched || /[^A-Za-z0-9]/.test(rest[0])) {
      const last = acc[acc.length - 1];
      const nacc = last && last.kind === 'lit' ? acc.slice(0, -1).concat([{ kind: 'lit', text: last.text + rest[0] }]) : acc.concat([{ kind: 'lit', text: rest[0] }]);
      rec(pos + 1, nacc);
    }
  };
  rec(0, []);
  // prefer programs with fewer literals
  return results.sort((a, b) => a.filter((p) => p.kind === 'lit').reduce((n, p) => n + (p as { text: string }).text.length, 0) - b.filter((p) => p.kind === 'lit').reduce((n, p) => n + (p as { text: string }).text.length, 0));
}

export function flashFill(): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  const { r: ar, c: col } = st.sel.active;
  // source columns: contiguous non-empty columns to the left (and right)
  const srcCols: number[] = [];
  for (let c = col - 1; c >= 0 && sheet.getCell(ar, c)?.v !== undefined; c--) srcCols.unshift(c);
  for (let c = col + 1; sheet.getCell(ar, c)?.v !== undefined; c++) srcCols.push(c);
  if (!srcCols.length) {
    alertBox('Flash Fill needs example data. Type an example of the desired output in the column next to your data, then try again.');
    return;
  }
  // data rows: contiguous block around active cell based on first source column
  let r1 = ar;
  while (r1 > 0 && sheet.getCell(r1 - 1, srcCols[0])?.v !== undefined) r1--;
  let r2 = ar;
  while (sheet.getCell(r2 + 1, srcCols[0])?.v !== undefined) r2++;
  const headerLike = r1 < ar && typeof sheet.getCell(r1, col)?.v === 'string' && r1 === 0;
  const examples: { srcs: string[]; target: string }[] = [];
  for (let r = r1; r <= r2; r++) {
    const t = sheet.getCell(r, col)?.v;
    if (t === undefined || t === null || t === '') continue;
    if (headerLike && r === r1) continue;
    examples.push({ srcs: srcCols.map((c) => String(sheet.getCell(r, c)?.v ?? '')), target: String(t) });
  }
  if (!examples.length) {
    alertBox('Flash Fill needs at least one example. Type the value you want in the first cell and try again.');
    return;
  }
  const progs = inferPrograms(examples[0].srcs, examples[0].target);
  const prog = progs.find((p) => examples.every((e) => runProgram(p, e.srcs) === e.target));
  if (!prog) {
    alertBox("Flash Fill couldn't recognize a pattern to autofill values. To use Flash Fill, enter a couple of examples of the output you want, keep the active cell in the column you want filled in, and click the Flash Fill button again.");
    return;
  }
  let filled = 0;
  transact('Flash Fill', (tx) => {
    for (let r = r1; r <= r2; r++) {
      const cur = sheet.getCell(r, col);
      if (cur?.v !== undefined && cur.v !== '') continue;
      if (headerLike && r === r1) continue;
      const out = runProgram(prog, srcCols.map((c) => String(sheet.getCell(r, c)?.v ?? '')));
      if (out === null) continue;
      writeInput(tx, sheet, r, col, "'" + out);
      filled++;
    }
  });
  setStatus(`Flash Fill Changed Cells: ${filled}`);
}

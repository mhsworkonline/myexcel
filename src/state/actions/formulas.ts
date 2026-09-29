import { addrToA1, quoteSheetName, Range, rangeToA1 } from '../../model/address';
import { primaryRange, singleSel } from '../../model/selection';
import type { DefinedName } from '../../model/types';
import { alertBox, bump, S, setState, transact } from '../store';
import { FnCommand } from '../../model/commands';
import { beginEdit, writeInput } from './edit';
import { getComputed } from '../values';

export type AutoFn = 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN';

function isNum(r: number, c: number): boolean {
  const sheet = S().wb.activeSheet;
  return typeof getComputed(sheet, r, c) === 'number';
}

/** AutoSum (Alt+=): proposes =SUM(...) over the numbers above or to the left. */
export function autoSum(fn: AutoFn = 'SUM'): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  const rg = primaryRange(st.sel);
  const name = fn === 'COUNT' ? 'COUNT' : fn;
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) {
    const { r, c } = st.sel.active;
    let r1 = r - 1;
    while (r1 >= 0 && isNum(r1, c)) r1--;
    let range: Range | null = null;
    if (r1 < r - 1) range = { r1: r1 + 1, r2: r - 1, c1: c, c2: c };
    else {
      let c1 = c - 1;
      while (c1 >= 0 && isNum(r, c1)) c1--;
      if (c1 < c - 1) range = { r1: r, r2: r, c1: c1 + 1, c2: c - 1 };
    }
    const text = `=${name}(${range ? rangeToA1(range) : ''})`;
    beginEdit('enter', text);
    const ed = S().edit;
    if (ed) setState({ edit: { ...ed, caret: range ? text.length - 1 : text.length - 1 } });
    return;
  }
  // range selected: write totals below each column (or right of each row if single row)
  transact('AutoSum', (tx) => {
    if (rg.r1 === rg.r2) {
      const c = rg.c2 + 1;
      writeInput(tx, sheet, rg.r1, c, `=${name}(${rangeToA1(rg)})`);
    } else {
      for (let c = rg.c1; c <= rg.c2; c++) {
        const last = sheet.getCell(rg.r2, c);
        const target = last && (last.v !== undefined || last.f) ? rg.r2 + 1 : rg.r2;
        const src: Range = { r1: rg.r1, r2: target - 1, c1: c, c2: c };
        if (src.r2 < src.r1) continue;
        writeInput(tx, sheet, target, c, `=${name}(${rangeToA1(src)})`);
      }
    }
  });
}

// ---------- defined names ----------

export function validName(name: string): string | null {
  if (!/^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(name)) return 'The name that you entered is not valid.\n\nReasons for this can include:\n• The name does not begin with a letter or underscore\n• The name contains a space or other invalid characters\n• The name conflicts with an Excel built-in name or the name of another object in the workbook';
  if (/^[A-Za-z]{1,3}\d+$/.test(name) || /^R\d*C\d*$/i.test(name)) return 'The name that you entered is not valid. It looks like a cell reference.';
  if (name.length > 255) return 'The name is too long.';
  return null;
}

export function setNames(next: DefinedName[], label = 'Define Name'): void {
  const before = structuredClone(S().wb.names);
  const after = structuredClone(next);
  transact(label, (tx) => {
    tx.run(new FnCommand(label, (w) => w.setNames(structuredClone(after)), (w) => w.setNames(structuredClone(before))));
  });
}

export function defineName(name: string, ref: string, scope?: string, comment?: string, replace?: string): boolean {
  const err = validName(name);
  if (err) {
    alertBox(err);
    return false;
  }
  const wb = S().wb;
  const exists = wb.names.find((n) => n.name.toLowerCase() === name.toLowerCase() && n.scope === scope && n.name !== replace);
  if (exists) {
    alertBox('The name already exists. Enter a unique name.');
    return false;
  }
  const refText = ref.startsWith('=') ? ref.slice(1) : ref;
  const names = wb.names.filter((n) => n.name !== replace || n.scope !== scope);
  names.push({ name, ref: refText, scope, comment });
  setNames(names);
  return true;
}

export function deleteName(name: string, scope?: string): void {
  setNames(S().wb.names.filter((n) => !(n.name === name && n.scope === scope)), 'Delete Name');
}

/** Absolute reference text for the current selection, e.g. Sheet1!$A$1:$B$5 */
export function selectionRefText(): string {
  const st = S();
  const sheet = st.wb.activeSheet;
  return st.sel.ranges.map((rg) => `${quoteSheetName(sheet.name)}!${rangeToA1(rg, true)}`).join(',');
}

export function createNamesFromSelection(top: boolean, left: boolean, bottom: boolean, right: boolean): void {
  const st = S();
  const sheet = st.wb.activeSheet;
  const rg = primaryRange(st.sel);
  const add: DefinedName[] = [];
  const clean = (v: unknown) => String(v ?? '').trim().replace(/[^A-Za-z0-9_.]/g, '_').replace(/^(\d)/, '_$1');
  const pre = quoteSheetName(sheet.name) + '!';
  const abs = (r: Range) => pre + rangeToA1(r, true);
  const r1 = rg.r1 + (top ? 1 : 0);
  const r2 = rg.r2 - (bottom ? 1 : 0);
  const c1 = rg.c1 + (left ? 1 : 0);
  const c2 = rg.c2 - (right ? 1 : 0);
  if (top) for (let c = c1; c <= c2; c++) add.push({ name: clean(getComputed(sheet, rg.r1, c)), ref: abs({ r1, r2, c1: c, c2: c }) });
  if (bottom) for (let c = c1; c <= c2; c++) add.push({ name: clean(getComputed(sheet, rg.r2, c)), ref: abs({ r1, r2, c1: c, c2: c }) });
  if (left) for (let r = r1; r <= r2; r++) add.push({ name: clean(getComputed(sheet, r, rg.c1)), ref: abs({ r1: r, r2: r, c1, c2 }) });
  if (right) for (let r = r1; r <= r2; r++) add.push({ name: clean(getComputed(sheet, r, rg.c2)), ref: abs({ r1: r, r2: r, c1, c2 }) });
  const valid = add.filter((n) => n.name && !validName(n.name));
  const names = st.wb.names.filter((n) => !valid.some((v) => v.name.toLowerCase() === n.name.toLowerCase() && !n.scope));
  setNames(names.concat(valid), 'Create Names');
}

/** Name Box: typing a new name with cells selected defines it (Excel behaviour). */
export function nameBoxDefine(name: string): boolean {
  return defineName(name, selectionRefText());
}

export function showCalcModeNote(): void {
  bump();
}

export function traceCell(): string {
  const st = S();
  return addrToA1(st.sel.active.r, st.sel.active.c);
}

export { singleSel };

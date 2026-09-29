'use client';
import { Check, ChevronDown, X } from 'lucide-react';
import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { addrToA1, isFullCols, isFullRows, rangeToA1 } from '../model/address';
import { primaryRange } from '../model/selection';
import { beginEdit, cancelEdit, commitEdit } from '../state/actions/edit';
import { goTo, resolveReference } from '../state/actions/find';
import { nameBoxDefine, validName } from '../state/actions/formulas';
import { alertBox, openDialog, S, setState, useStore } from '../state/store';
import { editText } from '../state/values';
import { handleEditKey } from './grid/editKeys';
import { FormulaHighlight } from './grid/FormulaHighlight';
import { ArgHint, AutoList, useFormulaAssist } from './grid/useFormulaAssist';
import { Menu } from './Menu';
import { focusGrid } from './grid/focus';

function nameBoxText(): string {
  const st = S();
  const sheet = st.wb.activeSheet;
  const rg = primaryRange(st.sel);
  const single = rg.r1 === rg.r2 && rg.c1 === rg.c2;
  const m = sheet.mergeAt(rg.r1, rg.c1);
  const isMerge = m && m.r1 === rg.r1 && m.c1 === rg.c1 && m.r2 === rg.r2 && m.c2 === rg.c2;
  // a defined name that matches the selection exactly
  if (!single && st.sel.ranges.length === 1) {
    for (const n of st.wb.names) {
      const res = resolveReference(n.ref);
      if (res && res.sheetId === sheet.id && res.range.r1 === rg.r1 && res.range.r2 === rg.r2 && res.range.c1 === rg.c1 && res.range.c2 === rg.c2) return n.name;
    }
    const t = sheet.tables.find((x) => x.range.r1 + 1 === rg.r1 && x.range.r2 === rg.r2 && x.range.c1 === rg.c1 && x.range.c2 === rg.c2);
    if (t) return t.name;
  }
  if (isFullCols(rg) && isFullRows(rg)) return addrToA1(st.sel.active.r, st.sel.active.c);
  void isMerge;
  return addrToA1(st.sel.active.r, st.sel.active.c);
}

export function NameBox() {
  const sel = useStore((s) => s.sel);
  useStore((s) => s.rev);
  const [text, setText] = useState<string | null>(null);
  const [menu, setMenu] = useState<{ x: number; y: number } | null>(null);
  const ref = useRef<HTMLInputElement>(null);
  void sel;
  const commit = (v: string) => {
    const t = v.trim();
    setText(null);
    if (!t) return;
    if (goTo(t)) return;
    if (!validName(t)) {
      nameBoxDefine(t);
      return;
    }
    alertBox('Reference isn\'t valid.');
  };
  return (
    <div className="flex items-center h-[24px] border rounded-[3px] w-[110px] shrink-0" style={{ borderColor: 'var(--border-strong)', background: 'var(--input-bg)' }}>
      <input
        ref={ref}
        data-testid="name-box"
        aria-label="Name Box"
        className="flex-1 min-w-0 bg-transparent outline-none px-1.5 text-[12px]"
        value={text ?? nameBoxText()}
        onFocus={(e) => {
          setText(nameBoxText());
          setTimeout(() => e.target.select(), 0);
        }}
        onChange={(e) => setText(e.target.value)}
        onBlur={() => setText(null)}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            commit((e.target as HTMLInputElement).value);
            (e.target as HTMLInputElement).blur();
            focusGrid();
          } else if (e.key === 'Escape') {
            setText(null);
            (e.target as HTMLInputElement).blur();
          }
        }}
      />
      <button className="h-full px-1 hover:bg-[var(--hover)]" aria-label="Names" onClick={(e) => { const r = e.currentTarget.parentElement!.getBoundingClientRect(); setMenu({ x: r.left, y: r.bottom + 1 }); }}>
        <ChevronDown size={12} />
      </button>
      {menu && (
        <Menu
          x={menu.x}
          y={menu.y}
          minWidth={110}
          onClose={() => setMenu(null)}
          items={
            S().wb.names.length || S().wb.sheets.some((s) => s.tables.length)
              ? [...S().wb.names.map((n) => ({ label: n.name, onClick: () => goTo(n.name) })), ...S().wb.sheets.flatMap((s) => s.tables.map((t) => ({ label: t.name, onClick: () => goTo(t.name) })))]
              : [{ label: '(no names defined)', disabled: true }]
          }
        />
      )}
    </div>
  );
}

export function FormulaBar() {
  const sel = useStore((s) => s.sel);
  const edit = useStore((s) => s.edit);
  const expanded = useStore((s) => s.formulaBarExpanded);
  useStore((s) => s.rev);
  const taRef = useRef<HTMLTextAreaElement>(null);
  const sheet = S().wb.activeSheet;
  const { r, c } = sel.active;
  const style = S().wb.styles.get(sheet.styleIdAt(r, c));
  const hiddenFormula = !!sheet.protection && style.hidden && !!sheet.getCell(r, c)?.f;
  const shown = edit ? edit.text : hiddenFormula ? '' : editText(sheet, r, c);
  const barEditing = !!edit && edit.fromBar;
  const { auto, hint, accept, close, move, setIndex } = useFormulaAssist(edit?.text ?? '', edit?.caret ?? 0);

  useLayoutEffect(() => {
    const el = taRef.current;
    if (!el || !edit || !edit.fromBar) return;
    if (document.activeElement !== el) el.focus();
    if (el.selectionStart !== edit.caret) el.setSelectionRange(edit.caret, edit.caret);
  }, [edit]);

  useEffect(() => {
    if (document.body.dataset.hideFormulaBar) return;
  });
  if (typeof document !== 'undefined' && document.body.dataset.hideFormulaBar) return null;

  const height = expanded ? 72 : 26;
  return (
    <div className="flex items-start gap-1.5 px-2 py-[3px] relative" style={{ background: 'var(--chrome-bg)' }} data-testid="formula-bar">
      <NameBox />
      <div className="flex flex-col items-center justify-center h-[24px] px-0.5 opacity-50 select-none">⋮</div>
      <div className="flex items-center h-[24px] gap-0.5">
        <button className="w-6 h-6 flex items-center justify-center rounded hover:bg-[var(--hover)] disabled:opacity-30" disabled={!edit} title="Cancel" onClick={cancelEdit}>
          <X size={15} className="text-[#C42B1C]" />
        </button>
        <button className="w-6 h-6 flex items-center justify-center rounded hover:bg-[var(--hover)] disabled:opacity-30" disabled={!edit} title="Enter" onClick={() => commitEdit('none')}>
          <Check size={15} className="text-[#107C41]" />
        </button>
        <button className="w-7 h-6 flex items-center justify-center rounded hover:bg-[var(--hover)] italic font-serif text-[14px]" title="Insert Function" onClick={() => openDialog('insertFunction')} data-testid="fx-button">
          fx
        </button>
      </div>
      <div className="flex-1 relative border rounded-[3px] overflow-hidden" style={{ height, borderColor: 'var(--border-strong)', background: 'var(--input-bg)' }}>
        <div aria-hidden className="absolute inset-0 px-1.5 py-[3px] text-[13px] whitespace-pre-wrap break-all pointer-events-none leading-[18px] overflow-hidden" style={{ fontFamily: 'Segoe UI, Calibri, sans-serif' }}>
          <FormulaHighlight text={shown} color="var(--text)" />
        </div>
        <textarea
          ref={taRef}
          data-testid="formula-input"
          aria-label="Formula Bar"
          className="absolute inset-0 w-full h-full px-1.5 py-[3px] text-[13px] leading-[18px] bg-transparent outline-none resize-none border-0 whitespace-pre-wrap break-all"
          style={{ color: 'transparent', caretColor: 'var(--text)', fontFamily: 'Segoe UI, Calibri, sans-serif' }}
          value={shown}
          spellCheck={false}
          onFocus={(e) => {
            if (!S().edit) {
              const pos = e.target.selectionStart;
              beginEdit('edit', undefined, true);
              const ed = S().edit;
              if (ed) setState({ edit: { ...ed, caret: pos ?? ed.text.length, fromBar: true } });
            } else if (!S().edit!.fromBar) setState({ edit: { ...S().edit!, fromBar: true, mode: 'edit' } });
          }}
          onChange={(e) => {
            const ed = S().edit;
            if (!ed) return;
            setState({ edit: { ...ed, text: e.target.value, caret: e.target.selectionStart, point: undefined, fromBar: true } });
          }}
          onSelect={(e) => {
            const ed = S().edit;
            const el = e.currentTarget;
            if (ed && ed.fromBar && ed.caret !== el.selectionStart) setState({ edit: { ...ed, caret: el.selectionStart } });
          }}
          onKeyDown={(e) => {
            if (e.nativeEvent.isComposing) return;
            if (handleEditKey(e.nativeEvent, e.currentTarget, auto, accept, close, move)) {
              e.preventDefault();
              if (!S().edit) (document.querySelector('[data-testid="grid-input"]') as HTMLElement | null)?.focus();
            }
          }}
        />
      </div>
      {barEditing && auto && <AutoList auto={auto} onPick={accept} onHover={setIndex} style={{ left: 190, top: height + 6 }} />}
      {barEditing && !auto && hint && <ArgHint hint={hint} style={{ left: 190, top: height + 6 }} />}
      <button className="w-6 h-[24px] flex items-center justify-center rounded hover:bg-[var(--hover)]" title={expanded ? 'Collapse Formula Bar (Ctrl+Shift+U)' : 'Expand Formula Bar (Ctrl+Shift+U)'} onClick={() => setState({ formulaBarExpanded: !expanded })}>
        <ChevronDown size={14} className={expanded ? 'rotate-180' : ''} />
      </button>
    </div>
  );
}

export { rangeToA1 };

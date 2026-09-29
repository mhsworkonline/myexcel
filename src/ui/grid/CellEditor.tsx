'use client';
import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { DEFAULT_FONT_SIZE } from '../../model/styles';
import { cssFont, measureText } from '../../state/measure';
import { S, setState, useStore } from '../../state/store';
import { adaptColor } from '../theme';
import { handleEditKey } from './editKeys';
import { FormulaHighlight } from './FormulaHighlight';
import { ArgHint, AutoList, useFormulaAssist } from './useFormulaAssist';

interface Props {
  rect: { x: number; y: number; w: number; h: number };
  viewportWidth: number;
  viewportHeight: number;
}

export function CellEditor({ rect, viewportWidth, viewportHeight }: Props) {
  const edit = useStore((s) => s.edit)!;
  const theme = useStore((s) => s.theme);
  const ref = useRef<HTMLTextAreaElement>(null);
  const sheet = S().wb.sheetById(edit.hostSheetId)!;
  const style = S().wb.styles.get(sheet.styleIdAt(edit.r, edit.c));
  const zoom = sheet.zoom / 100;
  const font = cssFont(style, zoom);
  const { auto, hint, accept, close, move, setIndex } = useFormulaAssist(edit.text, edit.caret);
  const [composing, setComposing] = useState(false);

  // size: grows right with content (no wrap) or down (wrap)
  const lines = edit.text.split('\n');
  const textW = Math.max(...lines.map((l) => measureText(l, style))) * zoom + 12 * zoom;
  const wrap = !!style.wrap;
  const lineH = Math.round(((style.fontSize ?? DEFAULT_FONT_SIZE) * 96) / 72 * 1.2 * zoom);
  const w = wrap ? rect.w + 1 : Math.min(Math.max(rect.w + 1, textW), Math.max(rect.w + 1, viewportWidth - rect.x - 4));
  const h = Math.max(rect.h + 1, Math.min(viewportHeight - rect.y - 4, (wrap ? Math.ceil(textW / Math.max(20, rect.w)) + lines.length - 1 : lines.length) * lineH + 6));

  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    if (document.activeElement !== el && !edit.fromBar) el.focus({ preventScroll: true });
    if (el.selectionStart !== edit.caret || el.selectionEnd !== edit.caret) {
      if (document.activeElement === el) el.setSelectionRange(edit.caret, edit.caret);
    }
  }, [edit.caret, edit.fromBar, edit.text]);

  useEffect(() => {
    // start in the right place: Enter mode = end; F2 edit mode = end as well (Excel)
    const el = ref.current;
    if (el && !edit.fromBar) {
      el.focus({ preventScroll: true });
      el.setSelectionRange(edit.caret, edit.caret);
    }
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const onChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    const ed = S().edit;
    if (!ed) return;
    setState({ edit: { ...ed, text: e.target.value, caret: e.target.selectionStart, point: undefined, fromBar: false } });
  };

  const onSelect = (e: React.SyntheticEvent<HTMLTextAreaElement>) => {
    const ed = S().edit;
    const el = e.currentTarget;
    if (ed && ed.caret !== el.selectionStart && document.activeElement === el) setState({ edit: { ...ed, caret: el.selectionStart, point: ed.point && ed.point.end === el.selectionStart ? ed.point : undefined } });
  };

  const onKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (composing || e.nativeEvent.isComposing) return;
    if (handleEditKey(e.nativeEvent, e.currentTarget, auto, accept, close, move)) {
      e.preventDefault();
      e.stopPropagation();
    }
  };

  const bg = style.fillColor ? adaptColor(style.fillColor, 'fill', theme === 'dark') : 'var(--grid-bg)';
  const color = style.fontColor ? adaptColor(style.fontColor, 'text', theme === 'dark')! : 'var(--cell-text)';
  const align = style.hAlign === 'right' ? 'right' : style.hAlign === 'center' ? 'center' : 'left';
  const common: React.CSSProperties = {
    font,
    lineHeight: lineH + 'px',
    padding: `${Math.max(0, h - lineH * lines.length - 3) > 0 && !wrap && lines.length === 1 ? Math.max(1, rect.h - lineH - 2) : 1}px ${Math.round(2 * zoom)}px 0 ${Math.round(2 * zoom)}px`,
    whiteSpace: wrap ? 'pre-wrap' : 'pre',
    wordBreak: 'break-word',
    textAlign: edit.text.startsWith('=') ? 'left' : (align as 'left'),
    textDecoration: [style.underline ? 'underline' : '', style.strike ? 'line-through' : ''].join(' ').trim() || undefined,
  };
  return (
    <>
      <div
        className="absolute z-20 overflow-hidden"
        style={{ left: rect.x - 1, top: rect.y - 1, width: w, height: h, background: bg, boxShadow: '0 0 0 2px var(--accent)', border: 'none' }}
        onPointerDown={(e) => e.stopPropagation()}
        onMouseDown={(e) => e.stopPropagation()}
        onDoubleClick={(e) => e.stopPropagation()}
      >
        <div aria-hidden className="absolute inset-0 pointer-events-none" style={{ ...common, color }}>
          <FormulaHighlight text={edit.text} color={color} />
        </div>
        <textarea
          ref={ref}
          data-testid="cell-editor"
          value={edit.text}
          onChange={onChange}
          onSelect={onSelect}
          onKeyDown={onKeyDown}
          onCompositionStart={() => setComposing(true)}
          onCompositionEnd={() => setComposing(false)}
          onFocus={() => {
            const ed = S().edit;
            if (ed?.fromBar) setState({ edit: { ...ed, fromBar: false } });
          }}
          spellCheck={false}
          className="absolute inset-0 w-full h-full resize-none outline-none border-0 bg-transparent overflow-hidden"
          style={{ ...common, color: 'transparent', caretColor: theme === 'dark' ? '#fff' : '#000' }}
        />
      </div>
      {auto && <AutoList auto={auto} onPick={accept} onHover={setIndex} style={{ left: rect.x - 1, top: rect.y + h + 2 }} />}
      {!auto && hint && <ArgHint hint={hint} style={{ left: rect.x - 1, top: rect.y + h + 2 }} />}
    </>
  );
}

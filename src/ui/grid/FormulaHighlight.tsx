'use client';
import type { ReactElement } from 'react';
import { extractRefs, REF_COLORS_EXCEL } from '../../model/formula';

/** Mirror text with coloured references, rendered underneath a transparent-text textarea. */
export function FormulaHighlight({ text, color }: { text: string; color: string }) {
  if (!text.startsWith('=')) return <>{text + '​'}</>;
  const refs = extractRefs(text);
  const parts: ReactElement[] = [];
  let pos = 0;
  refs.forEach((r, i) => {
    if (r.start > pos) parts.push(<span key={'t' + i}>{text.slice(pos, r.start)}</span>);
    parts.push(
      <span key={'r' + i} style={{ color: REF_COLORS_EXCEL[r.colorIndex % REF_COLORS_EXCEL.length] }}>
        {text.slice(r.start, r.end)}
      </span>,
    );
    pos = r.end;
  });
  if (pos < text.length) parts.push(<span key="end">{text.slice(pos)}</span>);
  return <span style={{ color }}>{parts}{'​'}</span>;
}

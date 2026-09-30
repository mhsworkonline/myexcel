'use client';
import { FileText, Grid3x3, Minus, Plus, Rows3, Accessibility } from 'lucide-react';
import { useMemo } from 'react';
import { isErrorVal } from '../engine/Engine';
import { addrToA1 } from '../model/address';
import { formatValue } from '../model/numfmt';
import { canPoint } from '../state/actions/point';
import { bump, openDialog, S, setState, useStore } from '../state/store';
import { getComputed } from '../state/values';

function useSelectionStats() {
  const sel = useStore((s) => s.sel);
  const rev = useStore((s) => s.rev);
  return useMemo(() => {
    const st = S();
    const sheet = st.wb.activeSheet;
    let count = 0;
    let numCount = 0;
    let sum = 0;
    let min = Infinity;
    let max = -Infinity;
    const single = sel.ranges.length === 1 && sel.ranges[0].r1 === sel.ranges[0].r2 && sel.ranges[0].c1 === sel.ranges[0].c2;
    if (single) return null;
    const seen = new Set<string>();
    for (const rg of sel.ranges) {
      sheet.forEachInRange(rg, (r, c, cell) => {
        if (cell.v === undefined && cell.f === undefined) return;
        if (sheet.isRowHidden(r)) return;
        const k = r + ',' + c;
        if (seen.has(k)) return;
        seen.add(k);
        const v = getComputed(sheet, r, c);
        if (v === null || v === '') return;
        count++;
        if (typeof v === 'number' && !isErrorVal(v)) {
          numCount++;
          sum += v;
          if (v < min) min = v;
          if (v > max) max = v;
        }
      });
    }
    let fmt: string | undefined;
    const a = sel.active;
    fmt = st.wb.styles.get(sheet.styleIdAt(a.r, a.c)).numFmt;
    if (fmt && /[dmyhs]/i.test(fmt.replace(/"[^"]*"/g, '')) && !/0/.test(fmt)) fmt = undefined;
    return { count, numCount, sum, avg: numCount ? sum / numCount : 0, min, max, fmt };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sel, rev]);
}

export function StatusBar() {
  const edit = useStore((s) => s.edit);
  const status = useStore((s) => s.status);
  const viewMode = useStore((s) => s.viewMode);
  const extend = useStore((s) => s.extendMode);
  const add = useStore((s) => s.addMode);
  useStore((s) => s.rev);
  const stats = useSelectionStats();
  const sheet = S().wb.activeSheet;
  const zoom = sheet.zoom;
  const circ = S().engine.findCircular(sheet);
  let mode = 'Ready';
  if (edit) mode = canPoint() ? 'Point' : edit.mode === 'enter' ? 'Enter' : 'Edit';
  const setZoom = (z: number) => {
    sheet.zoom = Math.max(10, Math.min(400, Math.round(z)));
    sheet.touch();
    bump();
  };
  const f = (n: number) => formatValue(n, stats?.fmt ?? undefined).text;
  // slider maps 10..400 with 100 in the middle
  const toSlider = (z: number) => (z <= 100 ? ((z - 10) / 90) * 50 : 50 + ((z - 100) / 300) * 50);
  const fromSlider = (p: number) => (p <= 50 ? 10 + (p / 50) * 90 : 100 + ((p - 50) / 50) * 300);
  return (
    <div className="flex items-center h-[24px] px-2 text-[12px] gap-3 select-none" style={{ background: 'var(--chrome-bg)', color: 'var(--muted)', borderTop: '1px solid var(--border)' }} data-testid="status-bar">
      <span data-testid="status-mode" className="min-w-[40px]" style={{ color: 'var(--text)' }}>{mode}</span>
      {sheet.autoFilter && Object.keys(sheet.autoFilter.filters).length > 0 && <span>Filter Mode</span>}
      {extend && <span>Extend Selection</span>}
      {add && <span>Add to Selection</span>}
      {circ && <span style={{ color: '#C42B1C' }}>Circular References: {circ.sheet !== sheet.name ? circ.sheet + '!' : ''}{addrToA1(circ.r, circ.c)}</span>}
      {status && <span className="truncate" style={{ color: 'var(--text)' }}>{status}</span>}
      <span className="hidden md:flex items-center gap-1"><Accessibility size={13} /> Accessibility: Good to go</span>
      <div className="flex-1" />
      {stats && stats.count > 0 && (
        <div className="flex items-center gap-4" data-testid="status-stats" style={{ color: 'var(--text)' }}>
          {stats.numCount > 0 && <span>Average: {f(stats.avg)}</span>}
          <span>Count: {stats.count}</span>
          {stats.numCount > 0 && <span data-testid="status-sum">Sum: {f(stats.sum)}</span>}
        </div>
      )}
      <div className="flex items-center gap-0.5">
        <button className={'xl-status-btn ' + (viewMode === 'normal' ? 'xl-status-btn-active' : '')} title="Normal" onClick={() => import('../state/print').then((m) => m.setViewMode('normal'))}>
          <Grid3x3 size={14} />
        </button>
        <button className="xl-status-btn" title="Page Layout" onClick={() => openDialog('print')}>
          <FileText size={14} />
        </button>
        <button className={'xl-status-btn ' + (viewMode === 'pageBreak' ? 'xl-status-btn-active' : '')} title="Page Break Preview" onClick={() => import('../state/print').then((m) => m.setViewMode('pageBreak'))}>
          <Rows3 size={14} />
        </button>
      </div>
      <div className="flex items-center gap-1">
        <button className="xl-status-btn !w-5" title="Zoom Out" onClick={() => setZoom(zoom - 10)}>
          <Minus size={13} />
        </button>
        <input
          type="range"
          min={0}
          max={100}
          value={toSlider(zoom)}
          onChange={(e) => setZoom(fromSlider(+e.target.value))}
          className="w-[100px] accent-[var(--muted)]"
          aria-label="Zoom"
          data-testid="zoom-slider"
        />
        <button className="xl-status-btn !w-5" title="Zoom In" onClick={() => setZoom(zoom + 10)}>
          <Plus size={13} />
        </button>
        <button className="min-w-[40px] text-right hover:underline" onClick={() => openDialog('zoom')} data-testid="zoom-pct">
          {zoom}%
        </button>
      </div>
    </div>
  );
}

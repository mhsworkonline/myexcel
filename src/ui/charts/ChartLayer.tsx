'use client';
import { useEffect, useMemo, useRef, useState } from 'react';
import type { ChartSpec } from '../../model/types';
import { chartData, deleteChart, updateChart } from '../../state/charts';
import { openDialog, S, setState, useStore } from '../../state/store';
import { getViewport } from '../grid/Grid';
import { onScrollChange } from '../grid/scrollBus';
import { Menu } from '../Menu';
import { ChartSvg } from './ChartSvg';

/** Positions charts of the active sheet over the grid, in sheet coordinates. */
export function ChartLayer() {
  const rev = useStore((s) => s.rev);
  const selected = useStore((s) => s.selectedChartId);
  const theme = useStore((s) => s.theme);
  const [, force] = useState(0);
  useEffect(() => onScrollChange(() => force((n) => n + 1)), []);
  const sheet = S().wb.activeSheet;
  const vp = getViewport();
  // redraw once after the grid has computed its viewport
  useEffect(() => {
    const id = requestAnimationFrame(() => force((n) => n + 1));
    return () => cancelAnimationFrame(id);
  }, [rev]);
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const st = S();
      if (!st.selectedChartId || st.edit || st.dialog) return;
      if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        e.stopPropagation();
        deleteChart(st.selectedChartId);
      } else if (e.key === 'Escape') setState({ selectedChartId: null });
    };
    window.addEventListener('keydown', onKey, true);
    return () => window.removeEventListener('keydown', onKey, true);
  }, []);
  if (!vp || !sheet.charts.length) return null;
  const cp = vp.colPanes[vp.colPanes.length - 1];
  const rp = vp.rowPanes[vp.rowPanes.length - 1];
  const toScreen = (x: number, y: number) => ({ x: cp.start + x * vp.z - cp.content, y: rp.start + y * vp.z - rp.content });
  return (
    <div className="absolute overflow-hidden pointer-events-none" style={{ left: vp.hw, top: vp.hh, right: 17, bottom: 0 }}>
      {sheet.charts.map((ch) => {
        const p = toScreen(ch.anchor.x, ch.anchor.y);
        return <ChartBox key={ch.id} spec={ch} left={p.x - vp.hw} top={p.y - vp.hh} z={vp.z} selected={selected === ch.id} dark={theme === 'dark'} rev={rev} />;
      })}
    </div>
  );
}

function ChartBox({ spec, left, top, z, selected, dark, rev }: { spec: ChartSpec; left: number; top: number; z: number; selected: boolean; dark: boolean; rev: number }) {
  const sheet = S().wb.activeSheet;
  const data = useMemo(() => chartData(spec, sheet), [spec, sheet, rev]); // eslint-disable-line react-hooks/exhaustive-deps
  const [temp, setTemp] = useState<{ dx: number; dy: number; dw: number; dh: number } | null>(null);
  const [menu, setMenu] = useState<{ x: number; y: number } | null>(null);
  const ref = useRef<HTMLDivElement>(null);
  const w = (spec.anchor.w + (temp?.dw ?? 0)) * z;
  const h = (spec.anchor.h + (temp?.dh ?? 0)) * z;
  const x = left + (temp?.dx ?? 0) * z;
  const y = top + (temp?.dy ?? 0) * z;
  const startDrag = (e: React.PointerEvent, mode: string) => {
    e.stopPropagation();
    e.preventDefault();
    setState({ selectedChartId: spec.id, menu: null });
    const sx = e.clientX;
    const sy = e.clientY;
    let last = { dx: 0, dy: 0, dw: 0, dh: 0 };
    const move = (ev: PointerEvent) => {
      const ddx = (ev.clientX - sx) / z;
      const ddy = (ev.clientY - sy) / z;
      const t = { dx: 0, dy: 0, dw: 0, dh: 0 };
      if (mode === 'move') {
        t.dx = ddx;
        t.dy = ddy;
      } else {
        if (mode.includes('e')) t.dw = ddx;
        if (mode.includes('s')) t.dh = ddy;
        if (mode.includes('w')) {
          t.dx = ddx;
          t.dw = -ddx;
        }
        if (mode.includes('n')) {
          t.dy = ddy;
          t.dh = -ddy;
        }
      }
      last = t;
      setTemp(t);
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      setTemp(null);
      if (last.dx || last.dy || last.dw || last.dh) {
        const a = spec.anchor;
        updateChart(spec.id, { anchor: { x: Math.max(0, a.x + last.dx), y: Math.max(0, a.y + last.dy), w: Math.max(80, a.w + last.dw), h: Math.max(60, a.h + last.dh) } }, mode === 'move' ? 'Move Chart' : 'Resize Chart');
      }
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };
  const handle = (pos: string, style: React.CSSProperties) => (
    <div key={pos} className="absolute w-[8px] h-[8px] bg-white border pointer-events-auto" style={{ ...style, borderColor: '#7F7F7F', cursor: `${pos}-resize` }} onPointerDown={(e) => startDrag(e, pos)} />
  );
  return (
    <div
      ref={ref}
      className="absolute pointer-events-auto"
      style={{ left: x, top: y, width: w, height: h, boxShadow: selected ? '0 0 0 1px #7F7F7F' : '0 1px 3px rgba(0,0,0,0.15)', cursor: 'move' }}
      onPointerDown={(e) => startDrag(e, 'move')}
      onDoubleClick={(e) => {
        e.stopPropagation();
        openDialog('chartFormat', { id: spec.id });
      }}
      onContextMenu={(e) => {
        e.preventDefault();
        e.stopPropagation();
        setState({ selectedChartId: spec.id });
        setMenu({ x: e.clientX, y: e.clientY });
      }}
      data-testid={`chart-${spec.id}`}
      data-chart
    >
      <ChartSvg spec={spec} data={data} width={w} height={h} dark={dark} />
      {selected && (
        <>
          {handle('nw', { left: -4, top: -4 })}
          {handle('n', { left: w / 2 - 4, top: -4 })}
          {handle('ne', { right: -4, top: -4 })}
          {handle('e', { right: -4, top: h / 2 - 4 })}
          {handle('se', { right: -4, bottom: -4 })}
          {handle('s', { left: w / 2 - 4, bottom: -4 })}
          {handle('sw', { left: -4, bottom: -4 })}
          {handle('w', { left: -4, top: h / 2 - 4 })}
        </>
      )}
      {menu && (
        <Menu
          x={menu.x}
          y={menu.y}
          onClose={() => setMenu(null)}
          items={[
            { label: 'Delete', onClick: () => deleteChart(spec.id) },
            { separator: true },
            { label: 'Change Chart Type...', onClick: () => openDialog('insertChart', { changeId: spec.id }) },
            { label: 'Select Data...', onClick: () => openDialog('chartData', { id: spec.id }) },
            { separator: true },
            { label: 'Format Chart Area...', onClick: () => openDialog('chartFormat', { id: spec.id }) },
          ]}
        />
      )}
    </div>
  );
}

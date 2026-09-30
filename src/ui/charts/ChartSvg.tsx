'use client';
import type { ReactElement } from 'react';
import type { ChartSpec } from '../../model/types';
import type { ChartData, SeriesData } from '../../state/charts';

const FONT = '"Calibri", "Carlito", "Segoe UI", Arial, sans-serif';

export function niceScale(min: number, max: number, maxTicks = 6): { min: number; max: number; step: number } {
  if (!isFinite(min) || !isFinite(max)) return { min: 0, max: 1, step: 0.2 };
  if (min === max) {
    if (min === 0) return { min: 0, max: 1, step: 0.2 };
    min = Math.min(0, min);
    max = Math.max(0, max);
  }
  const range = max - min;
  const rough = range / maxTicks;
  const mag = Math.pow(10, Math.floor(Math.log10(rough)));
  const norm = rough / mag;
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag;
  const nmin = Math.floor(min / step) * step;
  const nmax = Math.ceil(max / step) * step;
  return { min: nmin, max: nmax === nmin ? nmin + step : nmax, step };
}

function fmtNum(n: number): string {
  const a = Math.abs(n);
  if (a >= 1e9) return +(n / 1e9).toFixed(1) + 'B';
  if (a >= 1e6) return +(n / 1e6).toFixed(1) + 'M';
  if (a >= 1e4) return Math.round(n).toLocaleString('en-US');
  if (Number.isInteger(n)) return n.toLocaleString('en-US');
  return String(+n.toPrecision(6));
}

interface Props {
  spec: ChartSpec;
  data: ChartData;
  width: number;
  height: number;
  dark?: boolean;
}

export function ChartSvg({ spec, data, width, height, dark }: Props) {
  const text = dark ? '#D0D0D0' : '#595959';
  const grid = dark ? '#444' : '#D9D9D9';
  const axisLine = dark ? '#666' : '#BFBFBF';
  const bg = dark ? '#262626' : '#FFFFFF';
  const W = Math.max(60, width);
  const H = Math.max(40, height);
  const pad = 10;
  const titleH = spec.showTitle && spec.title ? 30 : 0;
  const isPie = spec.type === 'pie' || spec.type === 'doughnut';
  const legendItems = isPie ? data.categories.map((c, i) => ({ name: c, color: pieColor(data, i) })) : data.series.map((s) => ({ name: s.name, color: s.color }));
  const legend = spec.legend;
  const legendH = legend === 'bottom' || legend === 'top' ? 22 : 0;
  const legendW = legend === 'right' || legend === 'left' ? Math.min(140, 30 + Math.max(0, ...legendItems.map((l) => l.name.length)) * 6.5) : 0;
  const plot = {
    x: pad + (legend === 'left' ? legendW : 0),
    y: pad + titleH + (legend === 'top' ? legendH : 0),
    w: W - 2 * pad - legendW,
    h: H - 2 * pad - titleH - legendH,
  };
  return (
    <svg width={W} height={H} viewBox={`0 0 ${W} ${H}`} style={{ fontFamily: FONT, display: 'block', background: bg }} xmlns="http://www.w3.org/2000/svg">
      <rect x={0.5} y={0.5} width={W - 1} height={H - 1} fill={bg} stroke={dark ? '#555' : '#D9D9D9'} />
      {titleH > 0 && (
        <text x={W / 2} y={pad + 18} textAnchor="middle" fontSize={18.6} fill={text}>
          {spec.title}
        </text>
      )}
      {isPie ? <Pie spec={spec} data={data} box={plot} text={text} dark={!!dark} /> : spec.type === 'scatter' ? <Scatter spec={spec} data={data} box={plot} text={text} grid={grid} axisLine={axisLine} /> : <Cartesian spec={spec} data={data} box={plot} text={text} grid={grid} axisLine={axisLine} />}
      {legend !== 'none' && legendItems.length > 0 && <Legend items={legendItems} pos={legend} W={W} H={H} pad={pad} titleH={titleH} legendW={legendW} text={text} />}
    </svg>
  );
}

function pieColor(data: ChartData, i: number): string {
  const base = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E', '#636363', '#997300'];
  const s0 = data.series[0];
  if (s0 && s0.color !== '#4472C4') {
    // tint variations of a monochrome palette
    const n = parseInt(s0.color.slice(1), 16);
    const t = (i % 6) * 0.14;
    const ch = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => Math.round(v + (255 - v) * t));
    return '#' + ch.map((v) => v.toString(16).padStart(2, '0')).join('');
  }
  return base[i % base.length];
}

function Legend({ items, pos, W, H, pad, titleH, legendW, text }: { items: { name: string; color: string }[]; pos: string; W: number; H: number; pad: number; titleH: number; legendW: number; text: string }) {
  if (pos === 'bottom' || pos === 'top') {
    const widths = items.map((i) => 18 + i.name.length * 6.3 + 12);
    const total = widths.reduce((a, b) => a + b, 0);
    let x = Math.max(pad, (W - total) / 2);
    const y = pos === 'bottom' ? H - pad - 8 : pad + titleH + 8;
    return (
      <g>
        {items.map((it, i) => {
          const gx = x;
          x += widths[i];
          return (
            <g key={i} transform={`translate(${gx},${y})`}>
              <rect x={0} y={-5} width={9} height={9} fill={it.color} />
              <text x={14} y={3.5} fontSize={12} fill={text}>{it.name}</text>
            </g>
          );
        })}
      </g>
    );
  }
  const x = pos === 'right' ? W - pad - legendW + 8 : pad;
  const y0 = pad + titleH + Math.max(10, (H - titleH - items.length * 18) / 2);
  return (
    <g>
      {items.map((it, i) => (
        <g key={i} transform={`translate(${x},${y0 + i * 18})`}>
          <rect x={0} y={-5} width={9} height={9} fill={it.color} />
          <text x={14} y={3.5} fontSize={12} fill={text}>{it.name.length > 18 ? it.name.slice(0, 17) + '…' : it.name}</text>
        </g>
      ))}
    </g>
  );
}

interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

function Cartesian({ spec, data, box, text, grid, axisLine }: { spec: ChartSpec; data: ChartData; box: Box; text: string; grid: string; axisLine: string }) {
  const horizontal = spec.type === 'bar' || spec.type === 'stackedBar';
  const stacked = spec.type === 'stackedColumn' || spec.type === 'stackedBar' || spec.type === 'stackedArea';
  const primary = data.series.filter((s) => !s.secondary);
  const secondary = data.series.filter((s) => s.secondary);
  const n = data.categories.length;
  const range = (list: SeriesData[]) => {
    let lo = 0;
    let hi = 0;
    if (stacked) {
      for (let i = 0; i < n; i++) {
        let pos = 0;
        let neg = 0;
        for (const s of list) {
          const v = s.values[i] ?? 0;
          if (v >= 0) pos += v;
          else neg += v;
        }
        hi = Math.max(hi, pos);
        lo = Math.min(lo, neg);
      }
    } else
      for (const s of list)
        for (const v of s.values) {
          if (v === null) continue;
          hi = Math.max(hi, v);
          lo = Math.min(lo, v);
        }
    return niceScale(lo, hi);
  };
  const sc = range(primary.length ? primary : data.series);
  const sc2 = secondary.length ? range(secondary) : null;
  const labelW = horizontal ? Math.min(100, 8 + Math.max(0, ...data.categories.map((c) => c.length)) * 6) : Math.max(24, fmtNum(sc.max).length * 7 + 8, fmtNum(sc.min).length * 7 + 8);
  const labelH = horizontal ? 18 : 18;
  const rightW = sc2 ? Math.max(24, fmtNum(sc2.max).length * 7 + 8) : 6;
  const titleY = spec.yTitle ? 18 : 0;
  const titleX = spec.xTitle ? 18 : 0;
  const p = { x: box.x + labelW + titleY, y: box.y + 6, w: box.w - labelW - rightW - titleY, h: box.h - labelH - 6 - titleX };
  if (p.w < 10 || p.h < 10) return null;
  const valTo = (v: number, s = sc) => (horizontal ? p.x + ((v - s.min) / (s.max - s.min)) * p.w : p.y + p.h - ((v - s.min) / (s.max - s.min)) * p.h);
  const catSize = (horizontal ? p.h : p.w) / Math.max(1, n);
  const catPos = (i: number) => (horizontal ? p.y + catSize * i : p.x + catSize * i);
  const ticks: number[] = [];
  for (let t = sc.min; t <= sc.max + sc.step / 2; t += sc.step) ticks.push(+t.toPrecision(12));
  const els: ReactElement[] = [];
  // gridlines + value labels
  ticks.forEach((t, i) => {
    const pos = valTo(t);
    if (horizontal) {
      if (spec.gridlines) els.push(<line key={'g' + i} x1={pos} x2={pos} y1={p.y} y2={p.y + p.h} stroke={grid} />);
      els.push(<text key={'t' + i} x={pos} y={p.y + p.h + 14} fontSize={12} fill={text} textAnchor="middle">{fmtNum(t)}</text>);
    } else {
      if (spec.gridlines) els.push(<line key={'g' + i} x1={p.x} x2={p.x + p.w} y1={pos} y2={pos} stroke={grid} />);
      els.push(<text key={'t' + i} x={p.x - 6} y={pos + 4} fontSize={12} fill={text} textAnchor="end">{fmtNum(t)}</text>);
    }
  });
  if (sc2) {
    for (let t = sc2.min; t <= sc2.max + sc2.step / 2; t += sc2.step) {
      const pos = valTo(t, sc2);
      els.push(<text key={'t2' + t} x={p.x + p.w + 6} y={pos + 4} fontSize={12} fill={text}>{fmtNum(+t.toPrecision(12))}</text>);
    }
  }
  // category axis line + labels
  const zero = valTo(Math.max(sc.min, Math.min(sc.max, 0)));
  if (horizontal) els.push(<line key="ax" x1={zero} x2={zero} y1={p.y} y2={p.y + p.h} stroke={axisLine} />);
  else els.push(<line key="ax" x1={p.x} x2={p.x + p.w} y1={zero} y2={zero} stroke={axisLine} />);
  const every = Math.max(1, Math.ceil((horizontal ? 14 : 7 * Math.max(...data.categories.map((c) => c.length), 1)) / Math.max(1, catSize)));
  data.categories.forEach((c, i) => {
    if (i % every) return;
    if (horizontal) els.push(<text key={'c' + i} x={p.x - 6} y={catPos(i) + catSize / 2 + 4} fontSize={12} fill={text} textAnchor="end">{c}</text>);
    else els.push(<text key={'c' + i} x={catPos(i) + catSize / 2} y={p.y + p.h + 14} fontSize={12} fill={text} textAnchor="middle">{c.length > 14 ? c.slice(0, 13) + '…' : c}</text>);
  });
  // series
  const colSeries = data.series.filter((s) => {
    const t = s.seriesType ?? (spec.type === 'combo' ? 'column' : spec.type.includes('Column') || spec.type === 'column' || spec.type.includes('Bar') || spec.type === 'bar' ? 'column' : spec.type.includes('rea') ? 'area' : 'line');
    return t === 'column';
  });
  const groupW = catSize * 0.64;
  const barW = stacked ? groupW : groupW / Math.max(1, colSeries.length);
  const posStack = new Array(n).fill(0);
  const negStack = new Array(n).fill(0);
  colSeries.forEach((s, si) => {
    const sScale = s.secondary && sc2 ? sc2 : sc;
    s.values.forEach((v, i) => {
      if (v === null) return;
      const base = stacked ? (v >= 0 ? posStack[i] : negStack[i]) : 0;
      const top = base + v;
      if (stacked) {
        if (v >= 0) posStack[i] = top;
        else negStack[i] = top;
      }
      const a = valTo(base, sScale);
      const b = valTo(top, sScale);
      const off = catPos(i) + (catSize - groupW) / 2 + (stacked ? 0 : barW * si);
      const rect = horizontal
        ? { x: Math.min(a, b), y: off, width: Math.abs(b - a), height: barW - (stacked ? 0 : 1) }
        : { x: off, y: Math.min(a, b), width: barW - (stacked ? 0 : 1), height: Math.abs(b - a) };
      els.push(<rect key={`b${si}-${i}`} {...rect} fill={s.color} />);
      if (spec.dataLabels) {
        els.push(
          <text key={`bl${si}-${i}`} x={horizontal ? Math.max(a, b) + 4 : rect.x + rect.width / 2} y={horizontal ? rect.y + rect.height / 2 + 4 : Math.min(a, b) - 3} fontSize={11} fill={text} textAnchor={horizontal ? 'start' : 'middle'}>
            {fmtNum(v)}
          </text>,
        );
      }
    });
  });
  // area / line
  const others = data.series.filter((s) => !colSeries.includes(s));
  const areaStack = new Array(n).fill(0);
  others.forEach((s, si) => {
    const t = s.seriesType ?? (spec.type.includes('rea') ? 'area' : 'line');
    const sScale = s.secondary && sc2 ? sc2 : sc;
    const pts: [number, number, number][] = [];
    s.values.forEach((v, i) => {
      if (v === null) return;
      const base = t === 'area' && stacked ? areaStack[i] : 0;
      const val = base + v;
      if (t === 'area' && stacked) areaStack[i] = val;
      pts.push([catPos(i) + catSize / 2, valTo(val, sScale), valTo(base, sScale)]);
    });
    if (!pts.length) return;
    if (t === 'area') {
      const top = pts.map((q) => `${q[0]},${q[1]}`).join(' ');
      const bottom = [...pts].reverse().map((q) => `${q[0]},${q[2]}`).join(' ');
      els.push(<polygon key={'a' + si} points={`${top} ${bottom}`} fill={s.color} fillOpacity={stacked ? 1 : 0.85} />);
    } else {
      els.push(<polyline key={'l' + si} points={pts.map((q) => `${q[0]},${q[1]}`).join(' ')} fill="none" stroke={s.color} strokeWidth={2.25} strokeLinejoin="round" strokeLinecap="round" />);
      if (spec.type === 'lineMarkers' || spec.type === 'combo') pts.forEach((q, i) => els.push(<circle key={`m${si}-${i}`} cx={q[0]} cy={q[1]} r={3.5} fill={s.color} stroke="#fff" strokeWidth={1} />));
    }
    if (spec.dataLabels) s.values.forEach((v, i) => v !== null && els.push(<text key={`ll${si}-${i}`} x={catPos(i) + catSize / 2} y={valTo(v, sScale) - 7} fontSize={11} fill={text} textAnchor="middle">{fmtNum(v)}</text>));
  });
  if (spec.yTitle) els.push(<text key="yt" x={box.x + 10} y={p.y + p.h / 2} fontSize={13} fill={text} textAnchor="middle" transform={`rotate(-90 ${box.x + 10} ${p.y + p.h / 2})`}>{spec.yTitle}</text>);
  if (spec.xTitle) els.push(<text key="xt" x={p.x + p.w / 2} y={box.y + box.h - 2} fontSize={13} fill={text} textAnchor="middle">{spec.xTitle}</text>);
  return <g>{els}</g>;
}

function Scatter({ spec, data, box, text, grid, axisLine }: { spec: ChartSpec; data: ChartData; box: Box; text: string; grid: string; axisLine: string }) {
  const xs = data.series.flatMap((s) => (s.xValues ?? s.values.map((_, i) => i + 1)).filter((v): v is number => v !== null));
  const ys = data.series.flatMap((s) => s.values.filter((v): v is number => v !== null));
  const sx = niceScale(Math.min(0, ...xs), Math.max(...xs, 1));
  const sy = niceScale(Math.min(0, ...ys), Math.max(...ys, 1));
  const labelW = Math.max(24, fmtNum(sy.max).length * 7 + 8);
  const p = { x: box.x + labelW, y: box.y + 6, w: box.w - labelW - 8, h: box.h - 24 };
  if (p.w < 10 || p.h < 10) return null;
  const X = (v: number) => p.x + ((v - sx.min) / (sx.max - sx.min)) * p.w;
  const Y = (v: number) => p.y + p.h - ((v - sy.min) / (sy.max - sy.min)) * p.h;
  const els: ReactElement[] = [];
  for (let t = sy.min; t <= sy.max + sy.step / 2; t += sy.step) {
    if (spec.gridlines) els.push(<line key={'gy' + t} x1={p.x} x2={p.x + p.w} y1={Y(t)} y2={Y(t)} stroke={grid} />);
    els.push(<text key={'ty' + t} x={p.x - 6} y={Y(t) + 4} fontSize={12} fill={text} textAnchor="end">{fmtNum(+t.toPrecision(12))}</text>);
  }
  for (let t = sx.min; t <= sx.max + sx.step / 2; t += sx.step) els.push(<text key={'tx' + t} x={X(t)} y={p.y + p.h + 15} fontSize={12} fill={text} textAnchor="middle">{fmtNum(+t.toPrecision(12))}</text>);
  els.push(<line key="axx" x1={p.x} x2={p.x + p.w} y1={Y(Math.max(sy.min, 0))} y2={Y(Math.max(sy.min, 0))} stroke={axisLine} />);
  els.push(<line key="axy" x1={X(Math.max(sx.min, 0))} x2={X(Math.max(sx.min, 0))} y1={p.y} y2={p.y + p.h} stroke={axisLine} />);
  data.series.forEach((s, si) => {
    const xv = s.xValues ?? s.values.map((_, i) => i + 1);
    s.values.forEach((v, i) => {
      const x = xv[i];
      if (v === null || x === null || x === undefined) return;
      els.push(<circle key={`p${si}-${i}`} cx={X(x)} cy={Y(v)} r={3.8} fill={s.color} />);
      if (spec.dataLabels) els.push(<text key={`pl${si}-${i}`} x={X(x) + 5} y={Y(v) - 5} fontSize={11} fill={text}>{fmtNum(v)}</text>);
    });
  });
  return <g>{els}</g>;
}

function Pie({ spec, data, box, text, dark }: { spec: ChartSpec; data: ChartData; box: Box; text: string; dark: boolean }) {
  const s = data.series[0];
  if (!s) return null;
  const vals = s.values.map((v) => (v && v > 0 ? v : 0));
  const total = vals.reduce((a, b) => a + b, 0);
  if (!total) return null;
  const cx = box.x + box.w / 2;
  const cy = box.y + box.h / 2;
  const r = Math.max(10, Math.min(box.w, box.h) / 2 - 6);
  const inner = spec.type === 'doughnut' ? r * 0.5 : 0;
  let ang = -Math.PI / 2;
  const els: ReactElement[] = [];
  vals.forEach((v, i) => {
    if (!v) return;
    const a2 = ang + (v / total) * Math.PI * 2;
    const large = a2 - ang > Math.PI ? 1 : 0;
    const p = (a: number, rad: number) => `${cx + Math.cos(a) * rad},${cy + Math.sin(a) * rad}`;
    const d =
      v === total
        ? inner
          ? `M ${cx + r} ${cy} A ${r} ${r} 0 1 1 ${cx - r} ${cy} A ${r} ${r} 0 1 1 ${cx + r} ${cy} M ${cx + inner} ${cy} A ${inner} ${inner} 0 1 0 ${cx - inner} ${cy} A ${inner} ${inner} 0 1 0 ${cx + inner} ${cy}`
          : `M ${cx + r} ${cy} A ${r} ${r} 0 1 1 ${cx - r} ${cy} A ${r} ${r} 0 1 1 ${cx + r} ${cy} Z`
        : inner
          ? `M ${p(ang, r)} A ${r} ${r} 0 ${large} 1 ${p(a2, r)} L ${p(a2, inner)} A ${inner} ${inner} 0 ${large} 0 ${p(ang, inner)} Z`
          : `M ${cx} ${cy} L ${p(ang, r)} A ${r} ${r} 0 ${large} 1 ${p(a2, r)} Z`;
    els.push(<path key={i} d={d} fill={pieColor(data, i)} stroke={dark ? '#262626' : '#FFFFFF'} strokeWidth={1.5} fillRule="evenodd" />);
    if (spec.dataLabels) {
      const mid = (ang + a2) / 2;
      const lr = inner ? (r + inner) / 2 : r * 0.65;
      els.push(<text key={'t' + i} x={cx + Math.cos(mid) * lr} y={cy + Math.sin(mid) * lr + 4} fontSize={11} fill="#fff" textAnchor="middle">{Math.round((v / total) * 100)}%</text>);
    }
    ang = a2;
  });
  void text;
  return <g>{els}</g>;
}

'use client';
import { useMemo, useState } from 'react';
import type { ChartSeries, ChartSpec, ChartType } from '../../model/types';
import { changeChartType, chartData, chartSourceRange, insertChart, PALETTES, selectedChart, seriesFromRange, switchRowCol, updateChart } from '../../state/charts';
import { openDialog, S, setState, useStore } from '../../state/store';
import { Dialog, RefInput } from '../dialogs/Dialog';
import { Col, Group, LargeButton, Row, SmallButton } from '../ribbon/parts';
import { ChartSvg } from './ChartSvg';
import { ChartColumnIcon } from '../icons';
import { ArrowLeftRight, Database, Palette, PlusSquare, Trash2, Type } from 'lucide-react';
import { deleteChart } from '../../state/charts';

export const CHART_TYPES: { type: ChartType; name: string; group: string }[] = [
  { type: 'column', name: 'Clustered Column', group: 'Column' },
  { type: 'stackedColumn', name: 'Stacked Column', group: 'Column' },
  { type: 'bar', name: 'Clustered Bar', group: 'Bar' },
  { type: 'stackedBar', name: 'Stacked Bar', group: 'Bar' },
  { type: 'line', name: 'Line', group: 'Line' },
  { type: 'lineMarkers', name: 'Line with Markers', group: 'Line' },
  { type: 'pie', name: 'Pie', group: 'Pie' },
  { type: 'doughnut', name: 'Doughnut', group: 'Pie' },
  { type: 'area', name: 'Area', group: 'Area' },
  { type: 'stackedArea', name: 'Stacked Area', group: 'Area' },
  { type: 'scatter', name: 'Scatter', group: 'X Y (Scatter)' },
  { type: 'combo', name: 'Clustered Column - Line', group: 'Combo' },
];

function previewSpec(type: ChartType): { spec: ChartSpec; host: ReturnType<typeof S>['wb']['activeSheet'] } | null {
  const src = chartSourceRange();
  if (!src) return null;
  const { series, title } = seriesFromRange(src.sheet, src.range, type);
  return {
    spec: { id: 'preview', type, title: title ?? 'Chart Title', showTitle: true, legend: series.length > 1 || type === 'pie' || type === 'doughnut' ? 'bottom' : 'none', dataLabels: false, gridlines: type !== 'pie' && type !== 'doughnut', style: 1, series, anchor: { x: 0, y: 0, w: 480, h: 288 } },
    host: src.sheet,
  };
}

export function InsertChartDialog({ props }: { props: Record<string, unknown> }) {
  const changeId = props.changeId as string | undefined;
  const [type, setType] = useState<ChartType>((props.type as ChartType) ?? (changeId ? S().wb.activeSheet.charts.find((c) => c.id === changeId)?.type ?? 'column' : 'column'));
  const preview = useMemo(() => previewSpec(type), [type]);
  const groups = [...new Set(CHART_TYPES.map((t) => t.group))];
  return (
    <Dialog title={changeId ? 'Change Chart Type' : 'Insert Chart'} width={760} onOk={() => (changeId ? changeChartType(changeId, type) : insertChart(type))} testId="insert-chart">
      <div className="flex gap-3 h-[400px]">
        <div className="w-[150px] xl-list">
          {groups.map((g) => (
            <div key={g} className={CHART_TYPES.find((t) => t.type === type)?.group === g ? 'xl-list-sel' : ''} onClick={() => setType(CHART_TYPES.find((t) => t.group === g)!.type)}>{g}</div>
          ))}
        </div>
        <div className="flex-1 flex flex-col gap-2">
          <div className="flex gap-2 flex-wrap">
            {CHART_TYPES.filter((t) => t.group === CHART_TYPES.find((x) => x.type === type)?.group).map((t) => (
              <button key={t.type} className={'border px-2 py-1 rounded ' + (t.type === type ? 'xl-item-active' : '')} style={{ borderColor: 'var(--border)' }} onClick={() => setType(t.type)} data-testid={`ct-${t.type}`}>{t.name}</button>
            ))}
          </div>
          <div className="font-semibold">{CHART_TYPES.find((t) => t.type === type)?.name}</div>
          <div className="flex-1 flex items-center justify-center border rounded" style={{ borderColor: 'var(--border)' }}>
            {preview ? <ChartSvg spec={preview.spec} data={chartData(preview.spec, preview.host)} width={520} height={310} dark={S().theme === 'dark'} /> : <span className="opacity-60">Select data to see a preview.</span>}
          </div>
        </div>
      </div>
    </Dialog>
  );
}

export function ChartDataDialog({ props }: { props: Record<string, unknown> }) {
  const ch = S().wb.activeSheet.charts.find((c) => c.id === props.id);
  const [series, setSeries] = useState<ChartSeries[]>(ch?.series ?? []);
  const [sel, setSel] = useState(0);
  const [cats, setCats] = useState(ch?.series[0]?.categories ?? '');
  if (!ch) return null;
  const cur = series[sel];
  const upd = (p: Partial<ChartSeries>) => setSeries(series.map((s, i) => (i === sel ? { ...s, ...p } : s)));
  return (
    <Dialog title="Select Data Source" width={620} onOk={() => updateChart(ch.id, { series: series.map((s) => ({ ...s, categories: cats || undefined })) }, 'Select Data')} testId="chart-data">
      <label className="flex items-center gap-2 mb-2">Chart data range: <span className="opacity-70">{ch.sourceRange ?? '(custom)'}</span></label>
      <div className="flex gap-3">
        <div className="flex-1">
          <div className="flex items-center gap-1 mb-1">
            <b className="flex-1">Legend Entries (Series)</b>
            <button className="xl-btn !min-w-0" onClick={() => { setSeries([...series, { values: '', name: `Series${series.length + 1}` }]); setSel(series.length); }}>Add</button>
            <button className="xl-btn !min-w-0" disabled={series.length <= 1} onClick={() => { setSeries(series.filter((_, i) => i !== sel)); setSel(0); }}>Remove</button>
          </div>
          <div className="xl-list h-[120px]">
            {series.map((s, i) => <div key={i} className={i === sel ? 'xl-list-sel' : ''} onClick={() => setSel(i)}>{s.name ?? s.nameRef ?? `Series${i + 1}`}</div>)}
          </div>
          {cur && (
            <div className="mt-2 flex flex-col gap-1">
              <label className="flex items-center gap-2"><span className="w-20">Name:</span><input className="xl-input flex-1" value={cur.name ?? ''} onChange={(e) => upd({ name: e.target.value, nameRef: undefined })} /></label>
              <label className="flex items-center gap-2"><span className="w-20">Values:</span><RefInput value={'=' + cur.values} onChange={(v) => upd({ values: v.replace(/^=/, '') })} width={220} /></label>
              {ch.type === 'combo' && (
                <label className="flex items-center gap-2"><span className="w-20">Type:</span>
                  <select className="xl-input" value={cur.seriesType ?? 'column'} onChange={(e) => upd({ seriesType: e.target.value as ChartSeries['seriesType'] })}><option value="column">Column</option><option value="line">Line</option><option value="area">Area</option></select>
                  <label className="flex items-center gap-1"><input type="checkbox" checked={!!cur.secondaryAxis} onChange={(e) => upd({ secondaryAxis: e.target.checked })} /> Secondary axis</label>
                </label>
              )}
            </div>
          )}
        </div>
        <div className="w-[220px]">
          <b>Horizontal (Category) Axis Labels</b>
          <div className="mt-2"><RefInput value={cats ? '=' + cats : ''} onChange={(v) => setCats(v.replace(/^=/, ''))} width={210} /></div>
        </div>
      </div>
    </Dialog>
  );
}

export function ChartFormatDialog({ props }: { props: Record<string, unknown> }) {
  const ch = S().wb.activeSheet.charts.find((c) => c.id === props.id);
  const [s, setS] = useState<ChartSpec | undefined>(ch);
  if (!ch || !s) return null;
  const set = (p: Partial<ChartSpec>) => setS({ ...s, ...p });
  return (
    <Dialog title="Format Chart" width={420} onOk={() => updateChart(ch.id, { title: s.title, showTitle: s.showTitle, legend: s.legend, dataLabels: s.dataLabels, gridlines: s.gridlines, xTitle: s.xTitle, yTitle: s.yTitle, style: s.style, series: s.series })} testId="chart-format">
      <label className="flex items-center gap-2 my-1"><input type="checkbox" checked={s.showTitle} onChange={(e) => set({ showTitle: e.target.checked })} /> Chart title <input className="xl-input flex-1" value={s.title ?? ''} onChange={(e) => set({ title: e.target.value })} data-testid="chart-title" /></label>
      <label className="flex items-center gap-2 my-1"><span className="w-28">Horizontal axis title</span><input className="xl-input flex-1" value={s.xTitle ?? ''} onChange={(e) => set({ xTitle: e.target.value || undefined })} /></label>
      <label className="flex items-center gap-2 my-1"><span className="w-28">Vertical axis title</span><input className="xl-input flex-1" value={s.yTitle ?? ''} onChange={(e) => set({ yTitle: e.target.value || undefined })} /></label>
      <label className="flex items-center gap-2 my-1"><span className="w-28">Legend</span>
        <select className="xl-input" value={s.legend} onChange={(e) => set({ legend: e.target.value as ChartSpec['legend'] })}>
          {['right', 'bottom', 'top', 'left', 'none'].map((l) => <option key={l} value={l}>{l[0].toUpperCase() + l.slice(1)}</option>)}
        </select>
      </label>
      <label className="flex items-center gap-2 my-1"><input type="checkbox" checked={s.dataLabels} onChange={(e) => set({ dataLabels: e.target.checked })} /> Data labels</label>
      <label className="flex items-center gap-2 my-1"><input type="checkbox" checked={s.gridlines} onChange={(e) => set({ gridlines: e.target.checked })} /> Major gridlines</label>
      <div className="mt-2 mb-1">Series colors</div>
      <div className="flex flex-col gap-1">
        {s.series.map((se, i) => (
          <label key={i} className="flex items-center gap-2">
            <input type="color" value={se.color ?? PALETTES[(s.style - 1) % PALETTES.length][i % 6]} onChange={(e) => set({ series: s.series.map((x, k) => (k === i ? { ...x, color: e.target.value.toUpperCase() } : x)) })} />
            {se.name ?? se.nameRef ?? `Series${i + 1}`}
          </label>
        ))}
      </div>
    </Dialog>
  );
}

export function ChartDesignTab() {
  useStore((s) => s.rev);
  useStore((s) => s.selectedChartId);
  const ch = selectedChart();
  if (!ch) return null;
  const up = (p: Partial<ChartSpec>, label = 'Chart') => updateChart(ch.id, p, label);
  return (
    <>
      <Group label="Chart Layouts">
        <LargeButton
          icon={<PlusSquare size={26} className="text-[#2B7CD3]" />}
          label={'Add Chart\nElement'}
          menu={[
            { label: 'Chart Title', submenu: [{ label: 'None', onClick: () => up({ showTitle: false }) }, { label: 'Above Chart', onClick: () => up({ showTitle: true, title: ch.title ?? 'Chart Title' }) }] },
            { label: 'Data Labels', submenu: [{ label: 'None', onClick: () => up({ dataLabels: false }) }, { label: 'Show', onClick: () => up({ dataLabels: true }) }] },
            { label: 'Gridlines', submenu: [{ label: 'Primary Major Horizontal', checked: ch.gridlines, onClick: () => up({ gridlines: !ch.gridlines }) }] },
            { label: 'Legend', submenu: (['none', 'right', 'top', 'left', 'bottom'] as const).map((l) => ({ label: l[0].toUpperCase() + l.slice(1), checked: ch.legend === l, onClick: () => up({ legend: l }) })) },
            { label: 'Axis Titles...', onClick: () => openDialog('chartFormat', { id: ch.id }) },
          ]}
        />
        <LargeButton
          icon={<ChartColumnIcon />}
          label={'Quick\nLayout'}
          menu={[
            { label: 'Layout 1 (title, legend right)', onClick: () => up({ showTitle: true, legend: 'right', dataLabels: false }) },
            { label: 'Layout 2 (title, legend top, labels)', onClick: () => up({ showTitle: true, legend: 'top', dataLabels: true, gridlines: false }) },
            { label: 'Layout 3 (title, legend bottom)', onClick: () => up({ showTitle: true, legend: 'bottom', dataLabels: false }) },
            { label: 'Layout 4 (no title, labels)', onClick: () => up({ showTitle: false, legend: 'bottom', dataLabels: true }) },
          ]}
        />
      </Group>
      <Group label="Chart Styles">
        <LargeButton
          icon={<Palette size={26} className="text-[#ED7D31]" />}
          label={'Change\nColors'}
          panel={(close) => (
            <div className="p-2 w-[220px] flex flex-col gap-1">
              {PALETTES.map((p, i) => (
                <button key={i} className="flex gap-0.5 p-1 rounded hover:bg-[var(--hover)]" onClick={() => { up({ style: i + 1, colors: undefined, series: ch.series.map((s) => ({ ...s, color: undefined })) }, 'Chart Colors'); close(); }}>
                  {p.slice(0, 6).map((c) => <span key={c} className="w-6 h-4" style={{ background: c }} />)}
                </button>
              ))}
            </div>
          )}
        />
        <div className="flex gap-1 items-center px-1">
          {[1, 2, 3, 4, 5, 6].map((st) => (
            <button key={st} className={'w-11 h-14 border rounded overflow-hidden ' + (ch.style === st ? 'xl-item-active' : '')} style={{ borderColor: 'var(--border)' }} onClick={() => up({ style: st, series: ch.series.map((s) => ({ ...s, color: undefined })) }, 'Chart Style')} title={`Style ${st}`}>
              <div className="flex items-end justify-center gap-[2px] h-full pb-1 bg-white">
                {[0.5, 0.9, 0.7].map((hh, k) => <span key={k} className="w-2" style={{ height: `${hh * 36}px`, background: PALETTES[st - 1][k] }} />)}
              </div>
            </button>
          ))}
        </div>
      </Group>
      <Group label="Data">
        <LargeButton icon={<ArrowLeftRight size={26} className="text-[#2B7CD3]" />} label={'Switch Row/\nColumn'} onClick={() => switchRowCol(ch.id)} />
        <LargeButton icon={<Database size={26} className="text-[#2B7CD3]" />} label={'Select\nData'} onClick={() => openDialog('chartData', { id: ch.id })} />
      </Group>
      <Group label="Type">
        <LargeButton icon={<ChartColumnIcon />} label={'Change\nChart Type'} onClick={() => openDialog('insertChart', { changeId: ch.id })} />
      </Group>
      <Group label="Format">
        <Col>
          <SmallButton icon={<Type size={14} />} label="Format Chart..." onClick={() => openDialog('chartFormat', { id: ch.id })} />
          <SmallButton icon={<Trash2 size={14} />} label="Delete Chart" onClick={() => deleteChart(ch.id)} />
          <Row><span className="text-[11px] opacity-70 px-1">Size: {Math.round(ch.anchor.w)}×{Math.round(ch.anchor.h)} px</span></Row>
        </Col>
      </Group>
    </>
  );
}

export { setState };

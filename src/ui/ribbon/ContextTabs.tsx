'use client';
import { CircleDot, Eye, List, Palette, RefreshCw, Table2, Trash2 } from 'lucide-react';
import type { SparklineGroup } from '../../model/types';
import { deletePivot, findPivot, pivotAt, refreshAllPivots, refreshPivot } from '../../state/pivot';
import { openDialog, S, setState, transact, useStore } from '../../state/store';
import { Col, ColorPalette, Group, LargeButton, SmallButton } from './parts';

export function activePivotId(): string | null {
  const st = S();
  const p = pivotAt(st.wb.activeSheet, st.sel.active.r, st.sel.active.c);
  return p?.id ?? null;
}

export function PivotAnalyzeTab() {
  useStore((s) => s.rev);
  useStore((s) => s.sel);
  const id = activePivotId();
  const found = id ? findPivot(id) : null;
  if (!found) return null;
  const src = S().wb.sheetById(found.spec.sourceSheetId);
  return (
    <>
      <Group label="PivotTable">
        <Col className="pt-1">
          <span className="font-semibold">{found.spec.name}</span>
          <span className="opacity-70 text-[11px]">Source: {src?.name ?? '?'}</span>
        </Col>
      </Group>
      <Group label="Data">
        <LargeButton
          icon={<RefreshCw size={26} className="text-[#107C41]" />}
          label="Refresh"
          onClick={() => refreshPivot(found.spec.id)}
          menu={[
            { label: 'Refresh', onClick: () => refreshPivot(found.spec.id) },
            { label: 'Refresh All', onClick: refreshAllPivots },
          ]}
        />
      </Group>
      <Group label="Calculations">
        <LargeButton icon={<Table2 size={26} className="text-[#2B7CD3]" />} label={'Calculated\nField...'} onClick={() => openDialog('pivotCalcField', { id: found.spec.id })} />
      </Group>
      <Group label="Show">
        <LargeButton icon={<List size={26} className="text-[#2B7CD3]" />} label={'Field\nList'} checked={S().pivotPanel === found.spec.id} onClick={() => setState({ pivotPanel: S().pivotPanel ? null : found.spec.id })} />
      </Group>
      <Group label="Actions">
        <LargeButton icon={<Trash2 size={26} className="text-[#C00000]" />} label={'Clear\nPivotTable'} onClick={() => deletePivot(found.spec.id)} />
      </Group>
    </>
  );
}

export function sparklineGroupAtActive(): SparklineGroup | undefined {
  const st = S();
  const a = st.sel.active;
  return st.wb.activeSheet.sparklines.find((g) => g.items.some((i) => i.r === a.r && i.c === a.c));
}

export function SparklineTab() {
  useStore((s) => s.rev);
  useStore((s) => s.sel);
  const g = sparklineGroupAtActive();
  if (!g) return null;
  const sheet = S().wb.activeSheet;
  const up = (patch: Partial<SparklineGroup>, label = 'Sparkline') =>
    transact(label, (tx) => tx.setMeta(sheet, 'sparklines', sheet.sparklines.map((x) => (x.id === g.id ? { ...x, ...patch } : x))));
  return (
    <>
      <Group label="Type">
        {(['line', 'column', 'winloss'] as const).map((t) => (
          <LargeButton key={t} icon={<Eye size={24} className="text-[#2B7CD3]" />} label={t === 'winloss' ? 'Win/\nLoss' : t[0].toUpperCase() + t.slice(1)} checked={g.type === t} onClick={() => up({ type: t })} />
        ))}
      </Group>
      <Group label="Show">
        <Col className="pt-1">
          <label className="flex gap-1"><input type="checkbox" checked={!!g.highPoint} onChange={(e) => up({ highPoint: e.target.checked })} /> High Point</label>
          <label className="flex gap-1"><input type="checkbox" checked={!!g.lowPoint} onChange={(e) => up({ lowPoint: e.target.checked })} /> Low Point</label>
          <label className="flex gap-1"><input type="checkbox" checked={!!g.markers} onChange={(e) => up({ markers: e.target.checked })} /> Markers</label>
        </Col>
      </Group>
      <Group label="Style">
        <Col>
          <SmallButton icon={<Palette size={14} />} label="Sparkline Color" panel={(close) => <ColorPalette close={close} onPick={(c) => c && up({ color: c })} />} />
          <SmallButton icon={<CircleDot size={14} />} label="Negative Color" panel={(close) => <ColorPalette close={close} onPick={(c) => c && up({ negColor: c })} />} />
        </Col>
      </Group>
      <Group label="Group">
        <LargeButton icon={<Trash2 size={24} className="text-[#C00000]" />} label="Clear" onClick={() => transact('Clear Sparklines', (tx) => tx.setMeta(sheet, 'sparklines', sheet.sparklines.filter((x) => x.id !== g.id)))} />
      </Group>
    </>
  );
}

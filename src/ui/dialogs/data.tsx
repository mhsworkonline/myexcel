'use client';
import { ChevronDown, ChevronUp } from 'lucide-react';
import { useMemo, useState } from 'react';
import { isErrorVal } from '../../engine/Engine';
import { colToName, parseRange, quoteSheetName, Range, rangeToA1 } from '../../model/address';
import { tokenize } from '../../model/formula';
import { primaryRange } from '../../model/selection';
import type { CFOperator, CFRule, CFStyle, CFType, CFVO, DataValidation, DefinedName, DVType } from '../../model/types';
import {
  addCFRule,
  CF_PRESETS,
  COLOR_SCALES,
  createTable,
  DATA_BAR_COLORS,
  DEFAULT_ALLOW,
  ICON_SETS,
  protectSheet,
  removeDuplicates,
  selectionRanges,
  setCFRules,
  setValidation,
  splitText,
  tableSourceGuess,
  textToColumns,
  TextToColumnsOpts,
  unprotectSheet,
} from '../../state/actions/data';
import { createNamesFromSelection, defineName, deleteName, selectionRefText, setNames } from '../../state/actions/formulas';
import { guessHeader, displayTextAt } from '../../state/actions/sortFilter';
import { resolveReference } from '../../state/actions/find';
import { alertBox, closeDialog, openDialog, S, setState, useStore } from '../../state/store';
import { validationAt } from '../../state/validation';
import { getComputed } from '../../state/values';
import { IconSetPreview } from '../ribbon/HomeTab';
import { Dialog, RefInput } from './Dialog';

// ---------- conditional formatting ----------

function StylePresetSelect({ value, onChange }: { value: string; onChange: (v: string) => void }) {
  return (
    <select className="xl-input" value={value} onChange={(e) => onChange(e.target.value)} data-testid="cf-preset">
      {Object.keys(CF_PRESETS).map((k) => (
        <option key={k}>{k}</option>
      ))}
      <option value="custom">Custom Format...</option>
    </select>
  );
}

export function CFQuickDialog({ props }: { props: Record<string, unknown> }) {
  const kind = props.kind as string;
  const [v1, setV1] = useState('');
  const [v2, setV2] = useState('');
  const [preset, setPreset] = useState('Light Red Fill with Dark Red Text');
  const [period, setPeriod] = useState<NonNullable<CFRule['timePeriod']>>('yesterday');
  const [dup, setDup] = useState<'duplicateValues' | 'uniqueValues'>('duplicateValues');
  const [n, setN] = useState(10);
  const titles: Record<string, [string, string]> = {
    greaterThan: ['Greater Than', 'Format cells that are GREATER THAN:'],
    lessThan: ['Less Than', 'Format cells that are LESS THAN:'],
    between: ['Between', 'Format cells that are BETWEEN:'],
    equal: ['Equal To', 'Format cells that are EQUAL TO:'],
    containsText: ['Text That Contains', 'Format cells that contain the text:'],
    timePeriod: ['A Date Occurring', 'Format cells that contain a date occurring:'],
    duplicateValues: ['Duplicate Values', 'Format cells that contain:'],
    top10: ['Top 10 Items', 'Format cells that rank in the TOP:'],
    top10pct: ['Top 10%', 'Format cells that rank in the TOP:'],
    bottom10: ['Bottom 10 Items', 'Format cells that rank in the BOTTOM:'],
    bottom10pct: ['Bottom 10%', 'Format cells that rank in the BOTTOM:'],
    aboveAverage: ['Above Average', 'Format cells that are ABOVE AVERAGE:'],
    belowAverage: ['Below Average', 'Format cells that are BELOW AVERAGE:'],
  };
  const [title, prompt] = titles[kind] ?? ['Rule', ''];
  const style: CFStyle = CF_PRESETS[preset] ?? CF_PRESETS['Light Red Fill with Dark Red Text'];
  const q = (v: string) => (v.startsWith('=') ? v : isNaN(Number(v)) || v === '' ? `="${v}"` : '=' + v);
  const ok = () => {
    switch (kind) {
      case 'greaterThan':
      case 'lessThan':
      case 'equal':
        addCFRule({ type: 'cellIs', operator: kind as CFOperator, formulas: [q(v1)], style });
        break;
      case 'between':
        addCFRule({ type: 'cellIs', operator: 'between', formulas: [q(v1), q(v2)], style });
        break;
      case 'containsText':
        addCFRule({ type: 'containsText', text: v1, style });
        break;
      case 'timePeriod':
        addCFRule({ type: 'timePeriod', timePeriod: period, style });
        break;
      case 'duplicateValues':
        addCFRule({ type: dup, style });
        break;
      case 'top10':
      case 'top10pct':
      case 'bottom10':
      case 'bottom10pct':
        addCFRule({ type: 'top10', rank: n, percent: kind.endsWith('pct'), bottom: kind.startsWith('bottom'), style });
        break;
      case 'aboveAverage':
      case 'belowAverage':
        addCFRule({ type: 'aboveAverage', aboveAverage: kind === 'aboveAverage', style });
        break;
    }
  };
  return (
    <Dialog title={title} width={480} onOk={ok} testId="cf-quick">
      <div className="mb-2">{prompt}</div>
      <div className="flex items-center gap-2">
        {['greaterThan', 'lessThan', 'equal', 'containsText', 'between'].includes(kind) && <RefInput value={v1} onChange={setV1} width={170} testId="cf-v1" />}
        {kind === 'between' && <span>and</span>}
        {kind === 'between' && <RefInput value={v2} onChange={setV2} width={140} />}
        {kind === 'timePeriod' && (
          <select className="xl-input" value={period} onChange={(e) => setPeriod(e.target.value as typeof period)}>
            {['yesterday', 'today', 'tomorrow', 'last7Days', 'lastWeek', 'thisWeek', 'nextWeek', 'lastMonth', 'thisMonth', 'nextMonth'].map((p) => <option key={p} value={p}>{p.replace(/([A-Z0-9])/g, ' $1').replace(/^./, (x) => x.toUpperCase())}</option>)}
          </select>
        )}
        {kind === 'duplicateValues' && (
          <select className="xl-input" value={dup} onChange={(e) => setDup(e.target.value as typeof dup)}>
            <option value="duplicateValues">Duplicate</option>
            <option value="uniqueValues">Unique</option>
          </select>
        )}
        {kind.startsWith('top') || kind.startsWith('bottom') ? <><input type="number" className="xl-input w-16" value={n} onChange={(e) => setN(Math.max(1, +e.target.value))} />{kind.endsWith('pct') && '%'}</> : null}
        <span>with</span>
        <StylePresetSelect value={preset} onChange={(v) => (v === 'custom' ? openDialog('cfRule', { type: kind }) : setPreset(v))} />
      </div>
    </Dialog>
  );
}

const RULE_KINDS: [string, string][] = [
  ['colorScale', 'Format all cells based on their values'],
  ['cellIs', 'Format only cells that contain'],
  ['top10', 'Format only top or bottom ranked values'],
  ['aboveAverage', 'Format only values that are above or below average'],
  ['duplicateValues', 'Format only unique or duplicate values'],
  ['expression', 'Use a formula to determine which cells to format'],
];

export function CFRuleDialog({ props }: { props: Record<string, unknown> }) {
  const editing = props.rule as CFRule | undefined;
  const initialKind = editing ? (['dataBar', 'iconSet', 'colorScale'].includes(editing.type) ? 'colorScale' : ['containsText', 'notContainsText', 'beginsWith', 'endsWith', 'containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors', 'timePeriod'].includes(editing.type) ? 'cellIs' : editing.type === 'uniqueValues' ? 'duplicateValues' : editing.type) : props.type === 'dataBar' || props.type === 'iconSet' ? 'colorScale' : (props.type as string) ?? 'colorScale';
  const [kind, setKind] = useState(RULE_KINDS.some(([k]) => k === initialKind) ? initialKind : 'cellIs');
  const [formatStyle, setFormatStyle] = useState<'2color' | '3color' | 'dataBar' | 'iconSet'>(editing?.type === 'dataBar' || props.type === 'dataBar' ? 'dataBar' : editing?.type === 'iconSet' || props.type === 'iconSet' ? 'iconSet' : editing?.colorScale?.colors.length === 2 ? '2color' : '3color');
  const [style, setStyle] = useState<CFStyle>(editing?.style ?? { fillColor: '#FFC7CE', fontColor: '#9C0006' });
  // cellIs
  const [what, setWhat] = useState<string>(editing && editing.type !== 'cellIs' && editing.type !== 'expression' ? editing.type : 'cellValue');
  const [op, setOp] = useState<CFOperator>(editing?.operator ?? 'between');
  const [f1, setF1] = useState(editing?.formulas?.[0]?.replace(/^=/, '') ?? editing?.text ?? '');
  const [f2, setF2] = useState(editing?.formulas?.[1]?.replace(/^=/, '') ?? '');
  // top10
  const [bottom, setBottom] = useState(!!editing?.bottom);
  const [rank, setRank] = useState(editing?.rank ?? 10);
  const [pct, setPct] = useState(!!editing?.percent);
  const [above, setAbove] = useState(editing?.aboveAverage !== false);
  const [unique, setUnique] = useState(editing?.type === 'uniqueValues');
  // scales
  const [colors, setColors] = useState<string[]>(editing?.colorScale?.colors ?? ['#F8696B', '#FFEB84', '#63BE7B']);
  const [minT, setMinT] = useState<CFVO>(editing?.colorScale?.cfvos[0] ?? editing?.dataBar?.min ?? { type: 'min' });
  const [maxT, setMaxT] = useState<CFVO>(editing?.colorScale?.cfvos[editing.colorScale.cfvos.length - 1] ?? editing?.dataBar?.max ?? { type: 'max' });
  const [barColor, setBarColor] = useState(editing?.dataBar?.color ?? DATA_BAR_COLORS[0]);
  const [gradient, setGradient] = useState(editing?.dataBar?.gradient ?? true);
  const [iconSet, setIconSet] = useState(editing?.iconSet?.set ?? '3TrafficLights1');
  const [reverse, setReverse] = useState(!!editing?.iconSet?.reverse);
  const [showValue, setShowValue] = useState(editing?.iconSet?.showValue ?? editing?.dataBar?.showValue ?? true);
  const [stopIfTrue, setStop] = useState(!!editing?.stopIfTrue);

  const build = (): Omit<CFRule, 'id' | 'priority' | 'ranges'> | null => {
    const eq = (v: string) => (v.startsWith('=') ? v : isNaN(Number(v)) || v === '' ? `="${v}"` : '=' + v);
    switch (kind) {
      case 'colorScale':
        if (formatStyle === 'dataBar') return { type: 'dataBar', dataBar: { color: barColor, gradient, min: minT.type === 'min' ? { type: 'autoMin' } : minT, max: maxT.type === 'max' ? { type: 'autoMax' } : maxT, showValue } };
        if (formatStyle === 'iconSet') {
          const n = parseInt(iconSet, 10) || 3;
          return { type: 'iconSet', iconSet: { set: iconSet, cfvos: Array.from({ length: n }, (_, i) => ({ type: 'percent' as const, value: Math.round((i * 100) / n) })), reverse, showValue } };
        }
        return {
          type: 'colorScale',
          colorScale: formatStyle === '2color' ? { cfvos: [minT, maxT], colors: [colors[0], colors[colors.length - 1]] } : { cfvos: [minT, { type: 'percentile', value: 50 }, maxT], colors: colors.length === 3 ? colors : [colors[0], '#FFEB84', colors[colors.length - 1]] },
        };
      case 'cellIs':
        if (what === 'cellValue') return { type: 'cellIs', operator: op, formulas: op === 'between' || op === 'notBetween' ? [eq(f1), eq(f2)] : [eq(f1)], style, stopIfTrue };
        if (['containsText', 'notContainsText', 'beginsWith', 'endsWith'].includes(what)) return { type: what as CFType, text: f1, style, stopIfTrue };
        return { type: what as CFType, style, stopIfTrue };
      case 'top10':
        return { type: 'top10', rank, percent: pct, bottom, style, stopIfTrue };
      case 'aboveAverage':
        return { type: 'aboveAverage', aboveAverage: above, style, stopIfTrue };
      case 'duplicateValues':
        return { type: unique ? 'uniqueValues' : 'duplicateValues', style, stopIfTrue };
      case 'expression':
        if (!f1) return null;
        return { type: 'expression', formulas: [f1.startsWith('=') ? f1 : '=' + f1], style, stopIfTrue };
    }
    return null;
  };
  const ok = () => {
    const r = build();
    if (!r) {
      alertBox('Enter a formula.');
      return false;
    }
    if (editing && props.onSave) (props.onSave as (r: CFRule) => void)({ ...editing, ...r, id: editing.id, priority: editing.priority, ranges: editing.ranges } as CFRule);
    else addCFRule(r);
  };
  const vo = (v: CFVO, set: (v: CFVO) => void, isMin: boolean) => (
    <div className="flex flex-col gap-1">
      <select className="xl-input" value={v.type} onChange={(e) => set({ type: e.target.value as CFVO['type'], value: v.value })}>
        <option value={isMin ? 'min' : 'max'}>{isMin ? 'Lowest Value' : 'Highest Value'}</option>
        <option value="num">Number</option>
        <option value="percent">Percent</option>
        <option value="percentile">Percentile</option>
        <option value="formula">Formula</option>
      </select>
      <input className="xl-input" disabled={v.type === 'min' || v.type === 'max'} value={v.value ?? ''} onChange={(e) => set({ ...v, value: e.target.value })} />
    </div>
  );
  return (
    <Dialog title={editing ? 'Edit Formatting Rule' : 'New Formatting Rule'} width={560} onOk={ok} testId="cf-rule">
      <div className="mb-1">Select a Rule Type:</div>
      <div className="xl-list h-[100px] mb-3">
        {RULE_KINDS.map(([k, l]) => (
          <div key={k} className={kind === k ? 'xl-list-sel' : ''} onClick={() => setKind(k)}>► {l}</div>
        ))}
      </div>
      <div className="mb-1">Edit the Rule Description:</div>
      <div className="border rounded p-3 min-h-[180px]" style={{ borderColor: 'var(--border)' }}>
        {kind === 'colorScale' && (
          <div className="flex flex-col gap-2">
            <label className="flex items-center gap-2">
              Format Style:
              <select className="xl-input" value={formatStyle} onChange={(e) => setFormatStyle(e.target.value as typeof formatStyle)}>
                <option value="2color">2-Color Scale</option>
                <option value="3color">3-Color Scale</option>
                <option value="dataBar">Data Bar</option>
                <option value="iconSet">Icon Sets</option>
              </select>
            </label>
            {(formatStyle === '2color' || formatStyle === '3color' || formatStyle === 'dataBar') && (
              <div className="grid grid-cols-2 gap-4">
                <div>Minimum{vo(minT, setMinT, true)}</div>
                <div>Maximum{vo(maxT, setMaxT, false)}</div>
              </div>
            )}
            {(formatStyle === '2color' || formatStyle === '3color') && (
              <div className="flex gap-3 items-center">
                Colors:
                {(formatStyle === '2color' ? [0, colors.length - 1] : [0, 1, 2]).map((i) => (
                  <input key={i} type="color" value={colors[i] ?? '#FFFFFF'} onChange={(e) => { const c = [...colors]; c[i] = e.target.value.toUpperCase(); setColors(c); }} />
                ))}
                <select className="xl-input" onChange={(e) => { const s = COLOR_SCALES.find((x) => x.name === e.target.value); if (s) { setColors(s.colors); setFormatStyle(s.colors.length === 2 ? '2color' : '3color'); } }}>
                  <option>Presets…</option>
                  {COLOR_SCALES.map((s) => <option key={s.name}>{s.name}</option>)}
                </select>
              </div>
            )}
            {formatStyle === 'dataBar' && (
              <div className="flex gap-3 items-center">
                Bar color: <input type="color" value={barColor} onChange={(e) => setBarColor(e.target.value.toUpperCase())} />
                <select className="xl-input" value={gradient ? 'g' : 's'} onChange={(e) => setGradient(e.target.value === 'g')}><option value="g">Gradient Fill</option><option value="s">Solid Fill</option></select>
                <label className="flex gap-1"><input type="checkbox" checked={!showValue} onChange={(e) => setShowValue(!e.target.checked)} /> Show Bar Only</label>
              </div>
            )}
            {formatStyle === 'iconSet' && (
              <div className="flex flex-col gap-2">
                <div className="flex gap-2 items-center">
                  Icon Style:
                  <select className="xl-input" value={iconSet} onChange={(e) => setIconSet(e.target.value)}>
                    {ICON_SETS.map((s) => <option key={s}>{s}</option>)}
                  </select>
                  <IconSetPreview set={iconSet} />
                </div>
                <label className="flex gap-1"><input type="checkbox" checked={reverse} onChange={(e) => setReverse(e.target.checked)} /> Reverse Icon Order</label>
                <label className="flex gap-1"><input type="checkbox" checked={!showValue} onChange={(e) => setShowValue(!e.target.checked)} /> Show Icon Only</label>
                <div className="text-[11px] opacity-70">Icons are assigned by equal percent bands of the value range.</div>
              </div>
            )}
          </div>
        )}
        {kind === 'cellIs' && (
          <div className="flex flex-col gap-2">
            <div>Format only cells with:</div>
            <div className="flex gap-2 flex-wrap">
              <select className="xl-input" value={what} onChange={(e) => setWhat(e.target.value)}>
                <option value="cellValue">Cell Value</option>
                <option value="containsText">Specific Text: containing</option>
                <option value="notContainsText">Specific Text: not containing</option>
                <option value="beginsWith">Specific Text: beginning with</option>
                <option value="endsWith">Specific Text: ending with</option>
                <option value="containsBlanks">Blanks</option>
                <option value="notContainsBlanks">No Blanks</option>
                <option value="containsErrors">Errors</option>
                <option value="notContainsErrors">No Errors</option>
              </select>
              {what === 'cellValue' && (
                <select className="xl-input" value={op} onChange={(e) => setOp(e.target.value as CFOperator)}>
                  {(['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'] as CFOperator[]).map((o) => (
                    <option key={o} value={o}>{o.replace(/([A-Z])/g, ' $1').toLowerCase()}</option>
                  ))}
                </select>
              )}
              {(what === 'cellValue' || what.includes('Text') || what === 'beginsWith' || what === 'endsWith') && <input className="xl-input w-28" value={f1} onChange={(e) => setF1(e.target.value)} />}
              {what === 'cellValue' && (op === 'between' || op === 'notBetween') && <><span>and</span><input className="xl-input w-28" value={f2} onChange={(e) => setF2(e.target.value)} /></>}
            </div>
          </div>
        )}
        {kind === 'top10' && (
          <div className="flex gap-2 items-center">
            Format values that rank in the:
            <select className="xl-input" value={bottom ? 'b' : 't'} onChange={(e) => setBottom(e.target.value === 'b')}><option value="t">Top</option><option value="b">Bottom</option></select>
            <input type="number" className="xl-input w-16" value={rank} onChange={(e) => setRank(+e.target.value)} />
            <label className="flex gap-1"><input type="checkbox" checked={pct} onChange={(e) => setPct(e.target.checked)} /> % of the selected range</label>
          </div>
        )}
        {kind === 'aboveAverage' && (
          <div className="flex gap-2 items-center">
            Format values that are:
            <select className="xl-input" value={above ? 'a' : 'b'} onChange={(e) => setAbove(e.target.value === 'a')}><option value="a">above</option><option value="b">below</option></select>
            the average for the selected range
          </div>
        )}
        {kind === 'duplicateValues' && (
          <div className="flex gap-2 items-center">
            Format all:
            <select className="xl-input" value={unique ? 'u' : 'd'} onChange={(e) => setUnique(e.target.value === 'u')}><option value="d">duplicate</option><option value="u">unique</option></select>
            values in the selected range
          </div>
        )}
        {kind === 'expression' && (
          <div className="flex flex-col gap-2">
            <div>Format values where this formula is true:</div>
            <RefInput value={f1} onChange={setF1} width={480} testId="cf-formula" />
          </div>
        )}
        {kind !== 'colorScale' && (
          <div className="flex items-center gap-3 mt-4">
            <span>Preview:</span>
            <div className="flex-1 h-7 border flex items-center justify-center" style={{ background: style.fillColor ?? '#fff', color: style.fontColor ?? '#000', fontWeight: style.bold ? 700 : 400, fontStyle: style.italic ? 'italic' : 'normal', textDecoration: [style.underline ? 'underline' : '', style.strike ? 'line-through' : ''].join(' '), borderColor: style.borderColor ?? 'var(--border)' }}>
              AaBbCcYyZz
            </div>
            <CFFormatButton style={style} onChange={setStyle} />
          </div>
        )}
        {kind !== 'colorScale' && <label className="flex gap-1 mt-2"><input type="checkbox" checked={stopIfTrue} onChange={(e) => setStop(e.target.checked)} /> Stop If True</label>}
      </div>
    </Dialog>
  );
}

function CFFormatButton({ style, onChange }: { style: CFStyle; onChange: (s: CFStyle) => void }) {
  const [open, setOpen] = useState(false);
  return (
    <div className="relative">
      <button className="xl-btn" onClick={() => setOpen(!open)}>Format...</button>
      {open && (
        <div className="xl-popup absolute right-0 bottom-9 z-10 p-3 w-[250px] flex flex-col gap-2 rounded">
          <label className="flex items-center justify-between">Font color <input type="color" value={style.fontColor ?? '#000000'} onChange={(e) => onChange({ ...style, fontColor: e.target.value.toUpperCase() })} /></label>
          <label className="flex items-center justify-between">Fill color <input type="color" value={style.fillColor ?? '#FFFFFF'} onChange={(e) => onChange({ ...style, fillColor: e.target.value.toUpperCase() })} /></label>
          <label className="flex items-center justify-between">Border color <input type="color" value={style.borderColor ?? '#000000'} onChange={(e) => onChange({ ...style, borderColor: e.target.value.toUpperCase() })} /></label>
          <div className="flex gap-3">
            <label className="flex gap-1"><input type="checkbox" checked={!!style.bold} onChange={(e) => onChange({ ...style, bold: e.target.checked })} /> <b>B</b></label>
            <label className="flex gap-1"><input type="checkbox" checked={!!style.italic} onChange={(e) => onChange({ ...style, italic: e.target.checked })} /> <i>I</i></label>
            <label className="flex gap-1"><input type="checkbox" checked={!!style.underline} onChange={(e) => onChange({ ...style, underline: e.target.checked })} /> <u>U</u></label>
            <label className="flex gap-1"><input type="checkbox" checked={!!style.strike} onChange={(e) => onChange({ ...style, strike: e.target.checked })} /> <s>S</s></label>
          </div>
          <label className="flex items-center gap-2">Number format <input className="xl-input flex-1" value={style.numFmt ?? ''} onChange={(e) => onChange({ ...style, numFmt: e.target.value || undefined })} /></label>
          <div className="flex justify-between">
            <button className="xl-btn !min-w-0" onClick={() => onChange({})}>Clear</button>
            <button className="xl-btn xl-btn-primary !min-w-0" onClick={() => setOpen(false)}>Done</button>
          </div>
        </div>
      )}
    </div>
  );
}

function describeRule(r: CFRule): string {
  switch (r.type) {
    case 'cellIs': return `Cell Value ${(r.operator ?? '').replace(/([A-Z])/g, ' $1').toLowerCase()} ${(r.formulas ?? []).map((f) => f.replace(/^=/, '')).join(' and ')}`;
    case 'expression': return `Formula: ${r.formulas?.[0] ?? ''}`;
    case 'containsText': return `Cell value contains "${r.text}"`;
    case 'notContainsText': return `Cell value does not contain "${r.text}"`;
    case 'beginsWith': return `Cell value begins with "${r.text}"`;
    case 'endsWith': return `Cell value ends with "${r.text}"`;
    case 'top10': return `${r.bottom ? 'Bottom' : 'Top'} ${r.rank}${r.percent ? '%' : ''}`;
    case 'aboveAverage': return r.aboveAverage === false ? 'Below Average' : 'Above Average';
    case 'duplicateValues': return 'Duplicate Values';
    case 'uniqueValues': return 'Unique Values';
    case 'dataBar': return 'Data Bar';
    case 'colorScale': return 'Graded Color Scale';
    case 'iconSet': return 'Icon Set';
    case 'timePeriod': return `Date: ${r.timePeriod}`;
    default: return r.type.replace(/([A-Z])/g, ' $1');
  }
}

export function CFManagerDialog() {
  const sheet = S().wb.activeSheet;
  const [scope, setScope] = useState<'selection' | 'sheet'>('selection');
  const [rules, setRules] = useState<CFRule[]>(() => [...sheet.conditionalFormats].sort((a, b) => a.priority - b.priority));
  const [sel, setSel] = useState<string | null>(rules[0]?.id ?? null);
  const selRanges = selectionRanges();
  const visible = scope === 'sheet' ? rules : rules.filter((r) => r.ranges.some((rg) => selRanges.some((s) => rg.r1 <= s.r2 && s.r1 <= rg.r2 && rg.c1 <= s.c2 && s.c1 <= rg.c2)));
  const move = (d: number) => {
    const i = rules.findIndex((r) => r.id === sel);
    const j = i + d;
    if (i < 0 || j < 0 || j >= rules.length) return;
    const n = [...rules];
    [n[i], n[j]] = [n[j], n[i]];
    setRules(n);
  };
  return (
    <Dialog
      title="Conditional Formatting Rules Manager"
      width={660}
      okLabel="OK"
      onOk={() => setCFRules(rules.map((r, i) => ({ ...r, priority: i + 1 })))}
      extraButtons={<button className="xl-btn" onClick={() => setCFRules(rules.map((r, i) => ({ ...r, priority: i + 1 })))}>Apply</button>}
      testId="cf-manager"
    >
      <label className="flex items-center gap-2 mb-2">
        Show formatting rules for:
        <select className="xl-input" value={scope} onChange={(e) => setScope(e.target.value as typeof scope)}>
          <option value="selection">Current Selection</option>
          <option value="sheet">This Worksheet</option>
        </select>
      </label>
      <div className="flex gap-1.5 mb-2">
        <button className="xl-btn !min-w-0" onClick={() => { closeDialog(); openDialog('cfRule', {}); }}>New Rule...</button>
        <button className="xl-btn !min-w-0" disabled={!sel} onClick={() => {
          const r = rules.find((x) => x.id === sel);
          if (!r) return;
          openDialog('cfRule', { rule: r, onSave: (nr: CFRule) => { const next = rules.map((x) => (x.id === nr.id ? nr : x)); setCFRules(next.map((x, i) => ({ ...x, priority: i + 1 }))); setTimeout(() => openDialog('cfManager'), 0); } });
        }}>Edit Rule...</button>
        <button className="xl-btn !min-w-0" disabled={!sel} onClick={() => setRules(rules.filter((r) => r.id !== sel))}>Delete Rule</button>
        <button className="xl-btn !min-w-0" onClick={() => move(-1)}><ChevronUp size={13} /></button>
        <button className="xl-btn !min-w-0" onClick={() => move(1)}><ChevronDown size={13} /></button>
      </div>
      <div className="xl-list h-[200px]">
        <div className="grid grid-cols-[1fr_120px_140px_80px] font-semibold sticky top-0" style={{ background: 'var(--chrome-bg)' }}>
          <span>Rule (applied in order shown)</span><span>Format</span><span>Applies to</span><span>Stop If True</span>
        </div>
        {visible.map((r) => (
          <div key={r.id} className={'grid grid-cols-[1fr_120px_140px_80px] items-center ' + (sel === r.id ? 'xl-list-sel' : '')} onClick={() => setSel(r.id)}>
            <span className="truncate">{describeRule(r)}</span>
            <span className="h-4 mx-1 border text-center text-[10px]" style={{ background: r.style?.fillColor ?? r.dataBar?.color ?? r.colorScale?.colors[0] ?? '#fff', color: r.style?.fontColor ?? '#000' }}>AaBbCc</span>
            <input
              className="xl-input !h-[20px] text-[11px]"
              value={'=' + r.ranges.map((rg) => rangeToA1(rg, true)).join(',')}
              onChange={(e) => {
                const ranges = e.target.value.replace(/^=/, '').split(',').map((x) => parseRange(x.replace(/\$/g, ''))).filter((x): x is Range => !!x);
                if (ranges.length) setRules(rules.map((x) => (x.id === r.id ? { ...x, ranges } : x)));
              }}
            />
            <input type="checkbox" className="justify-self-center" checked={!!r.stopIfTrue} onChange={(e) => setRules(rules.map((x) => (x.id === r.id ? { ...x, stopIfTrue: e.target.checked } : x)))} />
          </div>
        ))}
      </div>
    </Dialog>
  );
}

// ---------- data validation ----------

export function DataValidationDialog() {
  const st = S();
  const sheet = st.wb.activeSheet;
  const cur = validationAt(sheet, st.sel.active.r, st.sel.active.c);
  const [tab, setTab] = useState<'Settings' | 'Input Message' | 'Error Alert'>('Settings');
  const [dv, setDv] = useState<Omit<DataValidation, 'id' | 'ranges'>>(
    cur ?? { type: 'any', operator: 'between', allowBlank: true, showDropdown: true, showInput: true, showError: true, errorStyle: 'stop' },
  );
  const upd = (p: Partial<DataValidation>) => setDv({ ...dv, ...p });
  const needsOp = ['whole', 'decimal', 'date', 'time', 'textLength'].includes(dv.type);
  const two = dv.operator === 'between' || dv.operator === 'notBetween';
  const lbl = dv.type === 'date' ? 'Date' : dv.type === 'time' ? 'Time' : dv.type === 'textLength' ? 'Length' : 'Value';
  return (
    <Dialog
      title="Data Validation"
      width={440}
      onOk={() => setValidation(dv.type === 'any' && !dv.prompt ? null : dv)}
      extraButtons={<button className="xl-btn" onClick={() => { setValidation(null); closeDialog(); }}>Clear All</button>}
      testId="dv-dialog"
    >
      <div className="flex border-b mb-3" style={{ borderColor: 'var(--border)' }}>
        {(['Settings', 'Input Message', 'Error Alert'] as const).map((t) => (
          <button key={t} className={'xl-dtab ' + (tab === t ? 'xl-dtab-active' : '')} onClick={() => setTab(t)}>{t}</button>
        ))}
      </div>
      <div className="min-h-[220px]">
        {tab === 'Settings' && (
          <div className="flex flex-col gap-2">
            <div className="font-semibold">Validation criteria</div>
            <div className="flex gap-3">
              <label className="flex flex-col gap-1 flex-1">
                Allow:
                <select className="xl-input" value={dv.type} onChange={(e) => upd({ type: e.target.value as DVType })} data-testid="dv-type">
                  <option value="any">Any value</option>
                  <option value="whole">Whole number</option>
                  <option value="decimal">Decimal</option>
                  <option value="list">List</option>
                  <option value="date">Date</option>
                  <option value="time">Time</option>
                  <option value="textLength">Text length</option>
                  <option value="custom">Custom</option>
                </select>
              </label>
              <div className="flex flex-col gap-1 pt-4">
                <label className="flex gap-1"><input type="checkbox" checked={dv.allowBlank} onChange={(e) => upd({ allowBlank: e.target.checked })} /> Ignore blank</label>
                {dv.type === 'list' && <label className="flex gap-1"><input type="checkbox" checked={dv.showDropdown} onChange={(e) => upd({ showDropdown: e.target.checked })} /> In-cell dropdown</label>}
              </div>
            </div>
            {needsOp && (
              <label className="flex flex-col gap-1">
                Data:
                <select className="xl-input" value={dv.operator ?? 'between'} onChange={(e) => upd({ operator: e.target.value as CFOperator })}>
                  {(['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'] as CFOperator[]).map((o) => (
                    <option key={o} value={o}>{o.replace(/([A-Z])/g, ' $1').toLowerCase()}</option>
                  ))}
                </select>
              </label>
            )}
            {dv.type === 'list' && (
              <label className="flex flex-col gap-1">
                Source:
                <RefInput value={dv.formula1 ?? ''} onChange={(v) => upd({ formula1: v })} width={380} testId="dv-source" />
                <span className="text-[11px] opacity-70">Comma-separated items (Yes,No,Maybe) or a range reference (=$A$1:$A$5).</span>
              </label>
            )}
            {dv.type === 'custom' && (
              <label className="flex flex-col gap-1">
                Formula:
                <RefInput value={dv.formula1 ?? ''} onChange={(v) => upd({ formula1: v.startsWith('=') ? v : '=' + v })} width={380} />
              </label>
            )}
            {needsOp && (
              <>
                <label className="flex flex-col gap-1">
                  {two ? 'Minimum:' : lbl + ':'}
                  <RefInput value={dv.formula1 ?? ''} onChange={(v) => upd({ formula1: v })} width={380} testId="dv-min" />
                </label>
                {two && (
                  <label className="flex flex-col gap-1">
                    Maximum:
                    <RefInput value={dv.formula2 ?? ''} onChange={(v) => upd({ formula2: v })} width={380} testId="dv-max" />
                  </label>
                )}
              </>
            )}
          </div>
        )}
        {tab === 'Input Message' && (
          <div className="flex flex-col gap-2">
            <label className="flex gap-1"><input type="checkbox" checked={dv.showInput} onChange={(e) => upd({ showInput: e.target.checked })} /> Show input message when cell is selected</label>
            <label className="flex flex-col gap-1">Title:<input className="xl-input" value={dv.promptTitle ?? ''} onChange={(e) => upd({ promptTitle: e.target.value })} /></label>
            <label className="flex flex-col gap-1">Input message:<textarea className="xl-input !h-24 py-1" value={dv.prompt ?? ''} onChange={(e) => upd({ prompt: e.target.value })} /></label>
          </div>
        )}
        {tab === 'Error Alert' && (
          <div className="flex flex-col gap-2">
            <label className="flex gap-1"><input type="checkbox" checked={dv.showError} onChange={(e) => upd({ showError: e.target.checked })} /> Show error alert after invalid data is entered</label>
            <label className="flex flex-col gap-1">Style:
              <select className="xl-input" value={dv.errorStyle} onChange={(e) => upd({ errorStyle: e.target.value as DataValidation['errorStyle'] })}>
                <option value="stop">Stop</option><option value="warning">Warning</option><option value="information">Information</option>
              </select>
            </label>
            <label className="flex flex-col gap-1">Title:<input className="xl-input" value={dv.errorTitle ?? ''} onChange={(e) => upd({ errorTitle: e.target.value })} /></label>
            <label className="flex flex-col gap-1">Error message:<textarea className="xl-input !h-20 py-1" value={dv.error ?? ''} onChange={(e) => upd({ error: e.target.value })} /></label>
          </div>
        )}
      </div>
    </Dialog>
  );
}

// ---------- names ----------

export function NameManagerDialog() {
  useStore((s) => s.rev);
  const wb = S().wb;
  const [sel, setSel] = useState<string | null>(wb.names[0]?.name ?? null);
  const [ref, setRef] = useState<string | null>(null);
  const cur = wb.names.find((n) => n.name === sel);
  const valueOf = (n: DefinedName) => {
    const v = S().engine.evaluateArray('=' + n.ref, wb.activeSheet);
    if (!v.length) return '';
    const flat = v.flat().slice(0, 6).map((x) => (isErrorVal(x) ? x.error : String(x ?? '')));
    return flat.length > 1 ? `{${flat.join(';')}${v.flat().length > 6 ? ';…' : ''}}` : flat[0];
  };
  return (
    <Dialog title="Name Manager" width={680} okLabel="Close" cancelLabel={null} testId="name-manager">
      <div className="flex gap-1.5 mb-2">
        <button className="xl-btn !min-w-0" onClick={() => openDialog('newName', { back: 'nameManager' })}>New...</button>
        <button className="xl-btn !min-w-0" disabled={!cur} onClick={() => cur && openDialog('newName', { edit: cur, back: 'nameManager' })}>Edit...</button>
        <button className="xl-btn !min-w-0" disabled={!cur} onClick={() => cur && deleteName(cur.name, cur.scope)}>Delete</button>
      </div>
      <div className="xl-list h-[220px]">
        <div className="grid grid-cols-[130px_160px_1fr_90px_1fr] font-semibold sticky top-0" style={{ background: 'var(--chrome-bg)' }}>
          <span>Name</span><span>Value</span><span>Refers To</span><span>Scope</span><span>Comment</span>
        </div>
        {wb.names.map((n) => (
          <div key={n.name + n.scope} className={'grid grid-cols-[130px_160px_1fr_90px_1fr] ' + (sel === n.name ? 'xl-list-sel' : '')} onClick={() => { setSel(n.name); setRef(null); }}>
            <span className="truncate">{n.name}</span>
            <span className="truncate">{valueOf(n)}</span>
            <span className="truncate">={n.ref}</span>
            <span>{n.scope ? wb.sheetById(n.scope)?.name : 'Workbook'}</span>
            <span className="truncate">{n.comment}</span>
          </div>
        ))}
      </div>
      <div className="mt-2">Refers to:</div>
      <div className="flex gap-2">
        <input className="xl-input flex-1" value={ref ?? (cur ? '=' + cur.ref : '')} disabled={!cur} onChange={(e) => setRef(e.target.value)} />
        <button className="xl-btn !min-w-0" disabled={!cur || ref === null} onClick={() => { if (cur && ref) { setNames(wb.names.map((n) => (n === cur ? { ...n, ref: ref.replace(/^=/, '') } : n))); setRef(null); } }}>✓</button>
      </div>
    </Dialog>
  );
}

export function NewNameDialog({ props }: { props: Record<string, unknown> }) {
  const wb = S().wb;
  const edit = props.edit as DefinedName | undefined;
  const st = S();
  const defaultName = (() => {
    const v = getComputed(st.wb.activeSheet, st.sel.active.r, st.sel.active.c - 1);
    return typeof v === 'string' ? v.replace(/[^A-Za-z0-9_.]/g, '_') : '';
  })();
  const [name, setName] = useState(edit?.name ?? defaultName);
  const [scope, setScope] = useState<string>(edit?.scope ?? '');
  const [comment, setComment] = useState(edit?.comment ?? '');
  const [ref, setRef] = useState(edit ? '=' + edit.ref : '=' + selectionRefText());
  return (
    <Dialog
      title={edit ? 'Edit Name' : 'New Name'}
      width={420}
      onOk={() => {
        const ok = defineName(name.trim(), ref, scope || undefined, comment || undefined, edit?.name);
        if (!ok) return false;
        if (props.back) setTimeout(() => openDialog(props.back as string), 0);
      }}
      onCancel={() => props.back && setTimeout(() => openDialog(props.back as string), 0)}
      testId="new-name"
    >
      <label className="flex items-center gap-2 my-1"><span className="w-20">Name:</span><input className="xl-input flex-1" value={name} onChange={(e) => setName(e.target.value)} data-testid="name-input" /></label>
      <label className="flex items-center gap-2 my-1"><span className="w-20">Scope:</span>
        <select className="xl-input flex-1" value={scope} onChange={(e) => setScope(e.target.value)}>
          <option value="">Workbook</option>
          {wb.sheets.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
        </select>
      </label>
      <label className="flex items-start gap-2 my-1"><span className="w-20">Comment:</span><textarea className="xl-input flex-1 !h-16 py-1" value={comment} onChange={(e) => setComment(e.target.value)} /></label>
      <label className="flex items-center gap-2 my-1"><span className="w-20">Refers to:</span><RefInput value={ref} onChange={setRef} width={290} testId="name-ref" /></label>
    </Dialog>
  );
}

export function CreateNamesDialog() {
  const [o, setO] = useState({ top: true, left: false, bottom: false, right: false });
  return (
    <Dialog title="Create Names from Selection" width={300} onOk={() => createNamesFromSelection(o.top, o.left, o.bottom, o.right)}>
      <div className="mb-1">Create names from values in the:</div>
      {(['top', 'left', 'bottom', 'right'] as const).map((k) => (
        <label key={k} className="flex gap-2 my-1"><input type="checkbox" checked={o[k]} onChange={(e) => setO({ ...o, [k]: e.target.checked })} /> {k[0].toUpperCase() + k.slice(1)} {k === 'top' || k === 'bottom' ? 'row' : 'column'}</label>
      ))}
    </Dialog>
  );
}

export function PasteNameDialog() {
  const names = S().wb.names;
  const [sel, setSel] = useState(names[0]?.name);
  return (
    <Dialog
      title="Paste Name"
      width={300}
      onOk={() => {
        if (!sel) return;
        const ed = S().edit;
        if (ed) setState({ edit: { ...ed, text: ed.text.slice(0, ed.caret) + sel + ed.text.slice(ed.caret), caret: ed.caret + sel.length } });
        else import('../../state/actions/edit').then((m) => m.beginEdit('enter', '=' + sel));
      }}
    >
      <div className="xl-list h-[160px]">
        {names.map((n) => <div key={n.name} className={sel === n.name ? 'xl-list-sel' : ''} onClick={() => setSel(n.name)}>{n.name}</div>)}
        {!names.length && <div className="opacity-60">(no names)</div>}
      </div>
    </Dialog>
  );
}

// ---------- tables ----------

export function CreateTableDialog({ props }: { props: Record<string, unknown> }) {
  const guess = tableSourceGuess();
  const rg0 = (props.range as Range | undefined) ?? guess?.range ?? primaryRange(S().sel);
  const [ref, setRef] = useState('=' + rangeToA1(rg0, true));
  const [header, setHeader] = useState((props.header as boolean | undefined) ?? guess?.header ?? guessHeader(S().wb.activeSheet, rg0));
  return (
    <Dialog
      title="Create Table"
      width={320}
      onOk={() => {
        const rg = parseRange(ref.replace(/^=/, '').replace(/\$/g, '').replace(/^.*!/, ''));
        if (!rg) {
          alertBox('The reference isn\'t valid.');
          return false;
        }
        createTable(rg, header, (props.style as string) ?? 'TableStyleMedium2');
      }}
      testId="create-table"
    >
      <div className="mb-1">Where is the data for your table?</div>
      <RefInput value={ref} onChange={setRef} width={270} />
      <label className="flex gap-2 mt-3"><input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} /> My table has headers</label>
    </Dialog>
  );
}

export function RemoveDuplicatesDialog({ props }: { props: Record<string, unknown> }) {
  const sheet = S().wb.activeSheet;
  const guess = tableSourceGuess();
  const rg = (props.range as Range | undefined) ?? guess?.range ?? primaryRange(S().sel);
  const [header, setHeader] = useState(guess?.header ?? true);
  const cols = Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => rg.c1 + i);
  const [checked, setChecked] = useState<Set<number>>(new Set(cols));
  return (
    <Dialog title="Remove Duplicates" width={380} onOk={() => removeDuplicates(rg, [...checked].sort((a, b) => a - b), header)} testId="remove-dups">
      <div className="mb-2">To delete duplicate values, select one or more columns that contain duplicates.</div>
      <div className="flex gap-2 mb-2">
        <button className="xl-btn !min-w-0" onClick={() => setChecked(new Set(cols))}>Select All</button>
        <button className="xl-btn !min-w-0" onClick={() => setChecked(new Set())}>Unselect All</button>
        <div className="flex-1" />
        <label className="flex gap-1 items-center"><input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} /> My data has headers</label>
      </div>
      <div className="xl-list h-[140px]">
        {cols.map((c) => (
          <label key={c} className="flex gap-2 px-1">
            <input type="checkbox" checked={checked.has(c)} onChange={() => { const n = new Set(checked); if (n.has(c)) n.delete(c); else n.add(c); setChecked(n); }} />
            {header ? displayTextAt(sheet, rg.r1, c) || `Column ${colToName(c)}` : `Column ${colToName(c)}`}
          </label>
        ))}
      </div>
    </Dialog>
  );
}

export function TextToColumnsDialog() {
  const st = S();
  const sheet = st.wb.activeSheet;
  const rg = primaryRange(st.sel);
  const samples = useMemo(() => {
    const out: string[] = [];
    for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 6); r++) out.push(String(getComputed(sheet, r, rg.c1) ?? ''));
    return out;
  }, [sheet, rg]);
  const [step, setStep] = useState(1);
  const [o, setO] = useState<TextToColumnsOpts>({ mode: 'delimited', delimiters: { tab: true, semicolon: false, comma: samples.some((s) => s.includes(',')), space: false, other: '' }, consecutive: false, qualifier: '"', widths: [] });
  const [dest, setDest] = useState('=$' + colToName(rg.c1) + '$' + (rg.r1 + 1));
  const preview = samples.map((s) => splitText(s, o));
  const run = () => {
    const d = parseRange(dest.replace(/^=/, '').replace(/\$/g, ''));
    textToColumns({ ...o, dest: d ? { r: d.r1, c: d.c1 } : undefined });
  };
  const footer = (
    <div className="flex justify-end gap-2 px-4 py-3">
      <button className="xl-btn" onClick={() => closeDialog()}>Cancel</button>
      <button className="xl-btn" disabled={step === 1} onClick={() => setStep(step - 1)}>&lt; Back</button>
      <button className="xl-btn" disabled={step === 3} onClick={() => setStep(step + 1)}>Next &gt;</button>
      <button className="xl-btn xl-btn-primary" onClick={() => { run(); closeDialog(); }}>Finish</button>
    </div>
  );
  return (
    <Dialog title={`Convert Text to Columns Wizard - Step ${step} of 3`} width={560} footer={footer} testId="text-to-columns">
      {step === 1 && (
        <div className="flex flex-col gap-2">
          <div>Choose the file type that best describes your data:</div>
          <label className="flex gap-2"><input type="radio" checked={o.mode === 'delimited'} onChange={() => setO({ ...o, mode: 'delimited' })} /> <b>Delimited</b> - Characters such as commas or tabs separate each field.</label>
          <label className="flex gap-2"><input type="radio" checked={o.mode === 'fixed'} onChange={() => setO({ ...o, mode: 'fixed' })} /> <b>Fixed width</b> - Fields are aligned in columns with spaces between each field.</label>
        </div>
      )}
      {step === 2 && o.mode === 'delimited' && (
        <div className="flex gap-4">
          <fieldset className="xl-fieldset">
            <legend>Delimiters</legend>
            {(['tab', 'semicolon', 'comma', 'space'] as const).map((k) => (
              <label key={k} className="flex gap-2"><input type="checkbox" checked={o.delimiters[k]} onChange={(e) => setO({ ...o, delimiters: { ...o.delimiters, [k]: e.target.checked } })} /> {k[0].toUpperCase() + k.slice(1)}</label>
            ))}
            <label className="flex gap-2 items-center"><input type="checkbox" checked={!!o.delimiters.other} readOnly /> Other: <input className="xl-input w-8" maxLength={1} value={o.delimiters.other} onChange={(e) => setO({ ...o, delimiters: { ...o.delimiters, other: e.target.value } })} /></label>
          </fieldset>
          <div className="flex flex-col gap-2">
            <label className="flex gap-2"><input type="checkbox" checked={o.consecutive} onChange={(e) => setO({ ...o, consecutive: e.target.checked })} /> Treat consecutive delimiters as one</label>
            <label className="flex gap-2 items-center">Text qualifier:
              <select className="xl-input" value={o.qualifier} onChange={(e) => setO({ ...o, qualifier: e.target.value as TextToColumnsOpts['qualifier'] })}><option value={'"'}>&quot;</option><option value="'">&apos;</option><option value="">{'{none}'}</option></select>
            </label>
          </div>
        </div>
      )}
      {step === 2 && o.mode === 'fixed' && (
        <div>
          <div className="mb-1">Click on the ruler to set break lines (positions in characters):</div>
          <input className="xl-input w-full" placeholder="e.g. 5, 12, 20" value={o.widths.join(', ')} onChange={(e) => setO({ ...o, widths: e.target.value.split(',').map((x) => parseInt(x, 10)).filter((x) => x > 0) })} />
        </div>
      )}
      {step === 3 && (
        <label className="flex items-center gap-2">Destination: <RefInput value={dest} onChange={setDest} width={200} /></label>
      )}
      <div className="mt-3 mb-1">Data preview</div>
      <div className="border overflow-auto h-[120px] font-mono text-[11px]" style={{ borderColor: 'var(--border)', background: 'var(--input-bg)' }}>
        <table className="border-collapse">
          <tbody>
            {preview.map((row, i) => (
              <tr key={i}>{row.map((c, j) => <td key={j} className="border px-1.5 whitespace-nowrap" style={{ borderColor: 'var(--border)' }}>{c}</td>)}</tr>
            ))}
          </tbody>
        </table>
      </div>
    </Dialog>
  );
}

// ---------- protection ----------

export function ProtectSheetDialog() {
  const [pw, setPw] = useState('');
  const [allow, setAllow] = useState({ ...DEFAULT_ALLOW });
  const labels: [keyof typeof allow, string][] = [
    ['selectLocked', 'Select locked cells'],
    ['selectUnlocked', 'Select unlocked cells'],
    ['formatCells', 'Format cells'],
    ['formatColumns', 'Format columns'],
    ['formatRows', 'Format rows'],
    ['insertColumns', 'Insert columns'],
    ['insertRows', 'Insert rows'],
    ['insertHyperlinks', 'Insert hyperlinks'],
    ['deleteColumns', 'Delete columns'],
    ['deleteRows', 'Delete rows'],
    ['sort', 'Sort'],
    ['autoFilter', 'Use AutoFilter'],
    ['pivotTables', 'Use PivotTable & PivotChart'],
    ['editObjects', 'Edit objects'],
    ['editScenarios', 'Edit scenarios'],
  ];
  return (
    <Dialog title="Protect Sheet" width={340} onOk={() => protectSheet(pw, allow)} testId="protect-sheet">
      <label className="flex gap-2 mb-2"><input type="checkbox" checked readOnly /> Protect worksheet and contents of locked cells</label>
      <div>Password to unprotect sheet:</div>
      <input type="password" className="xl-input w-full mb-2" value={pw} onChange={(e) => setPw(e.target.value)} data-testid="protect-pw" />
      <div className="mb-1">Allow all users of this worksheet to:</div>
      <div className="xl-list h-[180px]">
        {labels.map(([k, l]) => (
          <label key={k} className="flex gap-2 px-1"><input type="checkbox" checked={allow[k]} onChange={(e) => setAllow({ ...allow, [k]: e.target.checked })} /> {l}</label>
        ))}
      </div>
    </Dialog>
  );
}

export function UnprotectSheetDialog() {
  const [pw, setPw] = useState('');
  return (
    <Dialog title="Unprotect Sheet" width={300} onOk={async () => (await unprotectSheet(pw)) ? undefined : false}>
      <label className="flex items-center gap-2">Password: <input type="password" className="xl-input flex-1" value={pw} onChange={(e) => setPw(e.target.value)} /></label>
    </Dialog>
  );
}

// ---------- misc ----------

export function SymbolDialog() {
  const groups: [string, string][] = [
    ['Currency', '₹$€£¥¢₩₽₺₫₱'],
    ['Math', '±×÷≠≈≤≥∞√∑∏∫∂∆°µπΩ'],
    ['Arrows', '←↑→↓↔↕⇐⇒⇔'],
    ['Symbols', '©®™§¶•…‰†‡★☆✓✗☐☑♠♣♥♦'],
    ['Fractions', '½⅓⅔¼¾⅛'],
  ];
  const [sel, setSel] = useState('₹');
  return (
    <Dialog
      title="Symbol"
      width={420}
      okLabel="Insert"
      onOk={() => {
        const ed = S().edit;
        if (ed) setState({ edit: { ...ed, text: ed.text.slice(0, ed.caret) + sel + ed.text.slice(ed.caret), caret: ed.caret + sel.length } });
        else import('../../state/actions/edit').then((m) => m.beginEdit('edit', (S().wb.activeSheet.getCell(S().sel.active.r, S().sel.active.c)?.v ?? '') + sel));
      }}
    >
      {groups.map(([g, chars]) => (
        <div key={g} className="mb-2">
          <div className="text-[11px] opacity-70 mb-1">{g}</div>
          <div className="flex flex-wrap gap-1">
            {[...chars].map((ch) => (
              <button key={ch} className={'w-8 h-8 border text-[16px] ' + (sel === ch ? 'xl-btn-checked' : '')} style={{ borderColor: 'var(--border)' }} onClick={() => setSel(ch)} onDoubleClick={() => setSel(ch)}>{ch}</button>
            ))}
          </div>
        </div>
      ))}
    </Dialog>
  );
}

export function StatsDialog() {
  const wb = S().wb;
  const sheet = wb.activeSheet;
  let cells = 0;
  let formulas = 0;
  let notes = 0;
  for (const row of sheet.rows.values()) for (const c of row.values()) {
    if (c.v !== undefined || c.f) cells++;
    if (c.f) formulas++;
    if (c.note) notes++;
  }
  const used = sheet.usedRange();
  return (
    <Dialog title="Workbook Statistics" width={340} cancelLabel={null}>
      <div className="font-semibold mb-1">Current Sheet: {sheet.name}</div>
      <div className="grid grid-cols-[1fr_auto] gap-y-1">
        <span>End of sheet</span><span>{used ? `${colToName(used.c2)}${used.r2 + 1}` : 'A1'}</span>
        <span>Cells with data</span><span>{cells}</span>
        <span>Formulas</span><span>{formulas}</span>
        <span>Tables</span><span>{sheet.tables.length}</span>
        <span>Charts</span><span>{sheet.charts.length}</span>
        <span>Notes/Comments</span><span>{notes}</span>
      </div>
      <div className="font-semibold mt-3 mb-1">Workbook</div>
      <div className="grid grid-cols-[1fr_auto] gap-y-1">
        <span>Sheets</span><span>{wb.sheets.length}</span>
        <span>Defined names</span><span>{wb.names.length}</span>
      </div>
    </Dialog>
  );
}

export function EvaluateDialog() {
  const st = S();
  const sheet = st.wb.activeSheet;
  const { r, c } = st.sel.active;
  const f = sheet.getCell(r, c)?.f ?? '';
  const refs = tokenize(f).filter((t) => t.type === 'ref');
  const [step, setStep] = useState(0);
  let expr = f;
  const evalRef = (text: string) => {
    const v = st.engine.evaluate('=' + text, sheet);
    return isErrorVal(v) ? v.error : typeof v === 'string' ? `"${v}"` : String(v ?? 0);
  };
  for (let i = 0; i < Math.min(step, refs.length); i++) expr = expr.replace(refs[i].text, evalRef(refs[i].text));
  const done = step > refs.length;
  const result = st.engine.getValue(sheet, r, c);
  return (
    <Dialog title="Evaluate Formula" width={500} cancelLabel="Close" okLabel="" footer={
      <div className="flex justify-end gap-2 px-4 py-3">
        <button className="xl-btn xl-btn-primary" onClick={() => setStep(done ? 0 : step + 1)}>{done ? 'Restart' : 'Evaluate'}</button>
        <button className="xl-btn" onClick={() => closeDialog()}>Close</button>
      </div>
    }>
      <div className="mb-1">Reference: <b>{quoteSheetName(sheet.name)}!{colToName(c)}{r + 1}</b></div>
      <div className="mb-1">Evaluation:</div>
      <div className="border p-2 min-h-[70px] font-mono text-[12px] break-all" style={{ borderColor: 'var(--border)', background: 'var(--input-bg)' }}>
        {!f ? '(cell has no formula)' : done ? (isErrorVal(result) ? result.error : String(result)) : expr}
      </div>
      {refs[step] && !done && <div className="mt-2 text-[11px] opacity-75">Next: {refs[step].text} → {evalRef(refs[step].text)}</div>}
      <div className="mt-2 text-[11px] opacity-60">{resolveReference('A1') ? '' : ''}</div>
    </Dialog>
  );
}

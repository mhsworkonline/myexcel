'use client';
import { useMemo, useState } from 'react';
import {
  buildAccountingFormat,
  buildCurrencyFormat,
  buildNumberFormat,
  buildPercentFormat,
  buildScientificFormat,
  CUSTOM_FORMATS,
  DATE_FORMATS,
  formatValue,
  SPECIAL_FORMATS,
  TIME_FORMATS,
} from '../../model/numfmt';
import type { BorderEdge, BorderStyleName, CellStyle, HAlign, StylePatch, VAlign } from '../../model/styles';
import { applyStyleFn, autoRowHeights, FONT_SIZES } from '../../state/actions/format';
import { S, transact } from '../../state/store';
import { getComputed } from '../../state/values';
import { FONTS } from '../ribbon/HomeTab';
import { ColorPalette, THEME_BASE, STANDARD_COLORS } from '../ribbon/parts';
import { Dialog } from './Dialog';

type Tab = 'Number' | 'Alignment' | 'Font' | 'Border' | 'Fill' | 'Protection';
const TABS: Tab[] = ['Number', 'Alignment', 'Font', 'Border', 'Fill', 'Protection'];
const CATEGORIES = ['General', 'Number', 'Currency', 'Accounting', 'Date', 'Time', 'Percentage', 'Fraction', 'Scientific', 'Text', 'Special', 'Custom'];
const SYMBOLS: [string, string][] = [['', 'None'], ['$', '$ English (United States)'], ['€', '€ Euro'], ['£', '£ English (United Kingdom)'], ['¥', '¥ Japanese'], ['₹', '₹ Indian Rupee (lakh/crore)'], ['CHF', 'CHF Swiss']];
const FRACTIONS: [string, string][] = [['# ?/?', 'Up to one digit (1/4)'], ['# ??/??', 'Up to two digits (21/25)'], ['# ???/???', 'Up to three digits (312/943)'], ['# ?/2', 'As halves (1/2)'], ['# ?/4', 'As quarters (2/4)'], ['# ?/8', 'As eighths (4/8)'], ['# ??/16', 'As sixteenths (8/16)'], ['# ?/10', 'As tenths (3/10)'], ['# ??/100', 'As hundredths (30/100)']];
const LINE_STYLES: (BorderStyleName | null)[] = [null, 'hair', 'dotted', 'dashDotDot', 'dashDot', 'dashed', 'thin', 'mediumDashDotDot', 'slantDashDot', 'mediumDashDot', 'mediumDashed', 'medium', 'thick', 'double'];
const PATTERNS = ['solid', 'darkGray', 'mediumGray', 'lightGray', 'gray125', 'gray0625', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis'];

type EdgeKey = 'top' | 'bottom' | 'left' | 'right' | 'insideH' | 'insideV' | 'diagUp' | 'diagDown';

function guessCategory(fmt: string | undefined): string {
  if (!fmt) return 'General';
  if (fmt === '@') return 'Text';
  if (SPECIAL_FORMATS.some((s) => s.code === fmt) && !fmt.includes('₹')) return 'Special';
  if (/\*/.test(fmt)) return 'Accounting';
  if (fmt.includes('%')) return 'Percentage';
  if (/E[+-]/.test(fmt)) return 'Scientific';
  if (/\?\/[?\d]/.test(fmt)) return 'Fraction';
  if (/[$€£¥₹]|\[\$/.test(fmt)) return 'Currency';
  const noq = fmt.replace(/"[^"]*"|\[[^\]]*\]/g, '');
  if (/[dy]/i.test(noq) || /m{3,}/i.test(noq) || /^m\/|\/m/.test(noq)) return 'Date';
  if (/[hs]/i.test(noq)) return 'Time';
  if (/^[#0,.]+([_;\\()\-\[\]Red#0,. ]*)$/.test(fmt)) return 'Number';
  return 'Custom';
}

function decimalsOf(fmt: string | undefined): number {
  const m = /\.(0+)/.exec(fmt ?? '');
  return m ? m[1].length : fmt && fmt !== 'General' ? 0 : 2;
}

export function FormatCellsDialog({ props }: { props: Record<string, unknown> }) {
  const st = S();
  const sheet = st.wb.activeSheet;
  const a = st.sel.active;
  const cur = st.wb.styles.get(sheet.styleIdAt(a.r, a.c));
  const sampleVal = getComputed(sheet, a.r, a.c);
  const sample = typeof sampleVal === 'number' ? sampleVal : typeof sampleVal === 'string' ? sampleVal : 1234.5;
  const [tab, setTab] = useState<Tab>((props.tab as Tab) ?? 'Number');
  const [patch, setPatch] = useState<StylePatch>({});
  const eff: CellStyle = { ...cur, ...(patch as CellStyle) };
  const set = (p: StylePatch) => setPatch((o) => ({ ...o, ...p }));

  // number
  const [cat, setCat] = useState<string>((props.category as string) ?? guessCategory(cur.numFmt));
  const [decimals, setDecimals] = useState(decimalsOf(cur.numFmt));
  const [thousands, setThousands] = useState(!!cur.numFmt?.includes(','));
  const [neg, setNeg] = useState(cur.numFmt?.includes('[Red]') ? (cur.numFmt.includes('(') ? 3 : 1) : cur.numFmt?.includes('(') ? 2 : 0);
  const [symbol, setSymbol] = useState(cur.numFmt?.includes('₹') ? '₹' : cur.numFmt?.includes('€') ? '€' : cur.numFmt?.includes('£') ? '£' : '$');
  const [custom, setCustom] = useState(cur.numFmt ?? 'General');
  const [dateFmt, setDateFmt] = useState(cur.numFmt && DATE_FORMATS.includes(cur.numFmt) ? cur.numFmt : DATE_FORMATS[0]);
  const [timeFmt, setTimeFmt] = useState(cur.numFmt && TIME_FORMATS.includes(cur.numFmt) ? cur.numFmt : TIME_FORMATS[0]);
  const [fracFmt, setFracFmt] = useState(FRACTIONS[0][0]);
  const [special, setSpecial] = useState(SPECIAL_FORMATS[0].code);
  const [numTouched, setNumTouched] = useState(false);

  const numFmt = useMemo(() => {
    switch (cat) {
      case 'General': return 'General';
      case 'Number': return buildNumberFormat(decimals, thousands, neg);
      case 'Currency': return buildCurrencyFormat(decimals, symbol, neg);
      case 'Accounting': return buildAccountingFormat(decimals, symbol);
      case 'Date': return dateFmt;
      case 'Time': return timeFmt;
      case 'Percentage': return buildPercentFormat(decimals);
      case 'Fraction': return fracFmt;
      case 'Scientific': return buildScientificFormat(decimals);
      case 'Text': return '@';
      case 'Special': return special;
      default: return custom;
    }
  }, [cat, decimals, thousands, neg, symbol, dateFmt, timeFmt, fracFmt, special, custom]);

  // borders
  const [lineStyle, setLineStyle] = useState<BorderStyleName>('thin');
  const [lineColor, setLineColor] = useState('#000000');
  const [edges, setEdges] = useState<Partial<Record<EdgeKey, BorderEdge | null>>>(() => ({
    top: cur.bTop,
    bottom: cur.bBottom,
    left: cur.bLeft,
    right: cur.bRight,
    diagUp: cur.bDiagUp,
    diagDown: cur.bDiagDown,
  }));
  const [edgesTouched, setEdgesTouched] = useState<Set<EdgeKey>>(new Set());
  const multi = st.sel.ranges.some((rg) => rg.r1 !== rg.r2 || rg.c1 !== rg.c2);
  const toggleEdge = (k: EdgeKey) => {
    setEdges((e) => ({ ...e, [k]: e[k] ? null : { style: lineStyle, color: lineColor } }));
    setEdgesTouched((t) => new Set(t).add(k));
  };
  const preset = (p: 'none' | 'outline' | 'inside') => {
    const e: BorderEdge = { style: lineStyle, color: lineColor };
    if (p === 'none') {
      setEdges({ top: null, bottom: null, left: null, right: null, insideH: null, insideV: null, diagUp: null, diagDown: null });
      setEdgesTouched(new Set(['top', 'bottom', 'left', 'right', 'insideH', 'insideV', 'diagUp', 'diagDown']));
    } else if (p === 'outline') {
      setEdges((x) => ({ ...x, top: e, bottom: e, left: e, right: e }));
      setEdgesTouched((t) => new Set([...t, 'top', 'bottom', 'left', 'right'] as EdgeKey[]));
    } else {
      setEdges((x) => ({ ...x, insideH: e, insideV: e }));
      setEdgesTouched((t) => new Set([...t, 'insideH', 'insideV'] as EdgeKey[]));
    }
  };

  const ok = () => {
    const { merge: mergeFlag, ...rest } = patch as StylePatch & { merge?: boolean };
    const finalPatch: StylePatch = { ...rest };
    if (numTouched) finalPatch.numFmt = numFmt === 'General' ? null : numFmt;
    transact('Format Cells', (tx) => {
      for (const rg of st.sel.ranges) {
        applyStyleFn(tx, sheet, [rg], (_b, r, c) => {
          const p: StylePatch = { ...finalPatch };
          const fullRow = r < 0;
          const fullCol = c < 0;
          const top = r === rg.r1 || fullRow;
          const bottom = r === rg.r2 || fullRow;
          const left = c === rg.c1 || fullCol;
          const right = c === rg.c2 || fullCol;
          const val = (k: EdgeKey) => (edges[k] === undefined ? undefined : edges[k]);
          if (edgesTouched.has('top') && top) p.bTop = val('top') ?? null;
          if (edgesTouched.has('bottom') && bottom) p.bBottom = val('bottom') ?? null;
          if (edgesTouched.has('left') && left) p.bLeft = val('left') ?? null;
          if (edgesTouched.has('right') && right) p.bRight = val('right') ?? null;
          if (edgesTouched.has('insideH')) {
            if (!bottom) p.bBottom = val('insideH') ?? null;
            if (!top) p.bTop = val('insideH') ?? null;
          }
          if (edgesTouched.has('insideV')) {
            if (!right) p.bRight = val('insideV') ?? null;
            if (!left) p.bLeft = val('insideV') ?? null;
          }
          if (edgesTouched.has('diagUp')) p.bDiagUp = val('diagUp') ?? null;
          if (edgesTouched.has('diagDown')) p.bDiagDown = val('diagDown') ?? null;
          return p;
        });
      }
      if (finalPatch.wrap !== undefined || finalPatch.fontSize !== undefined) autoRowHeights(tx, sheet, st.sel.ranges);
    });
    if (mergeFlag !== undefined) {
      import('../../state/actions/format').then((m) => m.mergeCells(mergeFlag ? 'merge' : 'unmerge'));
    }
  };

  const preview = formatValue(typeof sample === 'number' ? sample : sample, numFmt);

  return (
    <Dialog title="Format Cells" width={560} onOk={ok} testId="format-cells">
      <div className="flex border-b mb-3" style={{ borderColor: 'var(--border)' }}>
        {TABS.map((t) => (
          <button key={t} className={'xl-dtab ' + (tab === t ? 'xl-dtab-active' : '')} onClick={() => setTab(t)} data-testid={`fc-tab-${t}`}>
            {t}
          </button>
        ))}
      </div>
      <div className="h-[330px]">
        {tab === 'Number' && (
          <div className="flex gap-3 h-full">
            <div className="w-[130px]">
              <div className="mb-1">Category:</div>
              <div className="xl-list h-[300px]">
                {CATEGORIES.map((c) => (
                  <div key={c} className={cat === c ? 'xl-list-sel' : ''} onClick={() => { setCat(c); setNumTouched(true); if (c === 'Custom') setCustom(numFmt); }} data-testid={`fc-cat-${c}`}>
                    {c}
                  </div>
                ))}
              </div>
            </div>
            <div className="flex-1 flex flex-col gap-2">
              <fieldset className="xl-fieldset">
                <legend>Sample</legend>
                <div className="h-5" style={{ color: preview.color }} data-testid="fc-sample">{preview.left !== undefined ? preview.left + preview.right : preview.text}</div>
              </fieldset>
              {(cat === 'Number' || cat === 'Currency' || cat === 'Accounting' || cat === 'Percentage' || cat === 'Scientific') && (
                <label className="flex items-center gap-2">
                  Decimal places:
                  <input type="number" min={0} max={30} className="xl-input w-16" value={decimals} onChange={(e) => { setDecimals(Math.max(0, Math.min(30, +e.target.value))); setNumTouched(true); }} />
                </label>
              )}
              {cat === 'Number' && (
                <label className="flex items-center gap-2">
                  <input type="checkbox" checked={thousands} onChange={(e) => { setThousands(e.target.checked); setNumTouched(true); }} /> Use 1000 Separator (,)
                </label>
              )}
              {(cat === 'Currency' || cat === 'Accounting') && (
                <label className="flex items-center gap-2">
                  Symbol:
                  <select className="xl-input flex-1" value={symbol} onChange={(e) => { setSymbol(e.target.value); setNumTouched(true); }} data-testid="fc-symbol">
                    {SYMBOLS.map(([s, l]) => (
                      <option key={s} value={s}>{l}</option>
                    ))}
                  </select>
                </label>
              )}
              {(cat === 'Number' || cat === 'Currency') && (
                <div>
                  <div className="mb-1">Negative numbers:</div>
                  <div className="xl-list h-[80px]">
                    {[0, 1, 2, 3].map((n) => {
                      const f = cat === 'Number' ? buildNumberFormat(decimals, thousands, n) : buildCurrencyFormat(decimals, symbol, n);
                      const r = formatValue(-1234.1, f);
                      return (
                        <div key={n} className={neg === n ? 'xl-list-sel' : ''} style={{ color: r.color }} onClick={() => { setNeg(n); setNumTouched(true); }}>
                          {r.text}
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
              {cat === 'Date' && <FmtList list={DATE_FORMATS} value={dateFmt} onChange={(v) => { setDateFmt(v); setNumTouched(true); }} sample={typeof sample === 'number' ? sample : 45293.5} label="Type:" />}
              {cat === 'Time' && <FmtList list={TIME_FORMATS} value={timeFmt} onChange={(v) => { setTimeFmt(v); setNumTouched(true); }} sample={typeof sample === 'number' ? sample : 45293.5625} label="Type:" />}
              {cat === 'Fraction' && (
                <div>
                  <div className="mb-1">Type:</div>
                  <div className="xl-list h-[160px]">
                    {FRACTIONS.map(([f, l]) => (
                      <div key={f} className={fracFmt === f ? 'xl-list-sel' : ''} onClick={() => { setFracFmt(f); setNumTouched(true); }}>{l}</div>
                    ))}
                  </div>
                </div>
              )}
              {cat === 'Special' && (
                <div>
                  <div className="mb-1">Type:</div>
                  <div className="xl-list h-[140px]">
                    {SPECIAL_FORMATS.map((s) => (
                      <div key={s.code} className={special === s.code ? 'xl-list-sel' : ''} onClick={() => { setSpecial(s.code); setNumTouched(true); }}>{s.name}</div>
                    ))}
                  </div>
                </div>
              )}
              {cat === 'Custom' && (
                <div className="flex flex-col gap-1">
                  <div>Type:</div>
                  <input className="xl-input" value={custom} onChange={(e) => { setCustom(e.target.value); setNumTouched(true); }} data-testid="fc-custom" />
                  <div className="xl-list h-[170px]">
                    {CUSTOM_FORMATS.map((f) => (
                      <div key={f} className={custom === f ? 'xl-list-sel' : ''} onClick={() => { setCustom(f); setNumTouched(true); }}>{f}</div>
                    ))}
                  </div>
                </div>
              )}
              <div className="text-[11px] opacity-70 leading-4">
                {cat === 'General' && 'General format cells have no specific number format.'}
                {cat === 'Number' && 'Number is used for general display of numbers. Currency and Accounting offer specialized formatting for monetary value.'}
                {cat === 'Currency' && 'Currency formats are used for general monetary values. Use Accounting formats to align decimal points in a column. ₹ uses Indian lakh/crore digit grouping.'}
                {cat === 'Accounting' && 'Accounting formats line up the currency symbols and decimal points in a column.'}
                {cat === 'Date' && 'Date formats display date and time serial numbers as date values.'}
                {cat === 'Time' && 'Time formats display date and time serial numbers as time values.'}
                {cat === 'Percentage' && 'Percentage formats multiply the cell value by 100 and display the result with a percent symbol.'}
                {cat === 'Text' && 'Text format cells are treated as text even when a number is in the cell. The cell is displayed exactly as entered.'}
                {cat === 'Special' && 'Special formats are useful for tracking list and database values.'}
                {cat === 'Custom' && 'Type the number format code, using one of the existing codes as a starting point.'}
              </div>
            </div>
          </div>
        )}

        {tab === 'Alignment' && (
          <div className="flex gap-4">
            <div className="flex-1 flex flex-col gap-2">
              <fieldset className="xl-fieldset">
                <legend>Text alignment</legend>
                <label className="flex items-center gap-2 my-1">
                  <span className="w-20">Horizontal:</span>
                  <select className="xl-input flex-1" value={eff.hAlign ?? 'general'} onChange={(e) => set({ hAlign: e.target.value === 'general' ? null : (e.target.value as HAlign) })} data-testid="fc-halign">
                    <option value="general">General</option>
                    <option value="left">Left (Indent)</option>
                    <option value="center">Center</option>
                    <option value="right">Right (Indent)</option>
                    <option value="fill">Fill</option>
                    <option value="justify">Justify</option>
                    <option value="centerContinuous">Center Across Selection</option>
                    <option value="distributed">Distributed (Indent)</option>
                  </select>
                </label>
                <label className="flex items-center gap-2 my-1">
                  <span className="w-20">Vertical:</span>
                  <select className="xl-input flex-1" value={eff.vAlign ?? 'bottom'} onChange={(e) => set({ vAlign: e.target.value === 'bottom' ? null : (e.target.value as VAlign) })}>
                    <option value="top">Top</option>
                    <option value="middle">Center</option>
                    <option value="bottom">Bottom</option>
                    <option value="justify">Justify</option>
                    <option value="distributed">Distributed</option>
                  </select>
                </label>
                <label className="flex items-center gap-2 my-1">
                  <span className="w-20">Indent:</span>
                  <input type="number" min={0} max={250} className="xl-input w-16" value={eff.indent ?? 0} onChange={(e) => set({ indent: +e.target.value || null })} />
                </label>
              </fieldset>
              <fieldset className="xl-fieldset">
                <legend>Text control</legend>
                <label className="flex gap-2 my-1"><input type="checkbox" checked={!!eff.wrap} onChange={(e) => set({ wrap: e.target.checked || null })} /> Wrap text</label>
                <label className="flex gap-2 my-1"><input type="checkbox" checked={!!eff.shrink} onChange={(e) => set({ shrink: e.target.checked || null })} /> Shrink to fit</label>
                <label className="flex gap-2 my-1">
                  <input type="checkbox" defaultChecked={!!sheet.mergeAt(a.r, a.c)} onChange={(e) => setPatch((o) => ({ ...o, merge: e.target.checked } as StylePatch))} /> Merge cells
                </label>
              </fieldset>
            </div>
            <fieldset className="xl-fieldset w-[150px]">
              <legend>Orientation</legend>
              <div className="flex gap-2">
                <button className={'w-8 h-24 border flex flex-col items-center justify-center text-[11px] leading-3 ' + (eff.rotation === 255 ? 'xl-btn-checked' : '')} style={{ borderColor: 'var(--border)' }} onClick={() => set({ rotation: eff.rotation === 255 ? null : 255 })}>
                  {'Text'.split('').map((ch, i) => <span key={i}>{ch}</span>)}
                </button>
                <div className="relative w-24 h-24 border rounded-full" style={{ borderColor: 'var(--border)' }}>
                  <div className="absolute left-1/2 top-1/2 origin-left h-[2px] w-10" style={{ background: 'var(--accent)', transform: `rotate(${-(eff.rotation && eff.rotation !== 255 ? (eff.rotation > 90 ? -(eff.rotation - 90) : eff.rotation) : 0)}deg)` }} />
                  <span className="absolute left-[46px] top-[38px] text-[10px]">Text</span>
                </div>
              </div>
              <label className="flex items-center gap-1 mt-2">
                <input
                  type="number"
                  min={-90}
                  max={90}
                  className="xl-input w-16"
                  value={eff.rotation && eff.rotation !== 255 ? (eff.rotation > 90 ? -(eff.rotation - 90) : eff.rotation) : 0}
                  onChange={(e) => {
                    const d = Math.max(-90, Math.min(90, +e.target.value || 0));
                    set({ rotation: d === 0 ? null : d < 0 ? 90 - d : d });
                  }}
                />
                Degrees
              </label>
            </fieldset>
          </div>
        )}

        {tab === 'Font' && (
          <div className="flex flex-col gap-2">
            <div className="flex gap-3">
              <div className="flex-1">
                <div>Font:</div>
                <input className="xl-input w-full" value={eff.fontName ?? 'Calibri'} onChange={(e) => set({ fontName: e.target.value })} />
                <div className="xl-list h-[110px] mt-1">
                  {FONTS.map((f) => (
                    <div key={f} className={(eff.fontName ?? 'Calibri') === f ? 'xl-list-sel' : ''} style={{ fontFamily: `"${f}"` }} onClick={() => set({ fontName: f })}>{f}</div>
                  ))}
                </div>
              </div>
              <div className="w-[120px]">
                <div>Font style:</div>
                <div className="xl-list h-[134px]">
                  {[['Regular', false, false], ['Italic', false, true], ['Bold', true, false], ['Bold Italic', true, true]].map(([l, b, i]) => (
                    <div key={l as string} className={!!eff.bold === b && !!eff.italic === i ? 'xl-list-sel' : ''} onClick={() => set({ bold: b ? true : null, italic: i ? true : null })}>{l as string}</div>
                  ))}
                </div>
              </div>
              <div className="w-[70px]">
                <div>Size:</div>
                <input className="xl-input w-full" value={eff.fontSize ?? 11} onChange={(e) => set({ fontSize: parseFloat(e.target.value) || 11 })} />
                <div className="xl-list h-[110px] mt-1">
                  {FONT_SIZES.map((s) => (
                    <div key={s} className={(eff.fontSize ?? 11) === s ? 'xl-list-sel' : ''} onClick={() => set({ fontSize: s })}>{s}</div>
                  ))}
                </div>
              </div>
            </div>
            <div className="flex gap-3">
              <label className="flex-1">
                Underline:
                <select className="xl-input w-full" value={eff.underline ?? 'none'} onChange={(e) => set({ underline: e.target.value === 'none' ? null : (e.target.value as CellStyle['underline']) })}>
                  <option value="none">None</option>
                  <option value="single">Single</option>
                  <option value="double">Double</option>
                  <option value="singleAccounting">Single Accounting</option>
                  <option value="doubleAccounting">Double Accounting</option>
                </select>
              </label>
              <label className="w-[200px]">
                Color:
                <div className="flex items-center gap-2">
                  <input type="color" value={eff.fontColor ?? '#000000'} onChange={(e) => set({ fontColor: e.target.value.toUpperCase() })} />
                  <button className="xl-btn !min-w-0" onClick={() => set({ fontColor: null })}>Automatic</button>
                </div>
              </label>
            </div>
            <div className="flex gap-3">
              <fieldset className="xl-fieldset flex-1">
                <legend>Effects</legend>
                <label className="flex gap-2"><input type="checkbox" checked={!!eff.strike} onChange={(e) => set({ strike: e.target.checked || null })} /> Strikethrough</label>
                <label className="flex gap-2"><input type="checkbox" checked={eff.vertAlign === 'superscript'} onChange={(e) => set({ vertAlign: e.target.checked ? 'superscript' : null })} /> Superscript</label>
                <label className="flex gap-2"><input type="checkbox" checked={eff.vertAlign === 'subscript'} onChange={(e) => set({ vertAlign: e.target.checked ? 'subscript' : null })} /> Subscript</label>
              </fieldset>
              <fieldset className="xl-fieldset flex-1">
                <legend>Preview</legend>
                <div className="h-14 flex items-center justify-center overflow-hidden" style={{ fontFamily: `"${eff.fontName ?? 'Calibri'}", Carlito, sans-serif`, fontSize: Math.min(28, (eff.fontSize ?? 11) * 1.33), fontWeight: eff.bold ? 700 : 400, fontStyle: eff.italic ? 'italic' : 'normal', textDecoration: [eff.underline ? 'underline' : '', eff.strike ? 'line-through' : ''].join(' '), color: eff.fontColor ?? 'var(--text)' }}>
                  AaBbCcYyZz
                </div>
              </fieldset>
            </div>
          </div>
        )}

        {tab === 'Border' && (
          <div className="flex gap-4">
            <fieldset className="xl-fieldset w-[150px]">
              <legend>Line</legend>
              <div>Style:</div>
              <div className="grid grid-cols-2 gap-1 border p-1" style={{ borderColor: 'var(--border)' }}>
                {LINE_STYLES.map((s, i) => (
                  <button key={i} className={'h-5 flex items-center px-1 ' + (lineStyle === s || (!s && false) ? 'xl-btn-checked' : '')} onClick={() => s && setLineStyle(s)} title={s ?? 'None'}>
                    {s ? <BorderSample s={s} /> : <span className="text-[11px]">None</span>}
                  </button>
                ))}
              </div>
              <div className="mt-2">Color:</div>
              <div className="flex flex-wrap gap-[2px] mt-1">
                {[...THEME_BASE, ...STANDARD_COLORS].map((c, i) => (
                  <button key={c + i} className="xl-color-cell" style={{ background: c, outline: lineColor === c ? '2px solid var(--accent)' : undefined }} onClick={() => setLineColor(c)} />
                ))}
              </div>
            </fieldset>
            <div className="flex-1">
              <div className="mb-1">Presets</div>
              <div className="flex gap-6 mb-3">
                {(['none', 'outline', 'inside'] as const).map((p) => (
                  <button key={p} className="flex flex-col items-center gap-1" onClick={() => preset(p)} disabled={p === 'inside' && !multi}>
                    <span className="w-9 h-9 border flex items-center justify-center" style={{ borderColor: 'var(--border-strong)', opacity: p === 'inside' && !multi ? 0.4 : 1 }}>
                      <svg width="24" height="24"><rect x="2" y="2" width="20" height="20" fill="none" stroke={p === 'outline' ? '#000' : '#bbb'} strokeWidth={p === 'outline' ? 2 : 1} /><path d="M12 2v20M2 12h20" stroke={p === 'inside' ? '#000' : '#ddd'} strokeWidth={p === 'inside' ? 2 : 1} /></svg>
                    </span>
                    <span className="capitalize text-[11px]">{p}</span>
                  </button>
                ))}
              </div>
              <div className="mb-1">Border</div>
              <div className="flex gap-2">
                <div className="flex flex-col gap-1">
                  {(['top', 'insideH', 'bottom', 'diagUp'] as EdgeKey[]).map((k) => (
                    <EdgeBtn key={k} k={k} on={!!edges[k]} onClick={() => toggleEdge(k)} disabled={(k === 'insideH' && !multi)} />
                  ))}
                </div>
                <BorderPreview edges={edges} multi={multi} onToggle={toggleEdge} />
              </div>
              <div className="flex gap-1 mt-1 ml-[44px]">
                {(['left', 'insideV', 'right', 'diagDown'] as EdgeKey[]).map((k) => (
                  <EdgeBtn key={k} k={k} on={!!edges[k]} onClick={() => toggleEdge(k)} disabled={k === 'insideV' && !multi} />
                ))}
              </div>
              <div className="text-[11px] opacity-70 mt-2">The selected border style can be applied by clicking the presets, preview diagram or the buttons above.</div>
            </div>
          </div>
        )}

        {tab === 'Fill' && (
          <div className="flex gap-4">
            <div className="flex-1">
              <div className="mb-1">Background Color:</div>
              <button className="xl-btn mb-2" onClick={() => set({ fillColor: null, patternType: null })}>No Color</button>
              <ColorPalette close={() => undefined} onPick={(c) => set({ fillColor: c })} />
            </div>
            <div className="w-[200px] flex flex-col gap-2">
              <label>
                Pattern Color:
                <input type="color" className="block" value={eff.patternColor ?? '#000000'} onChange={(e) => set({ patternColor: e.target.value.toUpperCase() })} />
              </label>
              <label>
                Pattern Style:
                <select className="xl-input w-full" value={eff.patternType ?? 'solid'} onChange={(e) => set({ patternType: e.target.value === 'solid' ? null : e.target.value })}>
                  {PATTERNS.map((p) => (
                    <option key={p} value={p}>{p}</option>
                  ))}
                </select>
              </label>
              <fieldset className="xl-fieldset mt-auto">
                <legend>Sample</legend>
                <div className="h-12" style={{ background: eff.fillColor ?? 'transparent' }} data-testid="fill-sample" />
              </fieldset>
            </div>
          </div>
        )}

        {tab === 'Protection' && (
          <div className="flex flex-col gap-2 pt-1">
            <label className="flex gap-2"><input type="checkbox" checked={eff.locked !== false} onChange={(e) => set({ locked: e.target.checked ? null : false })} data-testid="fc-locked" /> Locked</label>
            <label className="flex gap-2"><input type="checkbox" checked={!!eff.hidden} onChange={(e) => set({ hidden: e.target.checked || null })} /> Hidden</label>
            <div className="text-[11px] opacity-75 leading-4 mt-2">Locking cells or hiding formulas has no effect until you protect the worksheet (Review tab, Protect group, Protect Sheet button).</div>
          </div>
        )}
      </div>
    </Dialog>
  );
}

function FmtList({ list, value, onChange, sample, label }: { list: string[]; value: string; onChange: (v: string) => void; sample: number; label: string }) {
  return (
    <div>
      <div className="mb-1">{label}</div>
      <div className="xl-list h-[170px]">
        {list.map((f) => (
          <div key={f} className={value === f ? 'xl-list-sel' : ''} onClick={() => onChange(f)}>{formatValue(sample, f).text}</div>
        ))}
      </div>
    </div>
  );
}

function BorderSample({ s }: { s: BorderStyleName }) {
  const dash: Record<string, string> = { hair: '1 1', dotted: '1 2', dashDotDot: '6 2 2 2 2 2', dashDot: '6 2 2 2', dashed: '4 2', mediumDashDotDot: '6 2 2 2 2 2', slantDashDot: '6 2 2 2', mediumDashDot: '6 2 2 2', mediumDashed: '5 2' };
  const w = /medium|slant/.test(s) ? 2 : s === 'thick' ? 3 : 1;
  if (s === 'double') return <svg width="56" height="8"><path d="M0 2h56M0 6h56" stroke="currentColor" /></svg>;
  return <svg width="56" height="8"><path d="M0 4h56" stroke="currentColor" strokeWidth={w} strokeDasharray={dash[s]} /></svg>;
}

function EdgeBtn({ k, on, onClick, disabled }: { k: EdgeKey; on: boolean; onClick: () => void; disabled?: boolean }) {
  const paths: Record<EdgeKey, string> = {
    top: 'M3 4h14',
    insideH: 'M3 10h14',
    bottom: 'M3 16h14',
    diagUp: 'M3 17L17 3',
    left: 'M4 3v14',
    insideV: 'M10 3v14',
    right: 'M16 3v14',
    diagDown: 'M3 3l14 14',
  };
  return (
    <button disabled={disabled} className={'w-7 h-7 border flex items-center justify-center ' + (on ? 'xl-btn-checked' : '')} style={{ borderColor: 'var(--border-strong)', opacity: disabled ? 0.35 : 1 }} onClick={onClick} title={k} data-testid={`edge-${k}`}>
      <svg width="20" height="20"><rect x="3" y="3" width="14" height="14" fill="none" stroke="#ccc" strokeDasharray="1 1" /><path d={paths[k]} stroke="currentColor" strokeWidth="1.6" /></svg>
    </button>
  );
}

function BorderPreview({ edges, multi, onToggle }: { edges: Partial<Record<EdgeKey, BorderEdge | null>>; multi: boolean; onToggle: (k: EdgeKey) => void }) {
  const line = (e: BorderEdge | null | undefined) => (e ? { stroke: e.color ?? '#000', strokeWidth: e.style === 'thick' ? 3 : /medium/.test(e.style) ? 2 : 1 } : { stroke: 'transparent', strokeWidth: 1 });
  return (
    <svg width="130" height="120" className="border" style={{ borderColor: 'var(--border)', background: '#fff' }}>
      <text x="30" y="40" fontSize="11" fill="#555">Text</text>
      {multi && <text x="80" y="40" fontSize="11" fill="#555">Text</text>}
      {multi && <text x="30" y="90" fontSize="11" fill="#555">Text</text>}
      {multi && <text x="80" y="90" fontSize="11" fill="#555">Text</text>}
      <path d="M8 10v-4M4 10h4M122 10v-4M126 10h-4M8 110v4M4 110h4M122 110v4M126 110h-4" stroke="#999" />
      <line x1="10" y1="10" x2="120" y2="10" {...line(edges.top)} />
      <line x1="10" y1="110" x2="120" y2="110" {...line(edges.bottom)} />
      <line x1="10" y1="10" x2="10" y2="110" {...line(edges.left)} />
      <line x1="120" y1="10" x2="120" y2="110" {...line(edges.right)} />
      {multi && <line x1="10" y1="60" x2="120" y2="60" {...line(edges.insideH)} />}
      {multi && <line x1="65" y1="10" x2="65" y2="110" {...line(edges.insideV)} />}
      <line x1="10" y1="110" x2="120" y2="10" {...line(edges.diagUp)} />
      <line x1="10" y1="10" x2="120" y2="110" {...line(edges.diagDown)} />
      <rect x="0" y="0" width="130" height="20" fill="transparent" onClick={() => onToggle('top')} />
      <rect x="0" y="100" width="130" height="20" fill="transparent" onClick={() => onToggle('bottom')} />
      <rect x="0" y="20" width="20" height="80" fill="transparent" onClick={() => onToggle('left')} />
      <rect x="110" y="20" width="20" height="80" fill="transparent" onClick={() => onToggle('right')} />
    </svg>
  );
}

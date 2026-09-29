'use client';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  AArrowDown,
  AArrowUp,
  ArrowDownToLine,
  ArrowUpToLine,
  AlignVerticalJustifyCenter,
  Bold,
  Copy,
  Eraser,
  IndentDecrease,
  IndentIncrease,
  Italic,
  Paintbrush,
  Percent,
  Scissors,
  Strikethrough,
  Underline,
  WrapText,
  ArrowDownNarrowWide,
  ChevronsDown,
  RotateCcw,
} from 'lucide-react';
import { FMT, formatValue } from '../../model/numfmt';
import { primaryRange } from '../../model/selection';
import { copyToSystem, pasteFromSystem, DEFAULT_PASTE, PasteOptions } from '../../state/actions/clipboard';
import { clearSelection, protectedAlert } from '../../state/actions/edit';
import { fillDirection, flashFill } from '../../state/actions/fill';
import { goToSpecial } from '../../state/actions/find';
import {
  activeStyle,
  applyBorder,
  applyStyle,
  autoFitColumns,
  autoFitRows,
  CELL_STYLES,
  changeDecimals,
  changeIndent,
  commaStyle,
  FONT_SIZES,
  growFont,
  hideRowsCols,
  mergeCells,
  setHAlign,
  setNumFmt,
  setRotation,
  setVAlign,
  startFormatPainter,
  toggleStyle,
  toggleUnderline,
  toggleWrap,
  BorderPreset,
} from '../../state/actions/format';
import { autoSum } from '../../state/actions/formulas';
import { clearAllFilters, quickSort, reapplyFilter, toggleAutoFilter } from '../../state/actions/sortFilter';
import {
  addSheet,
  deleteCellsDialog,
  deleteCols,
  deleteRows,
  deleteSheet,
  hideSheet,
  insertCellsDialog,
  insertCols,
  insertRows,
  setTabColor,
} from '../../state/actions/structure';
import {
  addColorScale,
  addDataBar,
  addIconSet,
  clearCFRules,
  COLOR_SCALES,
  DATA_BAR_COLORS,
  ICON_SETS,
  tableSourceGuess,
  toggleLockCell,
  unprotectSheet,
} from '../../state/actions/data';
import { openDialog, S, setState, useStore } from '../../state/store';
import { getComputed } from '../../state/values';
import { TABLE_STYLE_NAMES, tableLook } from '../../state/tableStyles';
import { drawIcon } from '../grid/render';
import { MenuItem } from '../Menu';
import {
  AutoSumIcon,
  BordersIcon,
  CellStylesIcon,
  CondFormatIcon,
  DeleteCellsIcon,
  FillBucketIcon,
  FindSelectIcon,
  FontColorIcon,
  FormatCellsIcon,
  FormatTableIcon,
  InsertCellsIcon,
  MergeIcon,
  PasteIcon,
  SortFilterIcon,
} from '../icons';
import { Col, ColorIcon, ColorPalette, Combo, Group, LargeButton, Row, SmallButton } from './parts';
import { useEffect, useRef, useState } from 'react';
import type { TableDef } from '../../model/types';

export const FONTS = [
  'Calibri',
  'Calibri Light',
  'Aptos',
  'Aptos Narrow',
  'Arial',
  'Arial Black',
  'Arial Narrow',
  'Cambria',
  'Candara',
  'Century Gothic',
  'Comic Sans MS',
  'Consolas',
  'Constantia',
  'Corbel',
  'Courier New',
  'Franklin Gothic Medium',
  'Garamond',
  'Georgia',
  'Impact',
  'Lucida Console',
  'Palatino Linotype',
  'Segoe UI',
  'Tahoma',
  'Times New Roman',
  'Trebuchet MS',
  'Verdana',
];

let lastFill = '#FFFF00';
let lastFont = '#FF0000';
let lastBorder: BorderPreset = 'bottom';

export function pasteMenu(): MenuItem[] {
  const p = (o: Partial<PasteOptions>) => () => pasteFromSystem({ ...DEFAULT_PASTE, ...o });
  return [
    { label: 'Paste', disabled: false, onClick: p({}), shortcut: 'Ctrl+V' },
    { label: 'Formulas', onClick: p({ what: 'formulas' }) },
    { label: 'Formulas & Number Formatting', onClick: p({ what: 'formulasAndNumberFormats' }) },
    { label: 'Keep Source Formatting', onClick: p({}) },
    { label: 'No Borders', onClick: p({ what: 'allExceptBorders' }) },
    { label: 'Keep Source Column Widths', onClick: async () => { await pasteFromSystem(DEFAULT_PASTE); await pasteFromSystem({ ...DEFAULT_PASTE, what: 'columnWidths' }); } },
    { label: 'Transpose', onClick: p({ transpose: true }) },
    { separator: true },
    { label: 'Values', onClick: p({ what: 'values' }) },
    { label: 'Values & Number Formatting', onClick: p({ what: 'valuesAndNumberFormats' }) },
    { label: 'Values & Source Formatting', onClick: async () => { await pasteFromSystem({ ...DEFAULT_PASTE, what: 'values' }); await pasteFromSystem({ ...DEFAULT_PASTE, what: 'formats' }); } },
    { separator: true },
    { label: 'Formatting', onClick: p({ what: 'formats' }) },
    { label: 'Paste Link', onClick: p({ link: true }) },
    { separator: true },
    { label: 'Paste Special...', onClick: () => openDialog('pasteSpecial'), shortcut: 'Ctrl+Alt+V' },
  ];
}

export function bordersMenu(): MenuItem[] {
  const b = (p: BorderPreset) => () => {
    lastBorder = p;
    applyBorder(p);
  };
  return [
    { label: 'Bottom Border', onClick: b('bottom') },
    { label: 'Top Border', onClick: b('top') },
    { label: 'Left Border', onClick: b('left') },
    { label: 'Right Border', onClick: b('right') },
    { separator: true },
    { label: 'No Border', onClick: b('none') },
    { label: 'All Borders', onClick: b('all') },
    { label: 'Outside Borders', onClick: b('outside') },
    { label: 'Thick Outside Borders', onClick: b('thickOutside') },
    { separator: true },
    { label: 'Bottom Double Border', onClick: b('bottomDouble') },
    { label: 'Thick Bottom Border', onClick: b('thickBottom') },
    { label: 'Top and Bottom Border', onClick: b('topBottom') },
    { label: 'Top and Thick Bottom Border', onClick: b('topThickBottom') },
    { label: 'Top and Double Bottom Border', onClick: b('topDoubleBottom') },
    { separator: true },
    { label: 'More Borders...', onClick: () => openDialog('formatCells', { tab: 'Border' }) },
  ];
}

function numberFormatOptions(): { name: string; fmt: string }[] {
  return [
    { name: 'General', fmt: 'General' },
    { name: 'Number', fmt: '0.00' },
    { name: 'Currency', fmt: '"$"#,##0.00' },
    { name: 'Accounting', fmt: FMT.accounting },
    { name: 'Short Date', fmt: 'm/d/yyyy' },
    { name: 'Long Date', fmt: 'dddd, mmmm d, yyyy' },
    { name: 'Time', fmt: 'h:mm:ss AM/PM' },
    { name: 'Percentage', fmt: '0.00%' },
    { name: 'Fraction', fmt: '# ?/?' },
    { name: 'Scientific', fmt: '0.00E+00' },
    { name: 'Text', fmt: '@' },
  ];
}

export function formatCategoryName(fmt: string | undefined): string {
  if (!fmt || fmt === 'General') return 'General';
  const hit = numberFormatOptions().find((o) => o.fmt === fmt);
  if (hit) return hit.name;
  if (fmt === '@') return 'Text';
  if (fmt.includes('₹')) return 'Currency';
  if (/_\(.*\*/.test(fmt) || /\* /.test(fmt)) return 'Accounting';
  if (fmt.includes('%')) return 'Percentage';
  if (/E\+/.test(fmt)) return 'Scientific';
  if (/[$€£¥]/.test(fmt)) return 'Currency';
  if (/h|s/i.test(fmt.replace(/"[^"]*"/g, '')) && !/[dy]/i.test(fmt)) return 'Time';
  if (/[dmy]/i.test(fmt.replace(/"[^"]*"/g, ''))) return 'Date';
  if (/\?\/\?/.test(fmt)) return 'Fraction';
  if (/^[#0,.]+$/.test(fmt.split(';')[0].replace(/_\)|\\\(|\\\)|\[Red\]/g, ''))) return 'Number';
  return 'Custom';
}

function CFGallery({ kind, close }: { kind: 'bars' | 'scales' | 'icons'; close: () => void }) {
  if (kind === 'bars') {
    return (
      <div className="p-2 w-[210px]">
        <div className="text-[11px] font-semibold mb-1">Gradient Fill</div>
        <div className="grid grid-cols-3 gap-1.5">
          {DATA_BAR_COLORS.map((c) => (
            <button key={c} className="xl-gallery-item h-9 border flex flex-col justify-around p-1" style={{ borderColor: 'var(--border)' }} onClick={() => { addDataBar(c, true); close(); }} title={`${c} Data Bar`}>
              {[0.9, 0.6, 0.35].map((w) => (
                <span key={w} className="block h-[6px]" style={{ width: `${w * 100}%`, background: `linear-gradient(90deg, ${c}, #fff)` }} />
              ))}
            </button>
          ))}
        </div>
        <div className="text-[11px] font-semibold mt-2 mb-1">Solid Fill</div>
        <div className="grid grid-cols-3 gap-1.5">
          {DATA_BAR_COLORS.map((c) => (
            <button key={c} className="xl-gallery-item h-9 border flex flex-col justify-around p-1" style={{ borderColor: 'var(--border)' }} onClick={() => { addDataBar(c, false); close(); }}>
              {[0.9, 0.6, 0.35].map((w) => (
                <span key={w} className="block h-[6px]" style={{ width: `${w * 100}%`, background: c }} />
              ))}
            </button>
          ))}
        </div>
        <button className="xl-menu-item !mx-0 w-full mt-2" onClick={() => { close(); openDialog('cfRule', { type: 'dataBar' }); }}>More Rules...</button>
      </div>
    );
  }
  if (kind === 'scales') {
    return (
      <div className="p-2 w-[200px]">
        <div className="grid grid-cols-4 gap-1.5">
          {COLOR_SCALES.map((s) => (
            <button key={s.name} title={s.name} className="xl-gallery-item h-9 border grid grid-cols-1" style={{ borderColor: 'var(--border)' }} onClick={() => { addColorScale(s.colors); close(); }}>
              {[...s.colors].reverse().map((c, i) => (
                <span key={i} className="block" style={{ background: c }} />
              ))}
            </button>
          ))}
        </div>
        <button className="xl-menu-item !mx-0 w-full mt-2" onClick={() => { close(); openDialog('cfRule', { type: 'colorScale' }); }}>More Rules...</button>
      </div>
    );
  }
  return (
    <div className="p-2 w-[230px]">
      <div className="grid grid-cols-2 gap-1">
        {ICON_SETS.map((set) => (
          <button key={set} title={set} className="xl-gallery-item h-7 flex items-center gap-1 px-1" onClick={() => { addIconSet(set); close(); }}>
            <IconSetPreview set={set} />
          </button>
        ))}
      </div>
      <button className="xl-menu-item !mx-0 w-full mt-2" onClick={() => { close(); openDialog('cfRule', { type: 'iconSet' }); }}>More Rules...</button>
    </div>
  );
}

export function IconSetPreview({ set }: { set: string }) {
  const ref = useRef<HTMLCanvasElement>(null);
  const n = parseInt(set, 10) || 3;
  useEffect(() => {
    const c = ref.current;
    if (!c) return;
    const ctx = c.getContext('2d')!;
    const dpr = window.devicePixelRatio || 1;
    c.width = n * 18 * dpr;
    c.height = 16 * dpr;
    ctx.scale(dpr, dpr);
    for (let i = 0; i < n; i++) drawIcon(ctx, set, n - 1 - i, n, i * 18 + 1, 1, 14);
  }, [set, n]);
  return <canvas ref={ref} style={{ width: n * 18, height: 16 }} />;
}

export function TableStyleGallery({ onPick, close }: { onPick: (s: string) => void; close: () => void }) {
  const fake: TableDef = { id: '', name: '', range: { r1: 0, r2: 4, c1: 0, c2: 4 }, headerRow: true, totalRow: false, style: '', bandedRows: true, bandedCols: false, firstCol: false, lastCol: false, showFilterButton: false, columns: [] };
  const groups: [string, string[]][] = [
    ['Light', TABLE_STYLE_NAMES.filter((s) => s.includes('Light'))],
    ['Medium', TABLE_STYLE_NAMES.filter((s) => s.includes('Medium'))],
    ['Dark', TABLE_STYLE_NAMES.filter((s) => s.includes('Dark'))],
  ];
  return (
    <div className="p-2 w-[430px] max-h-[430px] overflow-auto">
      {groups.map(([g, list]) => (
        <div key={g}>
          <div className="text-[11px] font-semibold my-1">{g}</div>
          <div className="grid grid-cols-7 gap-1.5">
            {list.map((s) => (
              <button key={s} title={s.replace('TableStyle', 'Table Style ').replace(/(\d+)/, ' $1')} className="xl-gallery-item border p-[2px]" style={{ borderColor: 'var(--border)' }} onClick={() => { onPick(s); close(); }}>
                <div className="grid grid-cols-4" style={{ width: 48, height: 34 }}>
                  {Array.from({ length: 20 }, (_, k) => {
                    const r = Math.floor(k / 4);
                    const c = k % 4;
                    const look = tableLook({ ...fake, style: s, range: { r1: 0, r2: 4, c1: 0, c2: 3 } }, r, c);
                    return <span key={k} style={{ background: look.fill ?? '#fff', borderTop: look.bTop ? `1px solid ${look.bTop.color}` : undefined, borderBottom: look.bBottom ? `1px solid ${look.bBottom.color}` : undefined }} />;
                  })}
                </div>
              </button>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}

function CellStyleGallery({ close }: { close: () => void }) {
  const groups = ['Good, Bad and Neutral', 'Data and Model', 'Titles and Headings', 'Themed Cell Styles', 'Number Format'] as const;
  return (
    <div className="p-2 w-[520px] max-h-[440px] overflow-auto">
      {groups.map((g) => (
        <div key={g}>
          <div className="xl-menu-header !px-1 my-1 text-[11px]">{g}</div>
          <div className="flex flex-wrap gap-1.5">
            {CELL_STYLES.filter((s) => s.group === g).map((s) => (
              <button
                key={s.name}
                className="xl-gallery-item h-[26px] w-[92px] text-[12px] px-1 text-left truncate border"
                style={{
                  background: (s.patch.fillColor as string) ?? '#fff',
                  color: (s.patch.fontColor as string) ?? '#000',
                  fontWeight: s.patch.bold ? 700 : 400,
                  fontStyle: s.patch.italic ? 'italic' : 'normal',
                  fontSize: s.patch.fontSize ? Math.min(15, (s.patch.fontSize as number) * 0.95) : 12,
                  borderColor: 'var(--border)',
                  borderBottom: s.patch.bBottom ? `2px ${s.patch.bBottom.style === 'double' ? 'double' : 'solid'} ${s.patch.bBottom.color}` : undefined,
                }}
                onClick={() => {
                  applyStyle(s.patch, 'Cell Style');
                  close();
                }}
              >
                {s.name}
              </button>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}

export function cfMenu(): MenuItem[] {
  const hl = (kind: string) => () => openDialog('cfQuick', { kind });
  return [
    {
      label: 'Highlight Cells Rules',
      submenu: [
        { label: 'Greater Than...', onClick: hl('greaterThan') },
        { label: 'Less Than...', onClick: hl('lessThan') },
        { label: 'Between...', onClick: hl('between') },
        { label: 'Equal To...', onClick: hl('equal') },
        { label: 'Text that Contains...', onClick: hl('containsText') },
        { label: 'A Date Occurring...', onClick: hl('timePeriod') },
        { label: 'Duplicate Values...', onClick: hl('duplicateValues') },
        { separator: true },
        { label: 'More Rules...', onClick: () => openDialog('cfRule', { type: 'cellIs' }) },
      ],
    },
    {
      label: 'Top/Bottom Rules',
      submenu: [
        { label: 'Top 10 Items...', onClick: hl('top10') },
        { label: 'Top 10%...', onClick: hl('top10pct') },
        { label: 'Bottom 10 Items...', onClick: hl('bottom10') },
        { label: 'Bottom 10%...', onClick: hl('bottom10pct') },
        { label: 'Above Average...', onClick: hl('aboveAverage') },
        { label: 'Below Average...', onClick: hl('belowAverage') },
        { separator: true },
        { label: 'More Rules...', onClick: () => openDialog('cfRule', { type: 'top10' }) },
      ],
    },
    { separator: true },
    { label: 'Data Bars', submenu: [{ render: (close) => <CFGallery kind="bars" close={close} /> }] },
    { label: 'Color Scales', submenu: [{ render: (close) => <CFGallery kind="scales" close={close} /> }] },
    { label: 'Icon Sets', submenu: [{ render: (close) => <CFGallery kind="icons" close={close} /> }] },
    { separator: true },
    { label: 'New Rule...', onClick: () => openDialog('cfRule', {}) },
    {
      label: 'Clear Rules',
      submenu: [
        { label: 'Clear Rules from Selected Cells', onClick: () => clearCFRules('selection') },
        { label: 'Clear Rules from Entire Sheet', onClick: () => clearCFRules('sheet') },
      ],
    },
    { label: 'Manage Rules...', onClick: () => openDialog('cfManager') },
  ];
}

export function formatAsTable(style: string): void {
  const guess = tableSourceGuess();
  openDialog('createTable', { style, range: guess?.range, header: guess?.header ?? true });
}

export function formatMenu(): MenuItem[] {
  const st = S();
  const sheet = st.wb.activeSheet;
  const locked = st.wb.styles.get(sheet.styleIdAt(st.sel.active.r, st.sel.active.c)).locked !== false;
  return [
    { label: 'Cell Size', disabled: true, bold: true },
    { label: 'Row Height...', onClick: () => openDialog('rowHeight') },
    { label: 'AutoFit Row Height', onClick: () => autoFitRows() },
    { label: 'Column Width...', onClick: () => openDialog('colWidth') },
    { label: 'AutoFit Column Width', onClick: () => autoFitColumns() },
    { label: 'Default Width...', onClick: () => openDialog('colWidth', { standard: true }) },
    { separator: true },
    { label: 'Visibility', disabled: true, bold: true },
    {
      label: 'Hide & Unhide',
      submenu: [
        { label: 'Hide Rows', onClick: () => hideRowsCols('row', true), shortcut: 'Ctrl+9' },
        { label: 'Hide Columns', onClick: () => hideRowsCols('col', true), shortcut: 'Ctrl+0' },
        { label: 'Hide Sheet', onClick: () => hideSheet(sheet.id) },
        { separator: true },
        { label: 'Unhide Rows', onClick: () => hideRowsCols('row', false), shortcut: 'Ctrl+Shift+9' },
        { label: 'Unhide Columns', onClick: () => hideRowsCols('col', false) },
        { label: 'Unhide Sheet...', onClick: () => openDialog('unhideSheet'), disabled: !st.wb.sheets.some((s) => s.visibility === 'hidden') },
      ],
    },
    { separator: true },
    { label: 'Organize Sheets', disabled: true, bold: true },
    { label: 'Rename Sheet', onClick: () => window.dispatchEvent(new CustomEvent('xl-rename-sheet', { detail: sheet.id })) },
    { label: 'Move or Copy Sheet...', onClick: () => openDialog('moveCopySheet') },
    { label: 'Tab Color', submenu: [{ render: (close) => <ColorPalette onPick={(c) => setTabColor(sheet.id, c ?? undefined)} close={close} noneLabel="No Color" /> }] },
    { separator: true },
    { label: 'Protection', disabled: true, bold: true },
    sheet.protection
      ? { label: 'Unprotect Sheet...', onClick: () => (sheet.protection?.passwordHash ? openDialog('unprotectSheet') : unprotectSheet('')) }
      : { label: 'Protect Sheet...', onClick: () => openDialog('protectSheet') },
    { label: 'Lock Cell', checked: locked, onClick: () => (sheet.protection ? protectedAlert() : toggleLockCell()) },
    { label: 'Format Cells...', onClick: () => openDialog('formatCells'), shortcut: 'Ctrl+1' },
  ];
}

export function HomeTab() {
  useStore((s) => s.rev);
  const sel = useStore((s) => s.sel);
  const painter = useStore((s) => s.painter);
  const st = activeStyle();
  const sheet = S().wb.activeSheet;
  const [fillColor, setFillColor] = useState(lastFill);
  const [fontColor, setFontColor] = useState(lastFont);
  const merged = sheet.merges.some((m) => {
    const rg = primaryRange(sel);
    return m.r1 === rg.r1 && m.c1 === rg.c1;
  });
  const val = getComputed(sheet, sel.active.r, sel.active.c);
  const sample = typeof val === 'number' ? val : null;
  const nfOpts = numberFormatOptions();
  return (
    <>
      <Group label="Clipboard" launcher={() => openDialog('pasteSpecial')} title="Paste Special">
        <LargeButton icon={<PasteIcon />} label="Paste" onClick={() => pasteFromSystem()} menu={pasteMenu} testId="btn-paste" />
        <Col>
          <SmallButton icon={<Scissors size={15} />} label="Cut" onClick={() => copyToSystem(true)} title="Cut (Ctrl+X)" />
          <SmallButton icon={<Copy size={15} />} label="Copy" onClick={() => copyToSystem(false)} title="Copy (Ctrl+C)" menu={[{ label: 'Copy', onClick: () => copyToSystem(false) }, { label: 'Copy as Picture...', disabled: true }]} />
          <SmallButton icon={<Paintbrush size={15} />} label="Format Painter" checked={!!painter} onClick={() => startFormatPainter(false)} title="Format Painter" />
        </Col>
      </Group>

      <Group label="Font" launcher={() => openDialog('formatCells', { tab: 'Font' })}>
        <Col>
          <Row>
            <Combo value={st.fontName ?? 'Calibri'} options={FONTS} onCommit={(v) => applyStyle({ fontName: v }, 'Font')} width={128} title="Font" testId="font-name" renderOption={(o) => <span style={{ fontFamily: `"${o}"`, fontSize: 13 }}>{o}</span>} />
            <Combo value={String(st.fontSize ?? 11)} options={FONT_SIZES.map(String)} onCommit={(v) => { const n = parseFloat(v); if (n > 0 && n <= 409) applyStyle({ fontSize: n }, 'Font Size'); }} width={46} title="Font Size" testId="font-size" />
            <SmallButton icon={<AArrowUp size={16} />} title="Increase Font Size (Ctrl+Shift+>)" onClick={() => growFont(1)} />
            <SmallButton icon={<AArrowDown size={16} />} title="Decrease Font Size (Ctrl+Shift+<)" onClick={() => growFont(-1)} />
          </Row>
          <Row className="mt-[3px]">
            <SmallButton icon={<Bold size={15} strokeWidth={2.6} />} title="Bold (Ctrl+B)" checked={!!st.bold} onClick={() => toggleStyle('bold')} testId="btn-bold" />
            <SmallButton icon={<Italic size={15} />} title="Italic (Ctrl+I)" checked={!!st.italic} onClick={() => toggleStyle('italic')} testId="btn-italic" />
            <SmallButton icon={<Underline size={15} />} title="Underline (Ctrl+U)" checked={!!st.underline} onClick={() => toggleUnderline('single')} menu={[{ label: 'Underline', onClick: () => toggleUnderline('single') }, { label: 'Double Underline', onClick: () => toggleUnderline('double') }]} testId="btn-underline" />
            <div className="w-px h-4 mx-1" style={{ background: 'var(--border)' }} />
            <SmallButton icon={<BordersIcon />} title="Borders" onClick={() => applyBorder(lastBorder)} menu={bordersMenu} testId="btn-borders" />
            <div className="w-px h-4 mx-1" style={{ background: 'var(--border)' }} />
            <SmallButton
              icon={<ColorIcon icon={<FillBucketIcon />} color={fillColor} />}
              title="Fill Color"
              onClick={() => applyStyle({ fillColor }, 'Fill Color')}
              panel={(close) => <ColorPalette close={close} noneLabel="No Fill" onPick={(c) => { if (c) { lastFill = c; setFillColor(c); } applyStyle({ fillColor: c }, 'Fill Color'); }} />}
              testId="btn-fill"
            />
            <SmallButton
              icon={<ColorIcon icon={<FontColorIcon />} color={fontColor} />}
              title="Font Color"
              onClick={() => applyStyle({ fontColor }, 'Font Color')}
              panel={(close) => <ColorPalette close={close} autoLabel="Automatic" onPick={(c) => { if (c) { lastFont = c; setFontColor(c); } applyStyle({ fontColor: c }, 'Font Color'); }} />}
              testId="btn-fontcolor"
            />
            <SmallButton icon={<Strikethrough size={15} />} title="Strikethrough (Ctrl+5)" checked={!!st.strike} onClick={() => toggleStyle('strike')} />
          </Row>
        </Col>
      </Group>

      <Group label="Alignment" launcher={() => openDialog('formatCells', { tab: 'Alignment' })}>
        <Col>
          <Row>
            <SmallButton icon={<ArrowUpToLine size={15} />} title="Top Align" checked={st.vAlign === 'top'} onClick={() => setVAlign('top')} />
            <SmallButton icon={<AlignVerticalJustifyCenter size={15} />} title="Middle Align" checked={st.vAlign === 'middle'} onClick={() => setVAlign('middle')} />
            <SmallButton icon={<ArrowDownToLine size={15} />} title="Bottom Align" checked={!st.vAlign || st.vAlign === 'bottom'} onClick={() => setVAlign('bottom')} />
            <SmallButton
              icon={<RotateCcw size={15} />}
              title="Orientation"
              menu={[
                { label: 'Angle Counterclockwise', onClick: () => setRotation(45) },
                { label: 'Angle Clockwise', onClick: () => setRotation(135) },
                { label: 'Vertical Text', onClick: () => setRotation('vertical') },
                { label: 'Rotate Text Up', onClick: () => setRotation(90) },
                { label: 'Rotate Text Down', onClick: () => setRotation(180) },
                { separator: true },
                { label: 'Format Cell Alignment', onClick: () => openDialog('formatCells', { tab: 'Alignment' }) },
              ]}
            />
            <SmallButton icon={<WrapText size={15} />} label="Wrap Text" checked={!!st.wrap} onClick={toggleWrap} testId="btn-wrap" />
          </Row>
          <Row className="mt-[3px]">
            <SmallButton icon={<AlignLeft size={15} />} title="Align Left" checked={st.hAlign === 'left'} onClick={() => setHAlign('left')} />
            <SmallButton icon={<AlignCenter size={15} />} title="Center" checked={st.hAlign === 'center'} onClick={() => setHAlign('center')} testId="btn-center" />
            <SmallButton icon={<AlignRight size={15} />} title="Align Right" checked={st.hAlign === 'right'} onClick={() => setHAlign('right')} />
            <SmallButton icon={<IndentDecrease size={15} />} title="Decrease Indent" onClick={() => changeIndent(-1)} />
            <SmallButton icon={<IndentIncrease size={15} />} title="Increase Indent" onClick={() => changeIndent(1)} />
            <SmallButton
              icon={<MergeIcon />}
              label="Merge & Center"
              checked={merged}
              onClick={() => mergeCells('center')}
              menu={[
                { label: 'Merge & Center', onClick: () => mergeCells('center') },
                { label: 'Merge Across', onClick: () => mergeCells('across') },
                { label: 'Merge Cells', onClick: () => mergeCells('merge') },
                { label: 'Unmerge Cells', onClick: () => mergeCells('unmerge') },
              ]}
              testId="btn-merge"
            />
          </Row>
        </Col>
      </Group>

      <Group label="Number" launcher={() => openDialog('formatCells', { tab: 'Number' })}>
        <Col>
          <Row>
            <Combo
              value={formatCategoryName(st.numFmt)}
              options={[...nfOpts.map((o) => o.name), 'More Number Formats...']}
              onCommit={(v) => {
                if (v === 'More Number Formats...') return openDialog('formatCells', { tab: 'Number' });
                const o = nfOpts.find((x) => x.name === v);
                if (o) setNumFmt(o.fmt);
              }}
              width={136}
              title="Number Format"
              testId="numfmt"
              renderOption={(o) => {
                const f = nfOpts.find((x) => x.name === o);
                return (
                  <div className="flex flex-col py-0.5">
                    <span className="font-semibold">{o}</span>
                    {f && sample !== null && <span className="text-[11px] opacity-70">{formatValue(sample, f.fmt).text}</span>}
                  </div>
                );
              }}
            />
          </Row>
          <Row className="mt-[3px]">
            <SmallButton
              icon={<span className="text-[13px] font-semibold w-4 text-center">$</span>}
              title="Accounting Number Format"
              onClick={() => setNumFmt(FMT.accounting)}
              menu={[
                { label: '$ English (United States)', onClick: () => setNumFmt(FMT.accounting) },
                { label: '£ English (United Kingdom)', onClick: () => setNumFmt('_-[$£-809]* #,##0.00_-;-[$£-809]* #,##0.00_-;_-[$£-809]* "-"??_-;_-@_-') },
                { label: '€ Euro (€ 123)', onClick: () => setNumFmt('_-[$€-x-euro2] * #,##0.00_-;-[$€-x-euro2] * #,##0.00_-;_-[$€-x-euro2] * "-"??_-;_-@_-') },
                { label: '¥ Chinese (PRC)', onClick: () => setNumFmt('_-[$¥-804]* #,##0.00_-;-[$¥-804]* #,##0.00_-;_-[$¥-804]* "-"??_-;_-@_-') },
                { label: '₹ Indian Rupee (lakh/crore)', onClick: () => setNumFmt(FMT.inrAccounting) },
                { separator: true },
                { label: 'More Accounting Formats...', onClick: () => openDialog('formatCells', { tab: 'Number', category: 'Accounting' }) },
              ]}
            />
            <SmallButton icon={<Percent size={14} />} title="Percent Style (Ctrl+Shift+%)" onClick={() => setNumFmt('0%')} />
            <SmallButton icon={<span className="text-[15px] font-bold w-4 text-center leading-none">,</span>} title="Comma Style" onClick={commaStyle} />
            <SmallButton icon={<span className="text-[10px] font-semibold tracking-tighter">←.0<br /></span>} title="Increase Decimal" onClick={() => changeDecimals(1)} />
            <SmallButton icon={<span className="text-[10px] font-semibold tracking-tighter">.00→</span>} title="Decrease Decimal" onClick={() => changeDecimals(-1)} />
          </Row>
        </Col>
      </Group>

      <Group label="Styles">
        <LargeButton icon={<CondFormatIcon />} label={'Conditional\nFormatting'} menu={cfMenu} testId="btn-cf" />
        <LargeButton icon={<FormatTableIcon />} label={'Format as\nTable'} panel={(close) => <TableStyleGallery close={close} onPick={formatAsTable} />} />
        <LargeButton icon={<CellStylesIcon />} label={'Cell\nStyles'} panel={(close) => <CellStyleGallery close={close} />} />
      </Group>

      <Group label="Cells">
        <LargeButton
          icon={<InsertCellsIcon />}
          label="Insert"
          onClick={() => insertCellsDialog()}
          menu={[
            { label: 'Insert Cells...', onClick: insertCellsDialog },
            { label: 'Insert Sheet Rows', onClick: () => insertRows() },
            { label: 'Insert Sheet Columns', onClick: () => insertCols() },
            { label: 'Insert Sheet', onClick: () => addSheet() },
          ]}
        />
        <LargeButton
          icon={<DeleteCellsIcon />}
          label="Delete"
          onClick={() => deleteCellsDialog()}
          menu={[
            { label: 'Delete Cells...', onClick: deleteCellsDialog },
            { label: 'Delete Sheet Rows', onClick: deleteRows },
            { label: 'Delete Sheet Columns', onClick: deleteCols },
            { label: 'Delete Sheet', onClick: () => deleteSheet() },
          ]}
        />
        <LargeButton icon={<FormatCellsIcon />} label="Format" menu={formatMenu} />
      </Group>

      <Group label="Editing">
        <Col>
          <SmallButton
            icon={<AutoSumIcon />}
            label="AutoSum"
            onClick={() => autoSum('SUM')}
            menu={[
              { label: 'Sum', onClick: () => autoSum('SUM') },
              { label: 'Average', onClick: () => autoSum('AVERAGE') },
              { label: 'Count Numbers', onClick: () => autoSum('COUNT') },
              { label: 'Max', onClick: () => autoSum('MAX') },
              { label: 'Min', onClick: () => autoSum('MIN') },
              { separator: true },
              { label: 'More Functions...', onClick: () => openDialog('insertFunction') },
            ]}
            testId="btn-autosum"
          />
          <SmallButton
            icon={<ChevronsDown size={15} className="text-[#2B7CD3]" />}
            label="Fill"
            menu={[
              { label: 'Down', onClick: () => fillDirection('down'), shortcut: 'Ctrl+D' },
              { label: 'Right', onClick: () => fillDirection('right'), shortcut: 'Ctrl+R' },
              { label: 'Up', onClick: () => fillDirection('up') },
              { label: 'Left', onClick: () => fillDirection('left') },
              { separator: true },
              { label: 'Series...', onClick: () => openDialog('series') },
              { label: 'Flash Fill', onClick: flashFill, shortcut: 'Ctrl+E' },
            ]}
          />
          <SmallButton
            icon={<Eraser size={15} className="text-[#C00000]" />}
            label="Clear"
            menu={[
              { label: 'Clear All', onClick: () => clearSelection('all') },
              { label: 'Clear Formats', onClick: () => clearSelection('formats') },
              { label: 'Clear Contents', onClick: () => clearSelection('contents') },
              { label: 'Clear Comments and Notes', onClick: () => clearSelection('comments') },
              { separator: true },
              { label: 'Clear Hyperlinks', onClick: () => clearSelection('hyperlinks') },
            ]}
          />
        </Col>
        <LargeButton
          icon={<SortFilterIcon />}
          label={'Sort &\nFilter'}
          testId="btn-sortfilter"
          menu={() => [
            { label: 'Sort A to Z', onClick: () => quickSort(false), testId: 'sort-asc' },
            { label: 'Sort Z to A', onClick: () => quickSort(true), testId: 'sort-desc' },
            { label: 'Custom Sort...', icon: <ArrowDownNarrowWide size={14} />, onClick: () => openDialog('sort') },
            { separator: true },
            { label: 'Filter', checked: !!S().wb.activeSheet.autoFilter, onClick: toggleAutoFilter, shortcut: 'Ctrl+Shift+L' },
            { label: 'Clear', onClick: clearAllFilters, disabled: !S().wb.activeSheet.autoFilter },
            { label: 'Reapply', onClick: reapplyFilter, disabled: !S().wb.activeSheet.autoFilter },
          ]}
        />
        <LargeButton
          icon={<FindSelectIcon />}
          label={'Find &\nSelect'}
          menu={[
            { label: 'Find...', onClick: () => openDialog('findReplace', { tab: 'find' }), shortcut: 'Ctrl+F' },
            { label: 'Replace...', onClick: () => openDialog('findReplace', { tab: 'replace' }), shortcut: 'Ctrl+H' },
            { label: 'Go To...', onClick: () => openDialog('goto'), shortcut: 'Ctrl+G' },
            { label: 'Go To Special...', onClick: () => openDialog('gotoSpecial') },
            { separator: true },
            { label: 'Formulas', onClick: () => goToSpecial({ kind: 'formulas' }) },
            { label: 'Notes', onClick: () => goToSpecial({ kind: 'notes' }) },
            { label: 'Conditional Formatting', onClick: () => goToSpecial({ kind: 'conditionalFormats' }) },
            { label: 'Constants', onClick: () => goToSpecial({ kind: 'constants' }) },
            { label: 'Data Validation', onClick: () => goToSpecial({ kind: 'validation' }) },
          ]}
        />
      </Group>
    </>
  );
}

export { setState };

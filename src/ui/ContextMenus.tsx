'use client';
import { ClipboardPaste, Copy, Link, MessageSquarePlus, Scissors, StickyNote, Tag, Table2, Hash, ArrowRightLeft, Paintbrush, Link2, SquareFunction, Eraser } from 'lucide-react';
import { primaryRange } from '../model/selection';
import { copyToSystem, DEFAULT_PASTE, pasteFromSystem, PasteOptions } from '../state/actions/clipboard';
import { clearSelection } from '../state/actions/edit';
import { hideRowsCols } from '../state/actions/format';
import { clearAllFilters, displayTextAt, reapplyFilter, setColumnFilter, sortByColor, quickSort, toggleAutoFilter } from '../state/actions/sortFilter';
import { deleteCellsDialog, deleteCols, deleteRows, insertCellsDialog, insertCols, insertRows } from '../state/actions/structure';
import { openDialog, S, setState, useStore } from '../state/store';
import { validationAt } from '../state/validation';
import { Menu, MenuItem } from './Menu';

function PasteOptionsRow({ close }: { close: () => void }) {
  const p = (o: Partial<PasteOptions>) => () => {
    close();
    pasteFromSystem({ ...DEFAULT_PASTE, ...o });
  };
  const btn = (title: string, icon: React.ReactNode, fn: () => void) => (
    <button title={title} className="w-7 h-7 flex items-center justify-center rounded border hover:bg-[var(--menu-hover)]" style={{ borderColor: 'var(--border)' }} onClick={fn}>
      {icon}
    </button>
  );
  return (
    <div className="px-2 py-1">
      <div className="flex items-center gap-1.5 text-[12px] pl-[30px] mb-1">
        <span className="font-semibold">Paste Options:</span>
      </div>
      <div className="flex gap-1 pl-[30px]">
        {btn('Paste (P)', <ClipboardPaste size={15} />, p({}))}
        {btn('Values (V)', <Hash size={15} />, p({ what: 'values' }))}
        {btn('Formulas (F)', <SquareFunction size={15} />, p({ what: 'formulas' }))}
        {btn('Transpose (T)', <ArrowRightLeft size={15} />, p({ transpose: true }))}
        {btn('Formatting (R)', <Paintbrush size={15} />, p({ what: 'formats' }))}
        {btn('Paste Link (N)', <Link2 size={15} />, p({ link: true }))}
      </div>
    </div>
  );
}

function commonTop(): MenuItem[] {
  return [
    { label: 'Cut', icon: <Scissors size={15} />, onClick: () => copyToSystem(true), shortcut: '' },
    { label: 'Copy', icon: <Copy size={15} />, onClick: () => copyToSystem(false) },
    { render: (close) => <PasteOptionsRow close={close} /> },
    { label: 'Paste Special...', onClick: () => openDialog('pasteSpecial') },
  ];
}

export function cellMenuItems(): MenuItem[] {
  const st = S();
  const sheet = st.wb.activeSheet;
  const a = st.sel.active;
  const af = sheet.autoFilter;
  const style = st.wb.styles.get(sheet.styleIdAt(a.r, a.c));
  const cell = sheet.getCell(a.r, a.c);
  const inFilter = af && a.r > af.range.r1 && a.r <= af.range.r2 && a.c >= af.range.c1 && a.c <= af.range.c2;
  const filterBy = (fn: () => void) => () => {
    if (!sheet.autoFilter) toggleAutoFilter();
    fn();
  };
  const dv = validationAt(sheet, a.r, a.c);
  return [
    ...commonTop(),
    { separator: true },
    { label: 'Insert...', onClick: insertCellsDialog },
    { label: 'Delete...', onClick: deleteCellsDialog },
    { label: 'Clear Contents', icon: <Eraser size={15} />, onClick: () => clearSelection('contents') },
    { separator: true },
    {
      label: 'Filter',
      submenu: [
        { label: 'Reapply', disabled: !af, onClick: reapplyFilter },
        { label: `Clear Filter From "${af ? displayTextAt(sheet, af.range.r1, a.c) : ''}"`, disabled: !inFilter || !af?.filters[a.c], onClick: () => setColumnFilter(a.c, null) },
        { separator: true },
        { label: "Filter by Selected Cell's Value", onClick: filterBy(() => setColumnFilter(a.c, { type: 'values', values: [displayTextAt(sheet, a.r, a.c)], blanks: displayTextAt(sheet, a.r, a.c) === '' })) },
        { label: "Filter by Selected Cell's Color", disabled: !style.fillColor, onClick: filterBy(() => setColumnFilter(a.c, { type: 'color', color: style.fillColor! })) },
        { label: "Filter by Selected Cell's Font Color", disabled: !style.fontColor, onClick: filterBy(() => setColumnFilter(a.c, { type: 'color', color: style.fontColor!, font: true })) },
        { separator: true },
        { label: 'Clear All Filters', disabled: !af, onClick: clearAllFilters },
      ],
    },
    {
      label: 'Sort',
      submenu: [
        { label: 'Sort A to Z', onClick: () => quickSort(false) },
        { label: 'Sort Z to A', onClick: () => quickSort(true) },
        { label: 'Put Selected Cell Color On Top', disabled: !style.fillColor || !af, onClick: () => sortByColor(a.c, style.fillColor!, false) },
        { label: 'Put Selected Font Color On Top', disabled: !style.fontColor || !af, onClick: () => sortByColor(a.c, style.fontColor!, true) },
        { separator: true },
        { label: 'Custom Sort...', onClick: () => openDialog('sort') },
      ],
    },
    { label: 'Get Data from Table/Range...', icon: <Table2 size={15} />, onClick: () => openDialog('createTable') },
    { separator: true },
    { label: 'New Comment', icon: <MessageSquarePlus size={15} />, onClick: () => openDialog('note', { edit: true, threaded: true }) },
    cell?.note ? { label: 'Edit Note', icon: <StickyNote size={15} />, onClick: () => openDialog('note', { edit: true }) } : { label: 'New Note', icon: <StickyNote size={15} />, onClick: () => openDialog('note', { edit: true }) },
    ...(cell?.note ? [{ label: 'Delete Note', onClick: () => import('../state/actions/data').then((m) => m.setNote(a.r, a.c, undefined)) }] : []),
    { separator: true },
    { label: 'Format Cells...', onClick: () => openDialog('formatCells'), shortcut: 'Ctrl+1' },
    { label: 'Pick From Drop-down List...', onClick: () => setState({ listMenu: { ...a } }), disabled: !(dv && dv.type === 'list') && !hasColumnValues() },
    { label: 'Define Name...', icon: <Tag size={15} />, onClick: () => openDialog('newName') },
    cell?.link
      ? { label: 'Link', icon: <Link size={15} />, submenu: [
          { label: 'Edit Link...', onClick: () => openDialog('hyperlink') },
          { label: 'Open Link', onClick: () => window.open(cell.link!, '_blank', 'noopener') },
          { label: 'Remove Link', onClick: () => import('../state/actions/data').then((m) => m.setHyperlink(undefined)) },
        ] }
      : { label: 'Link', icon: <Link size={15} />, onClick: () => openDialog('hyperlink'), shortcut: 'Ctrl+K' },
  ];
}

function hasColumnValues(): boolean {
  return true;
}

function headerMenu(kind: 'row' | 'col'): MenuItem[] {
  return [
    ...commonTop(),
    { separator: true },
    { label: 'Insert', onClick: () => (kind === 'row' ? insertRows() : insertCols()) },
    { label: 'Delete', onClick: () => (kind === 'row' ? deleteRows() : deleteCols()) },
    { label: 'Clear Contents', onClick: () => clearSelection('contents') },
    { separator: true },
    { label: 'Format Cells...', onClick: () => openDialog('formatCells') },
    kind === 'row' ? { label: 'Row Height...', onClick: () => openDialog('rowHeight') } : { label: 'Column Width...', onClick: () => openDialog('colWidth') },
    { label: 'Hide', onClick: () => hideRowsCols(kind, true) },
    { label: 'Unhide', onClick: () => hideRowsCols(kind, false) },
  ];
}

export function ContextMenus() {
  const menu = useStore((s) => s.menu);
  useStore((s) => s.sel);
  if (!menu || menu.kind === 'tab' || menu.kind === 'chart') return null;
  const close = () => setState({ menu: null });
  const items = menu.kind === 'cell' ? cellMenuItems() : menu.kind === 'row' ? headerMenu('row') : menu.kind === 'col' ? headerMenu('col') : [
    ...commonTop(),
    { separator: true },
    { label: 'Format Cells...', onClick: () => openDialog('formatCells') },
    { label: 'Row Height...', onClick: () => openDialog('rowHeight') },
    { label: 'Column Width...', onClick: () => openDialog('colWidth') },
  ];
  void primaryRange;
  return <Menu items={items} x={menu.x} y={menu.y} onClose={close} minWidth={230} />;
}

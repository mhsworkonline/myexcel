'use client';
import { ComponentType } from 'react';
import { useStore } from '../../state/store';
import {
  AlertDialog,
  ConfirmDialog,
  GoToDialog,
  GoToSpecialDialog,
  InsertDeleteCellsDialog,
  MoveCopySheetDialog,
  RecoveryDialog,
  SaveAsDialog,
  SeriesDialog,
  SizeDialog,
  UnhideSheetDialog,
  ValidationErrorDialog,
  ZoomDialog,
} from './basic';
import {
  CFManagerDialog,
  CFQuickDialog,
  CFRuleDialog,
  CreateNamesDialog,
  CreateTableDialog,
  DataValidationDialog,
  EvaluateDialog,
  NameManagerDialog,
  NewNameDialog,
  PasteNameDialog,
  ProtectSheetDialog,
  RemoveDuplicatesDialog,
  StatsDialog,
  SymbolDialog,
  TextToColumnsDialog,
  UnprotectSheetDialog,
} from './data';
import { CustomFilterDialog, FindReplaceDialog, PasteSpecialDialog, SortDialog, Top10Dialog } from './FindSortPaste';
import { FormatCellsDialog } from './FormatCells';
import { ShortcutsDialog } from './Shortcuts';
import { FunctionArgsDialog, HyperlinkDialog, InsertFunctionDialog, NoteDialog } from './misc';

type P = { props: Record<string, unknown> };

const REGISTRY: Record<string, ComponentType<P>> = {
  alert: AlertDialog,
  confirm: ConfirmDialog,
  validationError: ValidationErrorDialog,
  insertCells: () => <InsertDeleteCellsDialog mode="insert" />,
  deleteCells: () => <InsertDeleteCellsDialog mode="delete" />,
  rowHeight: ({ props }) => <SizeDialog axis="row" props={props} />,
  colWidth: ({ props }) => <SizeDialog axis="col" props={props} />,
  zoom: ZoomDialog,
  series: SeriesDialog,
  moveCopySheet: MoveCopySheetDialog,
  unhideSheet: UnhideSheetDialog,
  recovery: RecoveryDialog,
  saveAs: SaveAsDialog,
  goto: GoToDialog,
  gotoSpecial: GoToSpecialDialog,
  formatCells: FormatCellsDialog,
  findReplace: FindReplaceDialog,
  sort: SortDialog,
  customFilter: CustomFilterDialog,
  top10Filter: Top10Dialog,
  pasteSpecial: PasteSpecialDialog,
  insertFunction: InsertFunctionDialog,
  functionArgs: FunctionArgsDialog,
  note: NoteDialog,
  hyperlink: HyperlinkDialog,
  cfQuick: CFQuickDialog,
  cfRule: CFRuleDialog,
  cfManager: CFManagerDialog,
  dataValidation: DataValidationDialog,
  nameManager: NameManagerDialog,
  newName: NewNameDialog,
  createNames: CreateNamesDialog,
  pasteName: PasteNameDialog,
  createTable: CreateTableDialog,
  removeDuplicates: RemoveDuplicatesDialog,
  textToColumns: TextToColumnsDialog,
  protectSheet: ProtectSheetDialog,
  unprotectSheet: UnprotectSheetDialog,
  symbol: SymbolDialog,
  stats: StatsDialog,
  evaluate: EvaluateDialog,
  shortcuts: ShortcutsDialog,
};

/** Later phases register additional dialogs here. */
export function registerDialogs(map: Record<string, ComponentType<P>>): void {
  Object.assign(REGISTRY, map);
}

export function DialogHost() {
  const dialog = useStore((s) => s.dialog);
  if (!dialog) return null;
  const C = REGISTRY[dialog.type];
  if (!C) return null;
  return <C key={dialog.type + JSON.stringify(Object.keys(dialog.props ?? {}))} props={dialog.props ?? {}} />;
}

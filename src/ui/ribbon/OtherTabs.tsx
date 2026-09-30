'use client';
import {
  ArrowDownAZ,
  ArrowUpAZ,
  BookOpen,
  Calculator,
  ChartArea,
  ChartBar,
  ChartLine,
  ChartPie,
  ChartScatter,
  Columns3,
  Eye,
  FileSpreadsheet,
  FileText,
  Filter,
  FilterX,
  Grid3x3,
  Link,
  ListChecks,
  Lock,
  MessageSquarePlus,
  MessageSquareX,
  Moon,
  PanelTop,
  Printer,
  RefreshCw,
  Rows3,
  ScanSearch,
  Snowflake,
  SplitSquareHorizontal,
  Sun,
  Table2,
  Target,
  TextCursorInput,
  Trash2,
  ZoomIn,
  ChevronLeft,
  ChevronRight,
  StickyNote,
  Braces,
  Tag,
  Layers,
  Ruler,
  Maximize,
  Sigma,
  Sparkles,
  Workflow,
  Shield,
  Search,
  Replace,
} from 'lucide-react';
import { fnCatalog } from '../../model/functions';
import { flashFill } from '../../state/actions/fill';
import { autoSum } from '../../state/actions/formulas';
import { clearAllFilters, quickSort, reapplyFilter, toggleAutoFilter } from '../../state/actions/sortFilter';
import { freezePanes, toggleSplit } from '../../state/actions/structure';
import { deleteNotesInSelection, nextNote, tableAtActive, tableSourceGuess, unprotectSheet, updateTable, convertTableToRange } from '../../state/actions/data';
import { bump, openDialog, S, setState, transact, useStore } from '../../state/store';
import { goToSpecial } from '../../state/actions/find';
import { toggleInvalidCircles } from '../grid/Grid';
import { MenuItem } from '../Menu';
import { ChartColumnIcon, PivotIcon, RecommendedChartsIcon, TableIcon, AutoSumIcon } from '../icons';
import { Col, Group, LargeButton, Row, SmallButton } from './parts';
import { formatAsTable, TableStyleGallery } from './HomeTab';
import { primaryRange } from '../../model/selection';
import { insertChart } from '../../state/charts';
import { getScroll } from '../grid/geometry';
import { TABLE_STYLE_NAMES } from '../../state/tableStyles';

function insertFnMenu(cat: string): MenuItem[] {
  const list = [...fnCatalog().values()].filter((f) => f.cat === cat);
  return [
    ...list.map((f) => ({ label: f.name, onClick: () => openDialog('functionArgs', { name: f.name }) })),
    { separator: true },
    { label: 'Insert Function...', onClick: () => openDialog('insertFunction') },
  ];
}

const sheetOf = () => S().wb.activeSheet;

function setSheetMeta<K extends keyof import('../../model/sheet').SheetMeta>(key: K, value: import('../../model/sheet').Sheet[K], label: string): void {
  const sheet = sheetOf();
  transact(label, (tx) => tx.setMeta(sheet, key, value));
}

function setPageSetup(patch: Partial<import('../../model/types').PageSetup>, label = 'Page Setup'): void {
  const sheet = sheetOf();
  setSheetMeta('pageSetup', { ...sheet.pageSetup, ...patch }, label);
}

export function chartMenu(kinds: [string, string][]): MenuItem[] {
  return [
    ...kinds.map(([label, type]) => ({ label, onClick: () => insertChart(type as import('../../model/types').ChartType) })),
    { separator: true },
    { label: 'More Charts...', onClick: () => openDialog('insertChart') },
  ];
}

export function InsertTab() {
  return (
    <>
      <Group label="Tables">
        <LargeButton icon={<PivotIcon />} label={'PivotTable'} onClick={() => openDialog('createPivot')} testId="btn-pivot" />
        <LargeButton icon={<TableIcon />} label="Table" onClick={() => { const g = tableSourceGuess(); openDialog('createTable', { range: g?.range, header: g?.header ?? true }); }} />
      </Group>
      <Group label="Charts" launcher={() => openDialog('insertChart')} title="See All Charts">
        <LargeButton icon={<RecommendedChartsIcon />} label={'Recommended\nCharts'} onClick={() => openDialog('insertChart')} />
        <Col>
          <Row>
            <SmallButton icon={<ChartColumnIcon />} title="Insert Column or Bar Chart" menu={chartMenu([['Clustered Column', 'column'], ['Stacked Column', 'stackedColumn'], ['Clustered Bar', 'bar'], ['Stacked Bar', 'stackedBar']])} testId="btn-chart-column" />
            <SmallButton icon={<ChartLine size={16} className="text-[#2B7CD3]" />} title="Insert Line or Area Chart" menu={chartMenu([['Line', 'line'], ['Line with Markers', 'lineMarkers'], ['Area', 'area'], ['Stacked Area', 'stackedArea']])} />
          </Row>
          <Row>
            <SmallButton icon={<ChartPie size={16} className="text-[#ED7D31]" />} title="Insert Pie or Doughnut Chart" menu={chartMenu([['Pie', 'pie'], ['Doughnut', 'doughnut']])} />
            <SmallButton icon={<ChartScatter size={16} className="text-[#2B7CD3]" />} title="Insert Scatter (X, Y) Chart" menu={chartMenu([['Scatter', 'scatter']])} />
          </Row>
          <Row>
            <SmallButton icon={<ChartBar size={16} className="text-[#70AD47]" />} title="Insert Combo Chart" onClick={() => insertChart('combo')} />
            <SmallButton icon={<ChartArea size={16} className="text-[#5B9BD5]" />} title="Insert Area Chart" onClick={() => insertChart('area')} />
          </Row>
        </Col>
      </Group>
      <Group label="Sparklines">
        <LargeButton icon={<ChartLine size={26} className="text-[#2B7CD3]" />} label="Line" onClick={() => openDialog('sparklines', { type: 'line' })} />
        <LargeButton icon={<ChartBar size={26} className="text-[#2B7CD3] rotate-90" />} label="Column" onClick={() => openDialog('sparklines', { type: 'column' })} />
        <LargeButton icon={<Columns3 size={26} className="text-[#2B7CD3]" />} label={'Win/\nLoss'} onClick={() => openDialog('sparklines', { type: 'winloss' })} />
      </Group>
      <Group label="Links">
        <LargeButton icon={<Link size={26} className="text-[#2B7CD3]" />} label="Link" onClick={() => openDialog('hyperlink')} />
      </Group>
      <Group label="Comments">
        <LargeButton icon={<MessageSquarePlus size={26} className="text-[#7B61FF]" />} label="Comment" onClick={() => openDialog('note', { edit: true, threaded: true })} />
      </Group>
      <Group label="Symbols">
        <LargeButton icon={<span className="text-[26px] leading-none">Ω</span>} label="Symbol" onClick={() => openDialog('symbol')} />
      </Group>
    </>
  );
}

export function PageLayoutTab() {
  useStore((s) => s.rev);
  const sheet = sheetOf();
  const ps = sheet.pageSetup;
  return (
    <>
      <Group label="Page Setup" launcher={() => openDialog('pageSetup')}>
        <LargeButton
          icon={<Maximize size={26} className="text-[#2B7CD3]" />}
          label="Margins"
          menu={[
            { label: 'Normal  (T 0.75" B 0.75" L 0.7" R 0.7")', onClick: () => setPageSetup({ margins: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7, header: 0.3, footer: 0.3 } }) },
            { label: 'Wide  (T 1" B 1" L 1" R 1")', onClick: () => setPageSetup({ margins: { top: 1, bottom: 1, left: 1, right: 1, header: 0.5, footer: 0.5 } }) },
            { label: 'Narrow  (T 0.75" B 0.75" L 0.25" R 0.25")', onClick: () => setPageSetup({ margins: { top: 0.75, bottom: 0.75, left: 0.25, right: 0.25, header: 0.3, footer: 0.3 } }) },
            { separator: true },
            { label: 'Custom Margins...', onClick: () => openDialog('pageSetup', { tab: 'Margins' }) },
          ]}
        />
        <LargeButton
          icon={<FileText size={26} className="text-[#2B7CD3]" />}
          label="Orientation"
          menu={[
            { label: 'Portrait', checked: ps.orientation === 'portrait', onClick: () => setPageSetup({ orientation: 'portrait' }) },
            { label: 'Landscape', checked: ps.orientation === 'landscape', onClick: () => setPageSetup({ orientation: 'landscape' }) },
          ]}
        />
        <LargeButton
          icon={<FileSpreadsheet size={26} className="text-[#2B7CD3]" />}
          label="Size"
          menu={(['letter', 'legal', 'a4', 'a3', 'tabloid'] as const).map((p) => ({ label: p === 'a4' || p === 'a3' ? p.toUpperCase() : p[0].toUpperCase() + p.slice(1), checked: ps.paperSize === p, onClick: () => setPageSetup({ paperSize: p }) }))}
        />
        <LargeButton
          icon={<ScanSearch size={26} className="text-[#2B7CD3]" />}
          label={'Print\nArea'}
          menu={[
            { label: 'Set Print Area', onClick: () => setPageSetup({ printArea: primaryRange(S().sel) }, 'Set Print Area') },
            { label: 'Clear Print Area', onClick: () => setPageSetup({ printArea: undefined }, 'Clear Print Area') },
          ]}
        />
        <LargeButton
          icon={<Rows3 size={26} className="text-[#2B7CD3]" />}
          label="Breaks"
          menu={[
            { label: 'Insert Page Break', onClick: () => import('../../state/print').then((m) => m.insertPageBreak()) },
            { label: 'Remove Page Break', onClick: () => import('../../state/print').then((m) => m.removePageBreak()) },
            { label: 'Reset All Page Breaks', onClick: () => import('../../state/print').then((m) => m.resetPageBreaks()) },
          ]}
        />
        <LargeButton icon={<PanelTop size={26} className="text-[#2B7CD3]" />} label={'Print\nTitles'} onClick={() => openDialog('pageSetup', { tab: 'Sheet' })} />
      </Group>
      <Group label="Scale to Fit" launcher={() => openDialog('pageSetup', { tab: 'Page' })}>
        <Col>
          <Row>
            <span className="w-12">Width:</span>
            <select className="xl-input !h-[22px] w-24" value={ps.fitToPage ? String(ps.fitToWidth) : 'auto'} onChange={(e) => setPageSetup(e.target.value === 'auto' ? { fitToPage: false } : { fitToPage: true, fitToWidth: +e.target.value })}>
              <option value="auto">Automatic</option>
              {[1, 2, 3, 4, 5].map((n) => (
                <option key={n} value={n}>{n} page{n > 1 ? 's' : ''}</option>
              ))}
            </select>
          </Row>
          <Row>
            <span className="w-12">Height:</span>
            <select className="xl-input !h-[22px] w-24" value={ps.fitToPage ? String(ps.fitToHeight) : 'auto'} onChange={(e) => setPageSetup(e.target.value === 'auto' ? { fitToPage: false } : { fitToPage: true, fitToHeight: +e.target.value })}>
              <option value="auto">Automatic</option>
              {[1, 2, 3, 4, 5].map((n) => (
                <option key={n} value={n}>{n} page{n > 1 ? 's' : ''}</option>
              ))}
            </select>
          </Row>
          <Row>
            <span className="w-12">Scale:</span>
            <input className="xl-input !h-[22px] w-24" type="number" min={10} max={400} value={ps.scale} disabled={ps.fitToPage} onChange={(e) => setPageSetup({ scale: Math.max(10, Math.min(400, +e.target.value || 100)) })} />
          </Row>
        </Col>
      </Group>
      <Group label="Sheet Options" launcher={() => openDialog('pageSetup', { tab: 'Sheet' })}>
        <div className="grid grid-cols-2 gap-x-4 gap-y-[2px] pt-1">
          <span className="font-semibold">Gridlines</span>
          <span className="font-semibold">Headings</span>
          <label className="flex items-center gap-1"><input type="checkbox" checked={sheet.showGridlines} onChange={(e) => setSheetMeta('showGridlines', e.target.checked, 'Gridlines')} /> View</label>
          <label className="flex items-center gap-1"><input type="checkbox" checked={sheet.showHeaders} onChange={(e) => setSheetMeta('showHeaders', e.target.checked, 'Headings')} /> View</label>
          <label className="flex items-center gap-1"><input type="checkbox" checked={ps.gridlines} onChange={(e) => setPageSetup({ gridlines: e.target.checked })} /> Print</label>
          <label className="flex items-center gap-1"><input type="checkbox" checked={ps.headings} onChange={(e) => setPageSetup({ headings: e.target.checked })} /> Print</label>
        </div>
      </Group>
    </>
  );
}

export function FormulasTab() {
  useStore((s) => s.rev);
  const showFormulas = useStore((s) => s.showFormulas);
  return (
    <>
      <Group label="Function Library">
        <LargeButton icon={<span className="text-[22px] italic font-serif">fx</span>} label={'Insert\nFunction'} onClick={() => openDialog('insertFunction')} />
        <LargeButton icon={<AutoSumIcon size={26} />} label="AutoSum" onClick={() => autoSum('SUM')} menu={[{ label: 'Sum', onClick: () => autoSum('SUM') }, { label: 'Average', onClick: () => autoSum('AVERAGE') }, { label: 'Count Numbers', onClick: () => autoSum('COUNT') }, { label: 'Max', onClick: () => autoSum('MAX') }, { label: 'Min', onClick: () => autoSum('MIN') }]} />
        <LargeButton icon={<Calculator size={24} className="text-[#107C41]" />} label="Financial" menu={() => insertFnMenu('Financial')} />
        <LargeButton icon={<span className="text-[20px] font-bold text-[#7030A0]">?</span>} label="Logical" menu={() => insertFnMenu('Logical')} />
        <LargeButton icon={<span className="text-[20px] font-bold text-[#2B7CD3]">A</span>} label="Text" menu={() => insertFnMenu('Text')} />
        <LargeButton icon={<span className="text-[18px] text-[#C55A11]">📅</span>} label={'Date &\nTime'} menu={() => insertFnMenu('Date & Time')} />
        <LargeButton icon={<ScanSearch size={24} className="text-[#2B7CD3]" />} label={'Lookup &\nReference'} menu={() => insertFnMenu('Lookup & Reference')} />
        <LargeButton icon={<span className="text-[20px] font-semibold text-[#107C41]">θ</span>} label={'Math &\nTrig'} menu={() => insertFnMenu('Math & Trig')} />
        <LargeButton
          icon={<Sigma size={24} className="text-[#2B7CD3]" />}
          label={'More\nFunctions'}
          menu={[
            { label: 'Statistical', submenu: insertFnMenu('Statistical') },
            { label: 'Engineering', submenu: insertFnMenu('Engineering') },
            { label: 'Information', submenu: insertFnMenu('Information') },
            { label: 'Database', submenu: insertFnMenu('Database') },
          ]}
        />
      </Group>
      <Group label="Defined Names">
        <LargeButton icon={<Tag size={26} className="text-[#2B7CD3]" />} label={'Name\nManager'} onClick={() => openDialog('nameManager')} />
        <Col>
          <SmallButton icon={<Tag size={14} />} label="Define Name" onClick={() => openDialog('newName')} />
          <SmallButton
            icon={<Braces size={14} />}
            label="Use in Formula"
            menu={() => [
              ...S().wb.names.map((n) => ({
                label: n.name,
                onClick: () => {
                  const ed = S().edit;
                  if (ed) setState({ edit: { ...ed, text: ed.text.slice(0, ed.caret) + n.name + ed.text.slice(ed.caret), caret: ed.caret + n.name.length } });
                  else import('../../state/actions/edit').then((m) => m.beginEdit('enter', '=' + n.name));
                },
              })),
              { separator: true },
              { label: 'Paste Names...', onClick: () => openDialog('pasteName') },
            ]}
          />
          <SmallButton icon={<Table2 size={14} />} label="Create from Selection" onClick={() => openDialog('createNames')} />
        </Col>
      </Group>
      <Group label="Formula Auditing">
        <Col>
          <SmallButton icon={<Workflow size={14} />} label="Trace Precedents" onClick={() => goToSpecial({ kind: 'precedents' })} />
          <SmallButton icon={<Workflow size={14} className="scale-x-[-1]" />} label="Trace Dependents" onClick={() => goToSpecial({ kind: 'dependents' })} />
          <SmallButton icon={<TextCursorInput size={14} />} label="Show Formulas" checked={showFormulas} onClick={() => { setState({ showFormulas: !showFormulas }); bump(); }} />
        </Col>
        <Col>
          <SmallButton icon={<ListChecks size={14} />} label="Error Checking" onClick={() => goToSpecial({ kind: 'formulas', numbers: false, text: false, logicals: false, errors: true })} />
          <SmallButton icon={<Target size={14} />} label="Evaluate Formula" onClick={() => openDialog('evaluate')} />
        </Col>
      </Group>
      <Group label="Calculation">
        <LargeButton icon={<Calculator size={26} className="text-[#2B7CD3]" />} label={'Calculation\nOptions'} menu={[{ label: 'Automatic', checked: true }, { label: 'Manual', disabled: true }]} />
        <Col>
          <SmallButton icon={<RefreshCw size={14} />} label="Calculate Now" onClick={() => { S().engine.recalc(); bump(); }} />
          <SmallButton icon={<RefreshCw size={14} />} label="Calculate Sheet" onClick={() => { S().engine.recalc(); bump(); }} />
        </Col>
      </Group>
    </>
  );
}

export function DataTab() {
  useStore((s) => s.rev);
  const sheet = sheetOf();
  return (
    <>
      <Group label="Get & Transform Data">
        <LargeButton icon={<FileText size={26} className="text-[#107C41]" />} label={'From\nText/CSV'} onClick={() => import('../../state/actions/file').then((m) => m.openFile())} />
      </Group>
      <Group label="Queries & Connections">
        <LargeButton icon={<RefreshCw size={26} className="text-[#107C41]" />} label={'Refresh\nAll'} onClick={() => import('../../state/pivot').then((m) => m.refreshAllPivots())} />
      </Group>
      <Group label="Sort & Filter">
        <Col>
          <SmallButton icon={<ArrowDownAZ size={16} />} title="Sort A to Z" onClick={() => quickSort(false)} testId="data-sort-asc" />
          <SmallButton icon={<ArrowUpAZ size={16} />} title="Sort Z to A" onClick={() => quickSort(true)} />
        </Col>
        <LargeButton icon={<Layers size={26} className="text-[#2B7CD3]" />} label="Sort" onClick={() => openDialog('sort')} />
        <LargeButton icon={<Filter size={26} className="text-[#2B7CD3]" />} label="Filter" checked={!!sheet.autoFilter} onClick={toggleAutoFilter} testId="btn-filter" />
        <Col>
          <SmallButton icon={<FilterX size={14} />} label="Clear" disabled={!sheet.autoFilter} onClick={clearAllFilters} />
          <SmallButton icon={<RefreshCw size={14} />} label="Reapply" disabled={!sheet.autoFilter} onClick={reapplyFilter} />
        </Col>
      </Group>
      <Group label="Data Tools">
        <LargeButton icon={<Columns3 size={26} className="text-[#2B7CD3]" />} label={'Text to\nColumns'} onClick={() => openDialog('textToColumns')} />
        <Col>
          <SmallButton icon={<Sparkles size={14} className="text-[#FFC000]" />} label="Flash Fill" onClick={flashFill} />
          <SmallButton icon={<Trash2 size={14} />} label="Remove Duplicates" onClick={() => openDialog('removeDuplicates')} />
          <SmallButton
            icon={<ListChecks size={14} className="text-[#107C41]" />}
            label="Data Validation"
            onClick={() => openDialog('dataValidation')}
            menu={[
              { label: 'Data Validation...', onClick: () => openDialog('dataValidation') },
              { label: 'Circle Invalid Data', onClick: () => toggleInvalidCircles(true) },
              { label: 'Clear Validation Circles', onClick: () => toggleInvalidCircles(false) },
            ]}
          />
        </Col>
      </Group>
      <Group label="Forecast">
        <LargeButton
          icon={<ScanSearch size={26} className="text-[#2B7CD3]" />}
          label={'What-If\nAnalysis'}
          menu={[
            { label: 'Scenario Manager...', onClick: () => openDialog('scenarioManager') },
            { label: 'Goal Seek...', onClick: () => openDialog('goalSeek') },
            { label: 'Data Table...', onClick: () => openDialog('dataTable') },
          ]}
        />
      </Group>
      <Group label="Analyze">
        <LargeButton icon={<Target size={26} className="text-[#C55A11]" />} label="Solver" onClick={() => openDialog('solver')} />
      </Group>
    </>
  );
}

export function ReviewTab() {
  useStore((s) => s.rev);
  const sheet = sheetOf();
  return (
    <>
      <Group label="Proofing">
        <LargeButton icon={<BookOpen size={26} className="text-[#2B7CD3]" />} label={'Workbook\nStatistics'} onClick={() => openDialog('stats')} />
      </Group>
      <Group label="Comments">
        <LargeButton icon={<MessageSquarePlus size={26} className="text-[#7B61FF]" />} label={'New\nComment'} onClick={() => openDialog('note', { edit: true, threaded: true })} />
        <LargeButton icon={<MessageSquareX size={26} className="text-[#C00000]" />} label="Delete" onClick={deleteNotesInSelection} />
        <Col>
          <SmallButton icon={<ChevronLeft size={14} />} label="Previous" onClick={() => nextNote(-1)} />
          <SmallButton icon={<ChevronRight size={14} />} label="Next" onClick={() => nextNote(1)} />
        </Col>
      </Group>
      <Group label="Notes">
        <LargeButton
          icon={<StickyNote size={26} className="text-[#FFC000]" />}
          label="Notes"
          menu={[
            { label: 'New Note', onClick: () => openDialog('note', { edit: true }) },
            { label: 'Edit Note', onClick: () => openDialog('note', { edit: true }) },
            { label: 'Previous Note', onClick: () => nextNote(-1) },
            { label: 'Next Note', onClick: () => nextNote(1) },
            { label: 'Delete Note', onClick: deleteNotesInSelection },
          ]}
        />
      </Group>
      <Group label="Protect">
        <LargeButton
          icon={sheet.protection ? <Shield size={26} className="text-[#C00000]" /> : <Lock size={26} className="text-[#2B7CD3]" />}
          label={sheet.protection ? 'Unprotect\nSheet' : 'Protect\nSheet'}
          onClick={() => (sheet.protection ? (sheet.protection.passwordHash ? openDialog('unprotectSheet') : unprotectSheet('')) : openDialog('protectSheet'))}
          testId="btn-protect"
        />
        <LargeButton icon={<Lock size={26} className="text-[#107C41]" />} label={'Allow Edit\nRanges'} onClick={() => openDialog('formatCells', { tab: 'Protection' })} />
      </Group>
    </>
  );
}

export function ViewTab() {
  useStore((s) => s.rev);
  const theme = useStore((s) => s.theme);
  const viewMode = useStore((s) => s.viewMode);
  const sheet = sheetOf();
  const fbHidden = useStore((s) => s.formulaBarExpanded);
  void fbHidden;
  const setZoom = (z: number) => {
    sheet.zoom = z;
    sheet.touch();
    bump();
  };
  return (
    <>
      <Group label="Workbook Views">
        <LargeButton icon={<Grid3x3 size={26} className="text-[#2B7CD3]" />} label="Normal" checked={viewMode === 'normal'} onClick={() => import('../../state/print').then((m) => m.setViewMode('normal'))} />
        <LargeButton icon={<Rows3 size={26} className="text-[#2B7CD3]" />} label={'Page Break\nPreview'} checked={viewMode === 'pageBreak'} onClick={() => import('../../state/print').then((m) => m.setViewMode('pageBreak'))} />
        <LargeButton icon={<FileText size={26} className="text-[#2B7CD3]" />} label={'Page\nLayout'} onClick={() => openDialog('print')} />
      </Group>
      <Group label="Show">
        <Col className="pt-1">
          <label className="flex items-center gap-1"><input type="checkbox" checked={sheet.showGridlines} onChange={(e) => setSheetMeta('showGridlines', e.target.checked, 'Gridlines')} /> Gridlines</label>
          <label className="flex items-center gap-1"><input type="checkbox" checked={sheet.showHeaders} onChange={(e) => setSheetMeta('showHeaders', e.target.checked, 'Headings')} /> Headings</label>
          <label className="flex items-center gap-1"><input type="checkbox" checked={!document.body.dataset.hideFormulaBar} onChange={(e) => { if (e.target.checked) delete document.body.dataset.hideFormulaBar; else document.body.dataset.hideFormulaBar = '1'; bump(); }} /> Formula Bar</label>
        </Col>
      </Group>
      <Group label="Zoom">
        <LargeButton icon={<ZoomIn size={26} className="text-[#2B7CD3]" />} label="Zoom" onClick={() => openDialog('zoom')} />
        <LargeButton icon={<span className="text-[15px] font-bold">100</span>} label="100%" onClick={() => setZoom(100)} />
        <LargeButton
          icon={<Maximize size={24} className="text-[#2B7CD3]" />}
          label={'Zoom to\nSelection'}
          onClick={() => {
            const rg = primaryRange(S().sel);
            const w = Math.min(rg.c2 - rg.c1 + 1, 200) * 64;
            const h = Math.min(rg.r2 - rg.r1 + 1, 2000) * 20;
            const el = document.querySelector('[data-grid-root]')?.getBoundingClientRect();
            const z = Math.max(10, Math.min(400, Math.floor(Math.min((el?.width ?? 1000) / w, (el?.height ?? 600) / h) * 100)));
            setZoom(z);
            const sc = getScroll(sheet.id);
            sc.x[1] = 0;
            sc.y[1] = 0;
          }}
        />
      </Group>
      <Group label="Window">
        <LargeButton
          icon={<Snowflake size={26} className="text-[#2B7CD3]" />}
          label={'Freeze\nPanes'}
          testId="btn-freeze"
          menu={[
            sheet.freeze.rows || sheet.freeze.cols ? { label: 'Unfreeze Panes', onClick: () => freezePanes('unfreeze') } : { label: 'Freeze Panes', onClick: () => freezePanes('panes') },
            { label: 'Freeze Top Row', onClick: () => freezePanes('topRow') },
            { label: 'Freeze First Column', onClick: () => freezePanes('firstCol') },
          ]}
        />
        <Col>
          <SmallButton icon={<SplitSquareHorizontal size={14} />} label="Split" checked={!!sheet.split} onClick={toggleSplit} />
          <SmallButton icon={<Eye size={14} />} label="Unhide" onClick={() => openDialog('unhideSheet')} />
        </Col>
      </Group>
      <Group label="Theme">
        <LargeButton icon={theme === 'dark' ? <Sun size={26} /> : <Moon size={26} />} label={theme === 'dark' ? 'Light\nMode' : 'Dark\nMode'} onClick={() => setTheme(theme === 'dark' ? 'light' : 'dark')} testId="btn-theme" />
      </Group>
    </>
  );
}

export function setTheme(t: 'light' | 'dark'): void {
  // applied immediately; persisted to settings (settings.json on desktop, localStorage on web)
  import('../../state/settings').then((m) => m.setThemeSetting(t));
}

export function TableDesignTab() {
  useStore((s) => s.rev);
  const t = tableAtActive();
  if (!t) return null;
  const set = (patch: Partial<import('../../model/types').TableDef>) => updateTable(t.id, patch);
  return (
    <>
      <Group label="Properties">
        <Col className="pt-1">
          <span>Table Name:</span>
          <input
            key={t.id + t.name}
            className="xl-input !h-[22px] w-28"
            defaultValue={t.name}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                const v = (e.target as HTMLInputElement).value.trim();
                if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(v)) set({ name: v });
              }
            }}
          />
        </Col>
      </Group>
      <Group label="Tools">
        <Col>
          <SmallButton icon={<Trash2 size={14} />} label="Remove Duplicates" onClick={() => openDialog('removeDuplicates', { range: t.range })} />
          <SmallButton icon={<Replace size={14} />} label="Convert to Range" onClick={() => convertTableToRange(t.id)} />
          <SmallButton icon={<Search size={14} />} label="Summarize with PivotTable" onClick={() => openDialog('createPivot', { range: t.range })} />
        </Col>
      </Group>
      <Group label="Table Style Options">
        <div className="grid grid-cols-3 gap-x-3 gap-y-[2px] pt-1">
          {(
            [
              ['headerRow', 'Header Row'],
              ['firstCol', 'First Column'],
              ['showFilterButton', 'Filter Button'],
              ['totalRow', 'Total Row'],
              ['lastCol', 'Last Column'],
              ['bandedRows', 'Banded Rows'],
              ['bandedCols', 'Banded Columns'],
            ] as const
          ).map(([k, label]) => (
            <label key={k} className="flex items-center gap-1 whitespace-nowrap">
              <input type="checkbox" checked={!!t[k]} disabled={k === 'headerRow'} onChange={(e) => set({ [k]: e.target.checked })} /> {label}
            </label>
          ))}
        </div>
      </Group>
      <Group label="Table Styles">
        <LargeButton icon={<TableIcon />} label={'Quick\nStyles'} panel={(close) => <TableStyleGallery close={close} onPick={(s) => set({ style: s })} />} />
        <div className="text-[11px] opacity-70 self-center px-2">{t.style.replace('TableStyle', '')}</div>
      </Group>
    </>
  );
}

export { formatAsTable, TABLE_STYLE_NAMES, Printer, Ruler };

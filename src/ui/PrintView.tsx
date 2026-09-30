'use client';
import { ArrowLeft, ChevronLeft, ChevronRight, FileDown, Printer } from 'lucide-react';
import { useEffect, useMemo, useState } from 'react';
import { colToName, parseRange, rangeToA1 } from '../model/address';
import type { PageSetup } from '../model/types';
import { collectPages, exportPdf, printPages, PrintWhat, renderPage } from '../state/print';
import { closeDialog, openDialog, S, transact, useStore } from '../state/store';
import { Dialog } from './dialogs/Dialog';

type Scaling = 'none' | 'fitSheet' | 'fitCols' | 'fitRows';

export function PrintView() {
  const sheet = S().wb.activeSheet;
  const [what, setWhat] = useState<PrintWhat>('active');
  const [ov, setOv] = useState<Partial<PageSetup>>({});
  const [page, setPage] = useState(0);
  const [img, setImg] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  useStore((s) => s.rev);
  const ps = { ...sheet.pageSetup, ...ov };
  const pages = useMemo(() => collectPages(what, ov), [what, ov]); // eslint-disable-line react-hooks/exhaustive-deps
  const scaling: Scaling = !ps.fitToPage ? 'none' : ps.fitToWidth === 1 && ps.fitToHeight === 1 ? 'fitSheet' : ps.fitToWidth === 1 && !ps.fitToHeight ? 'fitCols' : 'fitRows';
  useEffect(() => {
    let alive = true;
    const p = pages[Math.min(page, pages.length - 1)];
    if (!p) return;
    renderPage(p, pages.length, 1.25, ov).then((c) => alive && setImg(c.toDataURL('image/png')));
    return () => {
      alive = false;
    };
  }, [pages, page, ov]);
  const setScaling = (s: Scaling) => {
    if (s === 'none') setOv({ ...ov, fitToPage: false, scale: 100 });
    else if (s === 'fitSheet') setOv({ ...ov, fitToPage: true, fitToWidth: 1, fitToHeight: 1 });
    else if (s === 'fitCols') setOv({ ...ov, fitToPage: true, fitToWidth: 1, fitToHeight: 0 });
    else setOv({ ...ov, fitToPage: true, fitToWidth: 0, fitToHeight: 1 });
  };
  const persist = () => {
    if (Object.keys(ov).length) transact('Page Setup', (tx) => tx.setMeta(sheet, 'pageSetup', { ...sheet.pageSetup, ...ov }));
  };
  return (
    <div className="fixed inset-0 z-[66] flex" style={{ background: 'var(--ribbon-bg)' }} data-testid="print-view">
      <div className="w-[60px] flex flex-col items-center pt-3 text-white" style={{ background: '#107C41' }}>
        <button className="p-2 rounded hover:bg-white/10" onClick={() => { persist(); closeDialog(); }} aria-label="Back"><ArrowLeft size={18} /></button>
      </div>
      <div className="w-[300px] p-6 flex flex-col gap-3 border-r overflow-auto" style={{ borderColor: 'var(--border)' }}>
        <h1 className="text-[26px] font-light">Print</h1>
        <div className="flex gap-2">
          <button className="flex flex-col items-center gap-1 px-4 py-2 border rounded hover:bg-[var(--hover)]" style={{ borderColor: 'var(--border-strong)' }} disabled={busy} onClick={async () => { setBusy(true); await printPages(pages, ov); setBusy(false); }} data-testid="print-go">
            <Printer size={26} /> Print
          </button>
          <button className="flex flex-col items-center gap-1 px-4 py-2 border rounded hover:bg-[var(--hover)]" style={{ borderColor: 'var(--border-strong)' }} disabled={busy} onClick={async () => { setBusy(true); await exportPdf(pages, ov); setBusy(false); }} data-testid="print-pdf">
            <FileDown size={26} /> Export PDF
          </button>
        </div>
        <div className="font-semibold mt-2">Settings</div>
        <select className="xl-input" value={what} onChange={(e) => { setWhat(e.target.value as PrintWhat); setPage(0); }}>
          <option value="active">Print Active Sheets</option>
          <option value="workbook">Print Entire Workbook</option>
          <option value="selection">Print Selection</option>
        </select>
        <select className="xl-input" value={ps.orientation} onChange={(e) => setOv({ ...ov, orientation: e.target.value as PageSetup['orientation'] })}>
          <option value="portrait">Portrait Orientation</option>
          <option value="landscape">Landscape Orientation</option>
        </select>
        <select className="xl-input" value={ps.paperSize} onChange={(e) => setOv({ ...ov, paperSize: e.target.value as PageSetup['paperSize'] })}>
          <option value="letter">Letter 8.5" x 11"</option>
          <option value="legal">Legal 8.5" x 14"</option>
          <option value="a4">A4 8.27" x 11.69"</option>
          <option value="a3">A3 11.69" x 16.54"</option>
          <option value="tabloid">Tabloid 11" x 17"</option>
        </select>
        <select className="xl-input" value={ps.margins.left === 0.25 ? 'narrow' : ps.margins.left === 1 ? 'wide' : 'normal'} onChange={(e) => {
          const m = e.target.value === 'narrow' ? { top: 0.75, bottom: 0.75, left: 0.25, right: 0.25, header: 0.3, footer: 0.3 } : e.target.value === 'wide' ? { top: 1, bottom: 1, left: 1, right: 1, header: 0.5, footer: 0.5 } : { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7, header: 0.3, footer: 0.3 };
          setOv({ ...ov, margins: m });
        }}>
          <option value="normal">Normal Margins</option>
          <option value="wide">Wide Margins</option>
          <option value="narrow">Narrow Margins</option>
        </select>
        <select className="xl-input" value={scaling} onChange={(e) => setScaling(e.target.value as Scaling)} data-testid="print-scaling">
          <option value="none">No Scaling</option>
          <option value="fitSheet">Fit Sheet on One Page</option>
          <option value="fitCols">Fit All Columns on One Page</option>
          <option value="fitRows">Fit All Rows on One Page</option>
        </select>
        <button className="text-left underline text-[12px]" style={{ color: 'var(--accent)' }} onClick={() => { persist(); openDialog('pageSetup'); }}>Page Setup</button>
      </div>
      <div className="flex-1 flex flex-col items-center p-6 overflow-auto" style={{ background: 'var(--chrome-bg)' }}>
        <div className="flex-1 flex items-center justify-center min-h-0">
          {img && <img src={img} alt={`Page ${page + 1}`} className="max-h-full max-w-full shadow-lg bg-white" style={{ maxHeight: 'calc(100vh - 110px)' }} />}
        </div>
        <div className="flex items-center gap-2 mt-3" data-testid="print-nav">
          <button className="p-1 rounded hover:bg-[var(--hover)]" disabled={page === 0} onClick={() => setPage(page - 1)} aria-label="Previous page"><ChevronLeft size={16} /></button>
          <span>{Math.min(page + 1, pages.length)} of {pages.length}</span>
          <button className="p-1 rounded hover:bg-[var(--hover)]" disabled={page >= pages.length - 1} onClick={() => setPage(page + 1)} aria-label="Next page"><ChevronRight size={16} /></button>
        </div>
      </div>
    </div>
  );
}

type Tab = 'Page' | 'Margins' | 'Header/Footer' | 'Sheet';

export function PageSetupDialog({ props }: { props: Record<string, unknown> }) {
  const sheet = S().wb.activeSheet;
  const [tab, setTab] = useState<Tab>((props.tab as Tab) ?? 'Page');
  const [ps, setPs] = useState<PageSetup>(structuredClone(sheet.pageSetup));
  const [area, setArea] = useState(ps.printArea ? rangeToA1(ps.printArea, true) : '');
  const [rowsTop, setRowsTop] = useState(ps.printTitleRows ? `$${ps.printTitleRows[0] + 1}:$${ps.printTitleRows[1] + 1}` : '');
  const [colsLeft, setColsLeft] = useState(ps.printTitleCols ? `${colToName(ps.printTitleCols[0])}:${colToName(ps.printTitleCols[1])}` : '');
  const set = (p: Partial<PageSetup>) => setPs({ ...ps, ...p });
  const margin = (k: keyof PageSetup['margins'], label: string) => (
    <label className="flex flex-col gap-0.5">
      {label}
      <input type="number" step={0.05} min={0} className="xl-input w-20" value={ps.margins[k]} onChange={(e) => set({ margins: { ...ps.margins, [k]: Math.max(0, +e.target.value) } })} />
    </label>
  );
  const HF_PRESETS = ['(none)', 'Page &P', 'Page &P of &N', '&A', '&F', '&D', 'Confidential'];
  const section = (which: 'header' | 'footer') => (
    <div className="flex flex-col gap-1">
      <div className="flex items-center gap-2">
        <span className="font-semibold w-16 capitalize">{which}:</span>
        <select className="xl-input flex-1" onChange={(e) => set({ [which]: { left: '', center: e.target.value === '(none)' ? '' : e.target.value, right: '' } } as Partial<PageSetup>)} value="">
          <option value="" disabled>Presets…</option>
          {HF_PRESETS.map((p) => <option key={p} value={p}>{p.replace('&P', '1').replace('&N', '?').replace('&A', sheet.name).replace('&F', S().file.name).replace('&D', 'Date')}</option>)}
        </select>
      </div>
      <div className="grid grid-cols-3 gap-1">
        {(['left', 'center', 'right'] as const).map((k) => (
          <label key={k} className="flex flex-col gap-0.5 text-[11px] capitalize">{k} section
            <input className="xl-input" value={ps[which][k]} onChange={(e) => set({ [which]: { ...ps[which], [k]: e.target.value } } as Partial<PageSetup>)} />
          </label>
        ))}
      </div>
    </div>
  );
  return (
    <Dialog
      title="Page Setup"
      width={520}
      onOk={() => {
        const pa = area ? parseRange(area.replace(/\$/g, '').replace(/^.*!/, '')) : null;
        const rt = /(\d+):\$?(\d+)/.exec(rowsTop);
        const ct = /\$?([A-Z]+):\$?([A-Z]+)/i.exec(colsLeft);
        const col = (s: string) => s.toUpperCase().split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
        const next: PageSetup = { ...ps, printArea: pa ?? undefined, printTitleRows: rt ? [+rt[1] - 1, +rt[2] - 1] : undefined, printTitleCols: ct ? [col(ct[1]), col(ct[2])] : undefined };
        transact('Page Setup', (tx) => tx.setMeta(sheet, 'pageSetup', next));
      }}
      extraButtons={<button className="xl-btn" onClick={() => openDialog('print')}>Print Preview</button>}
      testId="page-setup"
    >
      <div className="flex border-b mb-3" style={{ borderColor: 'var(--border)' }}>
        {(['Page', 'Margins', 'Header/Footer', 'Sheet'] as Tab[]).map((t) => (
          <button key={t} className={'xl-dtab ' + (tab === t ? 'xl-dtab-active' : '')} onClick={() => setTab(t)}>{t}</button>
        ))}
      </div>
      <div className="min-h-[250px] flex flex-col gap-3">
        {tab === 'Page' && (
          <>
            <fieldset className="xl-fieldset">
              <legend>Orientation</legend>
              <label className="inline-flex gap-1 mr-6"><input type="radio" checked={ps.orientation === 'portrait'} onChange={() => set({ orientation: 'portrait' })} /> Portrait</label>
              <label className="inline-flex gap-1"><input type="radio" checked={ps.orientation === 'landscape'} onChange={() => set({ orientation: 'landscape' })} /> Landscape</label>
            </fieldset>
            <fieldset className="xl-fieldset flex flex-col gap-1">
              <legend>Scaling</legend>
              <label className="flex items-center gap-2"><input type="radio" checked={!ps.fitToPage} onChange={() => set({ fitToPage: false })} /> Adjust to: <input type="number" className="xl-input w-16" value={ps.scale} onChange={(e) => set({ scale: Math.max(10, Math.min(400, +e.target.value || 100)) })} /> % normal size</label>
              <label className="flex items-center gap-2"><input type="radio" checked={ps.fitToPage} onChange={() => set({ fitToPage: true })} /> Fit to: <input type="number" min={0} className="xl-input w-14" value={ps.fitToWidth} onChange={(e) => set({ fitToWidth: Math.max(0, +e.target.value) })} /> page(s) wide by <input type="number" min={0} className="xl-input w-14" value={ps.fitToHeight} onChange={(e) => set({ fitToHeight: Math.max(0, +e.target.value) })} /> tall</label>
            </fieldset>
            <label className="flex items-center gap-2">Paper size:
              <select className="xl-input flex-1" value={ps.paperSize} onChange={(e) => set({ paperSize: e.target.value as PageSetup['paperSize'] })}>
                <option value="letter">Letter</option><option value="legal">Legal</option><option value="a4">A4</option><option value="a3">A3</option><option value="tabloid">Tabloid</option>
              </select>
            </label>
          </>
        )}
        {tab === 'Margins' && (
          <>
            <div className="grid grid-cols-3 gap-3 items-end">
              <span />{margin('top', 'Top:')}{margin('header', 'Header:')}
              {margin('left', 'Left:')}<div className="h-24 border bg-white" style={{ borderColor: '#999' }} />{margin('right', 'Right:')}
              <span />{margin('bottom', 'Bottom:')}{margin('footer', 'Footer:')}
            </div>
            <fieldset className="xl-fieldset">
              <legend>Center on page</legend>
              <label className="inline-flex gap-1 mr-6"><input type="checkbox" checked={ps.centerH} onChange={(e) => set({ centerH: e.target.checked })} /> Horizontally</label>
              <label className="inline-flex gap-1"><input type="checkbox" checked={ps.centerV} onChange={(e) => set({ centerV: e.target.checked })} /> Vertically</label>
            </fieldset>
          </>
        )}
        {tab === 'Header/Footer' && (
          <>
            {section('header')}
            {section('footer')}
            <div className="text-[11px] opacity-70">Codes: &amp;P page, &amp;N pages, &amp;D date, &amp;T time, &amp;F file, &amp;A sheet.</div>
          </>
        )}
        {tab === 'Sheet' && (
          <>
            <label className="flex items-center gap-2"><span className="w-32">Print area:</span><input className="xl-input flex-1" value={area} onChange={(e) => setArea(e.target.value)} placeholder="$A$1:$F$40" /></label>
            <div className="font-semibold">Print titles</div>
            <label className="flex items-center gap-2"><span className="w-32">Rows to repeat at top:</span><input className="xl-input flex-1" value={rowsTop} onChange={(e) => setRowsTop(e.target.value)} placeholder="$1:$1" /></label>
            <label className="flex items-center gap-2"><span className="w-32">Columns to repeat at left:</span><input className="xl-input flex-1" value={colsLeft} onChange={(e) => setColsLeft(e.target.value)} placeholder="$A:$A" /></label>
            <fieldset className="xl-fieldset">
              <legend>Print</legend>
              <label className="inline-flex gap-1 mr-6"><input type="checkbox" checked={ps.gridlines} onChange={(e) => set({ gridlines: e.target.checked })} /> Gridlines</label>
              <label className="inline-flex gap-1"><input type="checkbox" checked={ps.headings} onChange={(e) => set({ headings: e.target.checked })} /> Row and column headings</label>
            </fieldset>
            <fieldset className="xl-fieldset">
              <legend>Page order</legend>
              <label className="inline-flex gap-1 mr-6"><input type="radio" checked={ps.pageOrder === 'downThenOver'} onChange={() => set({ pageOrder: 'downThenOver' })} /> Down, then over</label>
              <label className="inline-flex gap-1"><input type="radio" checked={ps.pageOrder === 'overThenDown'} onChange={() => set({ pageOrder: 'overThenDown' })} /> Over, then down</label>
            </fieldset>
          </>
        )}
      </div>
    </Dialog>
  );
}

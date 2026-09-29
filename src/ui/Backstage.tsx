'use client';
import { ArrowLeft, FilePlus, FileText, FolderOpen, Info, Printer, Save, SaveAll, Download, X } from 'lucide-react';
import { useEffect, useState } from 'react';
import { fileAdapter, RecentFile } from '../io/FileAdapter';
import { newWorkbook, openFile, saveAs, saveFile } from '../state/actions/file';
import { openDialog, S, setState, useStore } from '../state/store';

type Page = 'home' | 'new' | 'open' | 'info' | 'saveas' | 'export' | 'about';

export function Backstage() {
  const open = useStore((s) => s.backstage);
  const file = useStore((s) => s.file);
  const [page, setPage] = useState<Page>('home');
  const [recent, setRecent] = useState<RecentFile[]>([]);
  useEffect(() => {
    if (open) fileAdapter().recentFiles().then(setRecent).catch(() => setRecent([]));
  }, [open]);
  if (!open) return null;
  const close = () => setState({ backstage: false });
  const nav: [Page | 'save' | 'print' | 'close', string, React.ReactNode][] = [
    ['home', 'Home', <FileText key="h" size={16} />],
    ['new', 'New', <FilePlus key="n" size={16} />],
    ['open', 'Open', <FolderOpen key="o" size={16} />],
    ['info', 'Info', <Info key="i" size={16} />],
    ['save', 'Save', <Save key="s" size={16} />],
    ['saveas', 'Save As', <SaveAll key="sa" size={16} />],
    ['print', 'Print', <Printer key="p" size={16} />],
    ['export', 'Export', <Download key="e" size={16} />],
    ['close', 'Close', <X key="c" size={16} />],
  ];
  const templates = [
    { name: 'Blank workbook', action: () => { newWorkbook(); close(); } },
  ];
  return (
    <div className="fixed inset-0 z-[65] flex" style={{ background: 'var(--ribbon-bg)' }} data-testid="backstage">
      <div className="w-[200px] flex flex-col py-3 text-white" style={{ background: '#107C41' }}>
        <button className="flex items-center gap-2 px-5 py-3 hover:bg-white/10" onClick={close} aria-label="Back">
          <ArrowLeft size={18} />
        </button>
        {nav.map(([k, label, icon]) => (
          <button
            key={k}
            className={'flex items-center gap-3 px-5 py-2.5 text-left text-[13.5px] hover:bg-white/10 ' + (page === k ? 'bg-white/20' : '')}
            onClick={() => {
              if (k === 'save') { saveFile(); close(); return; }
              if (k === 'print') { close(); openDialog('print'); return; }
              if (k === 'close') { newWorkbook(); close(); return; }
              setPage(k as Page);
            }}
            data-testid={`bs-${k}`}
          >
            {icon}
            {label}
          </button>
        ))}
        <div className="flex-1" />
        <button className={'px-5 py-2.5 text-left text-[13.5px] hover:bg-white/10 ' + (page === 'about' ? 'bg-white/20' : '')} onClick={() => setPage('about')}>About</button>
      </div>
      <div className="flex-1 overflow-auto p-10">
        {(page === 'home' || page === 'new') && (
          <div>
            <h1 className="text-[26px] mb-6 font-light">{page === 'home' ? 'Good day' : 'New'}</h1>
            <div className="flex gap-4 mb-10">
              {templates.map((t) => (
                <button key={t.name} className="flex flex-col items-center gap-2 group" onClick={t.action} data-testid="bs-blank">
                  <div className="w-[160px] h-[120px] border bg-white group-hover:outline group-hover:outline-2 group-hover:outline-[var(--accent)] grid grid-cols-4 grid-rows-6" style={{ borderColor: 'var(--border-strong)' }}>
                    {Array.from({ length: 24 }, (_, i) => <span key={i} className="border-r border-b" style={{ borderColor: '#e1e1e1' }} />)}
                  </div>
                  <span>{t.name}</span>
                </button>
              ))}
            </div>
            <h2 className="text-[15px] font-semibold mb-2">Recent</h2>
            <RecentList recent={recent} />
          </div>
        )}
        {page === 'open' && (
          <div>
            <h1 className="text-[26px] mb-6 font-light">Open</h1>
            <button className="flex items-center gap-3 px-4 py-3 rounded hover:bg-[var(--hover)] mb-4" onClick={async () => { await openFile(); }} data-testid="bs-browse">
              <FolderOpen size={28} className="text-[#E8A200]" /> <span className="text-[14px]">Browse</span>
            </button>
            <h2 className="text-[15px] font-semibold mb-2">Recent</h2>
            <RecentList recent={recent} />
          </div>
        )}
        {page === 'info' && (
          <div className="max-w-[640px]">
            <h1 className="text-[26px] mb-6 font-light">Info</h1>
            <div className="text-[18px] mb-4">{file.name}</div>
            <div className="grid grid-cols-[180px_1fr] gap-y-2">
              <span className="opacity-70">Format</span><span>{file.format.toUpperCase()}</span>
              <span className="opacity-70">Sheets</span><span>{S().wb.sheets.length}</span>
              <span className="opacity-70">Last saved</span><span>{file.lastSaved ? new Date(file.lastSaved).toLocaleString() : 'Never'}</span>
              <span className="opacity-70">Autosave (recovery)</span><span>{file.autosaved ? new Date(file.autosaved).toLocaleTimeString() : 'Waiting for changes'}</span>
              <span className="opacity-70">Protection</span><span>{S().wb.sheets.filter((s) => s.protection).map((s) => s.name).join(', ') || 'None'}</span>
            </div>
          </div>
        )}
        {(page === 'saveas' || page === 'export') && (
          <div className="max-w-[560px]">
            <h1 className="text-[26px] mb-6 font-light">{page === 'saveas' ? 'Save As' : 'Export'}</h1>
            {([
              ['xlsx', 'Excel Workbook (*.xlsx)', 'Save as a standard Excel workbook'],
              ['csv', 'CSV UTF-8 (*.csv)', 'Active sheet only, comma separated'],
              ['tsv', 'Text (Tab delimited) (*.tsv)', 'Active sheet only'],
              ['ods', 'OpenDocument Spreadsheet (*.ods)', 'For LibreOffice / OpenOffice'],
              ['pdf', 'PDF (*.pdf)', 'Fixed layout, uses Page Setup'],
            ] as const).map(([f, l, d]) => (
              <button key={f} className="flex flex-col items-start w-full px-4 py-3 rounded hover:bg-[var(--hover)] border mb-2" style={{ borderColor: 'var(--border)' }} onClick={() => saveAs(f)} data-testid={`bs-save-${f}`}>
                <span className="text-[14px]">{l}</span>
                <span className="opacity-70 text-[12px]">{d}</span>
              </button>
            ))}
          </div>
        )}
        {page === 'about' && (
          <div className="max-w-[560px] leading-6">
            <h1 className="text-[26px] mb-6 font-light">About MyExcel</h1>
            <p>A local-first, single-user spreadsheet. Your files never leave this device.</p>
            <p className="mt-2 opacity-80">Calculation by HyperFormula (GPLv3), file I/O by ExcelJS and Papaparse, icons by Lucide.</p>
            <p className="mt-2 opacity-80">Not affiliated with Microsoft. Excel is a trademark of Microsoft Corporation.</p>
          </div>
        )}
      </div>
    </div>
  );
}

function RecentList({ recent }: { recent: RecentFile[] }) {
  if (!recent.length) return <div className="opacity-60">No recent files.</div>;
  return (
    <div className="flex flex-col max-w-[640px]">
      {recent.map((r) => (
        <button
          key={r.name + r.opened}
          className="flex items-center gap-3 px-3 py-2 rounded hover:bg-[var(--hover)] text-left"
          onClick={async () => {
            const fa = fileAdapter();
            if (fa.openRecent) {
              const f = await fa.openRecent(r);
              if (f) {
                const { openFromData } = await import('../state/actions/file');
                await openFromData(f.name, f.data, f.handle);
                return;
              }
            }
            openFile();
          }}
        >
          <FileText size={18} className="text-[#107C41]" />
          <span className="flex-1 truncate">{r.name}</span>
          <span className="opacity-60 text-[11px]">{new Date(r.opened).toLocaleDateString()}</span>
        </button>
      ))}
    </div>
  );
}

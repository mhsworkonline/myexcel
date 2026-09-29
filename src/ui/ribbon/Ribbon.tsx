'use client';
import { ChevronUp, MessageSquare, Share2, ChevronDown } from 'lucide-react';
import { RibbonTab, setState, S, useStore } from '../../state/store';
import { tableAtActive } from '../../state/actions/data';
import { HomeTab } from './HomeTab';
import { DataTab, FormulasTab, InsertTab, PageLayoutTab, ReviewTab, TableDesignTab, ViewTab } from './OtherTabs';
import { useState } from 'react';

const TABS: RibbonTab[] = ['File', 'Home', 'Insert', 'Page Layout', 'Formulas', 'Data', 'Review', 'View'];

export function Ribbon() {
  const tab = useStore((s) => s.ribbonTab) as string;
  const collapsed = useStore((s) => s.ribbonCollapsed);
  useStore((s) => s.sel);
  useStore((s) => s.rev);
  const [peek, setPeek] = useState(false);
  const table = tableAtActive();
  const showBody = !collapsed || peek;
  const current = tab === 'Table Design' && !table ? 'Home' : tab;
  return (
    <div className="flex flex-col px-2 pb-1" style={{ background: 'var(--chrome-bg)' }} data-testid="ribbon">
      <div className="flex items-center h-[32px] gap-[2px]" role="tablist">
        {TABS.map((t) => (
          <button
            key={t}
            role="tab"
            aria-selected={current === t}
            data-testid={`tab-${t}`}
            className={'xl-tab ' + (current === t && t !== 'File' ? 'xl-tab-active' : '')}
            onClick={() => {
              if (t === 'File') {
                setState({ backstage: true });
                return;
              }
              if (collapsed) setPeek(!(peek && current === t));
              setState({ ribbonTab: t });
            }}
            onDoubleClick={() => t !== 'File' && setState({ ribbonCollapsed: !collapsed })}
          >
            {t}
          </button>
        ))}
        {table && (
          <button
            className={'xl-tab ' + (current === 'Table Design' ? 'xl-tab-active' : '')}
            style={{ color: current === 'Table Design' ? undefined : 'var(--accent)' }}
            onClick={() => setState({ ribbonTab: 'Table Design' as RibbonTab })}
          >
            Table Design
          </button>
        )}
        <div className="flex-1" />
        <button className="xl-btn-sm !h-[28px] !px-2 border !border-[var(--border-strong)] rounded-[4px] mr-1" title="Comments" onClick={() => import('../../state/store').then((m) => m.openDialog('note', { edit: true, threaded: true }))}>
          <MessageSquare size={15} /> Comments
        </button>
        <button className="flex items-center gap-1.5 h-[28px] px-3 rounded-[4px] text-white" style={{ background: 'var(--accent)' }} title="Share" onClick={() => import('../../state/actions/file').then((m) => m.saveAs('xlsx'))}>
          <Share2 size={14} /> Share <ChevronDown size={12} />
        </button>
      </div>
      {showBody && (
        <div className="xl-ribbon flex items-stretch h-[96px] relative overflow-x-auto overflow-y-hidden" onClick={() => collapsed && setTimeout(() => setPeek(false), 0)}>
          <div className="flex items-stretch py-[2px]">
            {current === 'Home' && <HomeTab />}
            {current === 'Insert' && <InsertTab />}
            {current === 'Page Layout' && <PageLayoutTab />}
            {current === 'Formulas' && <FormulasTab />}
            {current === 'Data' && <DataTab />}
            {current === 'Review' && <ReviewTab />}
            {current === 'View' && <ViewTab />}
            {current === 'Table Design' && <TableDesignTab />}
          </div>
          <div className="flex-1" />
          <button className="absolute right-1 bottom-1 w-5 h-5 flex items-center justify-center rounded hover:bg-[var(--hover)]" title="Collapse the Ribbon (Ctrl+F1)" onClick={() => setState({ ribbonCollapsed: !S().ribbonCollapsed })}>
            <ChevronUp size={14} className={collapsed ? 'rotate-180' : ''} />
          </button>
        </div>
      )}
    </div>
  );
}

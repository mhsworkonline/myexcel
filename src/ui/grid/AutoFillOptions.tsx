'use client';
import { ClipboardList } from 'lucide-react';
import { useState } from 'react';
import type { Range } from '../../model/address';
import { autoFill, flashFill, FillMode } from '../../state/actions/fill';
import { undo } from '../../state/store';
import { Menu } from '../Menu';

/** The Auto Fill Options button shown after dragging the fill handle. */
export function AutoFillOptions({ x, y, src, target, onClose }: { x: number; y: number; src: Range; target: Range; onClose: () => void }) {
  const [open, setOpen] = useState(false);
  const [mode, setMode] = useState<FillMode>('auto');
  const redoWith = (m: FillMode) => () => {
    undo();
    autoFill(src, target, m);
    setMode(m);
  };
  return (
    <>
      <button
        className="absolute z-20 flex items-center gap-0.5 border px-0.5 h-[20px] shadow-sm"
        style={{ left: x, top: y, background: 'var(--panel)', borderColor: 'var(--border-strong)' }}
        onPointerDown={(e) => e.stopPropagation()}
        onClick={() => setOpen(true)}
        title="Auto Fill Options"
      >
        <ClipboardList size={14} />
        <span className="text-[8px]">▼</span>
      </button>
      {open && (
        <Menu
          x={x + 20}
          y={y + 60}
          onClose={() => {
            setOpen(false);
            onClose();
          }}
          items={[
            { label: 'Copy Cells', checked: mode === 'copy', onClick: redoWith('copy') },
            { label: 'Fill Series', checked: mode === 'auto' || mode === 'series', onClick: redoWith('series') },
            { label: 'Fill Formatting Only', checked: mode === 'formats', onClick: redoWith('formats') },
            { label: 'Fill Without Formatting', checked: mode === 'noFormats', onClick: redoWith('noFormats') },
            { separator: true },
            { label: 'Fill Days', onClick: redoWith('days') },
            { label: 'Fill Weekdays', onClick: redoWith('weekdays') },
            { label: 'Fill Months', onClick: redoWith('months') },
            { label: 'Fill Years', onClick: redoWith('years') },
            { separator: true },
            { label: 'Flash Fill', onClick: () => { undo(); flashFill(); } },
          ]}
        />
      )}
    </>
  );
}

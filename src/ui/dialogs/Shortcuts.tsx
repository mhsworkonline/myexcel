'use client';
import { Dialog } from './Dialog';

const GROUPS: [string, [string, string][]][] = [
  ['Workbook', [['Ctrl+N', 'New workbook'], ['Ctrl+O', 'Open'], ['Ctrl+S', 'Save (writes back to the opened file)'], ['F12', 'Save As'], ['Ctrl+P', 'Print preview'], ['Alt+F', 'File tab']]],
  ['Editing', [['F2', 'Edit cell'], ['Enter / Tab', 'Commit and move'], ['Esc', 'Cancel edit'], ['Alt+Enter', 'New line in cell'], ['Ctrl+Enter', 'Fill selection'], ['F4', 'Toggle $ references / repeat'], ['Ctrl+Z / Ctrl+Y', 'Undo / Redo'], ['Ctrl+C / X / V', 'Copy / Cut / Paste'], ['Ctrl+Alt+V', 'Paste Special'], ['Ctrl+D / Ctrl+R', 'Fill down / right'], ['Ctrl+E', 'Flash Fill'], ['Alt+=', 'AutoSum'], ['Ctrl+;', 'Insert date']]],
  ['Navigation', [['Ctrl+Arrow', 'Edge of data'], ['Ctrl+Home / Ctrl+End', 'First / last cell'], ['Ctrl+PgUp / PgDn', 'Previous / next sheet'], ['Ctrl+A', 'Select region / all'], ['Ctrl+G, F5', 'Go To'], ['Ctrl+F / Ctrl+H', 'Find / Replace']]],
  ['Formatting', [['Ctrl+B / I / U', 'Bold / Italic / Underline'], ['Ctrl+1', 'Format Cells'], ['Ctrl+Shift+$ % #', 'Currency / Percent / Date'], ['Ctrl+Shift+L', 'Filter'], ['Ctrl+T', 'Create table'], ['Ctrl+K', 'Hyperlink']]],
];

export function ShortcutsDialog() {
  return (
    <Dialog title="Keyboard Shortcuts" width={640} cancelLabel={null} okLabel="Close" testId="shortcuts">
      <div className="grid grid-cols-2 gap-x-6 gap-y-3 max-h-[60vh] overflow-auto">
        {GROUPS.map(([g, items]) => (
          <div key={g}>
            <div className="font-semibold mb-1">{g}</div>
            {items.map(([k, d]) => (
              <div key={k} className="flex gap-3 py-[2px]">
                <span className="w-[140px] shrink-0 font-mono text-[11.5px]">{k}</span>
                <span className="opacity-85">{d}</span>
              </div>
            ))}
          </div>
        ))}
      </div>
    </Dialog>
  );
}

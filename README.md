# MyExcel

A local-first, single-user spreadsheet that looks and works like current Microsoft Excel. Runs in the browser (Next.js) or as a desktop app (Tauri). Your files never leave your machine.

## Setup

Requirements: Node.js 20+ (tested on 22). For the Windows desktop app: Rust stable with the MSVC toolchain, Visual Studio 2022 Build Tools (Desktop C++), and the WebView2 runtime (built into Windows 10/11).

```bash
npm install
npx playwright install chromium   # only for the end-to-end smoke test
```

## Run

| Command | What it does |
|---|---|
| `npm run dev` | Web development server on http://localhost:3100 |
| `npm run build` | Static export of the whole app to `out/` (used by both web and desktop) |
| `npm start` | Serve `out/` locally on http://localhost:3100 |
| `npm run typecheck` | TypeScript check |
| `npm test` | Unit tests: xlsx/csv round-trip and undo/redo |
| `npm run e2e` | Playwright smoke test: open, edit, format, sort, undo, save |
| `npm run tauri:dev` | Desktop app against the dev server (hot reload) |
| `npm run tauri:build` | Release desktop build + installers in `src-tauri/target/release/bundle/` (`nsis/*.exe`, `msi/*.msi`) |

One-off checks (not part of the test suite): `node scripts/scroll-check.mjs` (200k-row scroll performance, needs `npm start`) and `node scripts/desktop-check.mjs` (drives the packaged desktop app over WebView2 remote debugging).

## Desktop app (Windows)

Everything stays on your computer; the app makes no network requests.

- **Files:** native Open/Save dialogs. Ctrl+S writes straight back to the opened file (`.xlsx`, `.csv`, `.tsv`, `.ods`); `.xls` opens read-only and saves as `.xlsx`. Double-clicking an associated file (`.xlsx`, `.xls`, `.csv`, `.ods`) opens it in MyExcel.
- **Where data lives** (`%APPDATA%\app.myexcel.desktop\`): `settings.json` holds theme, default zoom, default font and recent files; `recovery\snapshot.json` holds the crash-recovery copy (written every 30 s while there are unsaved changes, removed when you save or discard).
- **Recent files:** File > Open (and the File tab's Home page), or the menu's File > Open Recent.
- **Options:** File tab > Options (theme, default font, default zoom).
- **Closing** with unsaved changes asks Save / Don't Save / Cancel.
- **Clipboard:** copy/paste exchanges formatted data with Excel in both directions.

## Features

- **Grid**: canvas-rendered and virtualized (1,048,576 × 16,384), frozen panes, split view, zoom 10–400%, hidden rows/columns, merges, text overflow/wrap/rotation, all border styles, dark mode.
- **Editing**: in-cell editor with colored formula references, point mode (click/arrow to build references, including other sheets), F4 reference toggle, function autocomplete and argument hints, Alt+Enter line breaks, Ctrl+Enter fill selection.
- **Formulas**: every HyperFormula function (~390), named ranges, cross-sheet references, error values, circular-reference detection.
- **Formatting**: Home ribbon (fonts, colors, borders, alignment, indent, wrap, merge, number formats), Format Cells dialog (Number, Alignment, Font, Border, Fill, Protection), cell styles, Format Painter, custom number codes, **INR with lakh/crore grouping** (`[$₹-4009] #,##0.00`).
- **Data**: fill handle (series, dates, months, linear trends), Flash Fill, copy/cut/paste with Excel HTML interop, Paste Special (values, formats, operations, transpose, link), sort (multi-level, custom lists, by color), AutoFilter (values, text/number conditions, color, top 10, above/below average), Find & Replace with wildcards, Go To / Go To Special, Remove Duplicates, Text to Columns.
- **Data tools**: conditional formatting (highlight rules, top/bottom, averages, duplicates, text/date rules, data bars, color scales, icon sets, formula rules, rule manager), data validation (lists with dropdowns, number/date/time/length/custom, input messages, error alerts, Circle Invalid Data), tables (60 styles, banding, total row), Name Manager, notes and threaded comments, hyperlinks, cell locking and sheet protection.
- **Analysis**: PivotTables (field list with drag and drop, filters, date/number grouping, calculated fields, refresh), charts (column, bar, line, pie, doughnut, area, scatter, combo with secondary axis), sparklines, Goal Seek, Solver, Scenario Manager, Data Tables.
- **Output**: Page Setup, print preview, manual page breaks, print titles, headers/footers, print to printer, PDF export.
- **Files**: `.xlsx` (read/write, including charts, conditional formats, validations, tables, names), `.csv`/`.tsv`, `.ods` (read/write), `.xls` (read). Auto-save to IndexedDB with recovery after a crash.
- **Undo/redo** covers every edit, including structural changes and sheet operations.

## Keyboard shortcuts

### Navigation and selection
| Keys | Action |
|---|---|
| Arrow keys | Move one cell |
| Ctrl+Arrow | Jump to the edge of the data region |
| Shift+Arrow / Ctrl+Shift+Arrow | Extend selection |
| Tab / Shift+Tab | Move right / left (Enter after Tab returns to the starting column) |
| Enter / Shift+Enter | Move down / up (cycles inside a selection) |
| Home / Ctrl+Home / Ctrl+End | Start of row / top-left / last used cell |
| PageUp / PageDown, Alt+PageUp / Alt+PageDown | Scroll a screen up/down, left/right |
| Ctrl+PageUp / Ctrl+PageDown | Previous / next sheet |
| Ctrl+A | Select the current region, then the whole sheet |
| Ctrl+Space / Shift+Space | Select column / row |
| F8 / Shift+F8 | Extend / add-to-selection mode |
| Ctrl+G or F5 | Go To |

### Editing
| Keys | Action |
|---|---|
| F2 | Edit the active cell (toggles Enter/Edit mode while editing) |
| F4 | While editing: cycle absolute/relative reference; otherwise repeat last action |
| Esc | Cancel editing / clear the copy marquee |
| Alt+Enter | New line in the cell |
| Ctrl+Enter | Fill the selection with the entry |
| Delete / Backspace | Clear contents / clear and edit |
| Ctrl+; / Ctrl+Shift+; | Insert current date / time |
| Ctrl+' / Ctrl+Shift+" | Copy formula / value from the cell above |
| Ctrl+D / Ctrl+R | Fill down / right |
| Ctrl+E | Flash Fill |
| Alt+= | AutoSum |
| Shift+F3 | Insert Function |
| Alt+Down | Open validation list |
| Ctrl+Z / Ctrl+Y | Undo / redo |
| Ctrl+C / Ctrl+X / Ctrl+V | Copy / cut / paste |

### Formatting
| Keys | Action |
|---|---|
| Ctrl+B / Ctrl+I / Ctrl+U / Ctrl+5 | Bold / italic / underline / strikethrough |
| Ctrl+1 | Format Cells |
| Ctrl+Shift+~ ! @ # $ % ^ | General / number / time / date / currency / percent / scientific |
| Ctrl+Shift+& / Ctrl+Shift+_ | Outline border / remove borders |
| Ctrl+9 / Ctrl+0 (with Shift to unhide) | Hide rows / columns |

### Workbook
| Keys | Action |
|---|---|
| Ctrl+N / Ctrl+O / Ctrl+S / F12 | New / open / save / save as |
| Ctrl+P | Print preview |
| Ctrl+F / Ctrl+H | Find / replace |
| Ctrl+K | Insert hyperlink |
| Ctrl+T or Ctrl+L | Create table |
| Ctrl+Shift+L | Toggle AutoFilter |
| Ctrl+- / Ctrl++ | Delete / insert cells |
| Ctrl+` | Show formulas |
| Shift+F11 | Insert sheet |
| F11 | Insert chart |
| F9 | Recalculate |
| Shift+F2 | Edit note |
| Shift+F10 | Context menu |
| Ctrl+F1 | Collapse the ribbon |
| Alt+F | File (backstage) |

## Architecture

```
src/engine   HyperFormula wrapper (the only module that imports HyperFormula; no React)
src/model    cells, styles, sheets, workbook, selection, number formats, formula tokenizer, command layer
src/state    Zustand store and all user actions (every edit goes through Tx → History)
src/io       FileAdapter (browser / Tauri), xlsx, csv, ods, xls, autosave
src/ui       React UI: ribbon, grid (canvas), dialogs, charts, pivot pane, print preview
src-tauri    Desktop shell (native menus, dialogs, file associations)
```

See `DECISIONS.md` for design decisions and `KNOWN_GAPS.md` for what differs from Excel.

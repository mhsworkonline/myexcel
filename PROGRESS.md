# Progress

Resume guide: `npm install`, `npm run build`, `./scripts/restart.sh` (serves :3100), `node scripts/shot.mjs <name> '<actions>'` for screenshots, `npx vitest run`, `npx playwright test`.

## Phase 0 – Scaffold
- [x] Next.js 15 + TS 5.9 + Tailwind 3 + Zustand + HyperFormula + ExcelJS + Papaparse + Vitest + Playwright
- [x] DECISIONS.md / PROGRESS.md

## Phase 1 – Core  ✅
- [x] model: A1 addresses, sparse sheet storage (1,048,576 × 16,384), deduplicated style table, number formats (all categories + custom codes + INR lakh/crore), input parsing, formula tokenizer, structural reference adjustment
- [x] command layer (Tx / History) — every edit undoable incl. structural ops & sheet ops
- [x] engine: HyperFormula wrapper, lazy attach, names, circular detection
- [x] io: FileAdapter + BrowserFileAdapter, xlsx (values/formulas/styles/widths/heights/freeze/merges/names/CF/DV/tables/protection/page setup), csv/tsv, IndexedDB autosave + recovery
- [x] UI shell: title bar + QAT, ribbon (all tabs; Home complete with launchers), name box, formula bar (expandable), sheet tabs, status bar, backstage
- [x] canvas grid: virtualized, frozen panes, split view, zoom, DPR-aware, text overflow, wrap, rotation, borders (all styles), merges, headers highlight, fill handle, marching ants
- [x] selection (single/range/multi/rows/cols/all), full keyboard map, point mode, F4, autocomplete + arg hints
- [x] fill handle (copy/series/dates/months/flash fill lite), clipboard (internal + HTML/TSV interop), paste special incl. transpose
- [x] formatting + Format Cells (6 tabs), Format Painter, cell styles
- [x] sort (multi-level/custom lists/colors), AutoFilter (values/custom/color/top10/average), Find & Replace, Go To, Go To Special
- [x] sheets: add/rename/delete/move/copy/color/hide; insert/delete rows/cols/cells; hide/unhide; autofit
- [x] tests: xlsx+csv round-trip, undo/redo (6 tests, ~2s); Playwright smoke passes

## Phase 2 – Data tools  ✅
- [x] conditional formatting engine + renderer (highlight, top/bottom, averages, duplicates, text/date, data bars, color scales, icon sets, formula rules) + quick-rule dialogs, new/edit rule, rule manager
- [x] data validation (list dropdown, number/date/time/text-length/custom, input messages, stop/warning/info alerts, circle invalid)
- [x] named ranges + Name Manager + create from selection + name box define
- [x] notes / threaded comments, hyperlinks
- [x] tables (styles gallery, banding, total row, Table Design tab, auto-expand), Remove Duplicates, Text to Columns wizard, paste transpose
- [x] cell locking + sheet protection (password hash, allow-list)
- [x] visual verification pass + smoke test at end of phase

## Phase 3 – Analysis & output  ✅
- [x] pivot tables (field list with drag & drop, filters, date/number grouping, calculated fields, refresh, PivotTable Analyze tab)
- [x] charts (column/bar/line/pie/doughnut/area/scatter/combo + secondary axis) with Chart Design tab, select data, format
- [x] sparklines (dialog + Sparkline tab), Goal Seek, Solver-lite, Scenario Manager (+ summary), Data Tables
- [x] page setup, print preview, manual page breaks, headers/footers, print titles, PDF export (jsPDF)
- [x] xlsx round-trip for charts (native DrawingML), CF, DV, tables, names incl. constants/scoped, MyExcel meta part; .xls read (SheetJS); .ods read/write

## Phase 4 – Desktop  ✅
- [x] Tauri v2 wrapper (cargo build verified, app launches), TauriFileAdapter (native open/save, save-in-place), recent files, file associations (xlsx/xlsm/xls/csv/tsv/ods), native menus bridged to the UI; web build unchanged

## Finish  ✅
- [x] 200k-row scroll check (`scripts/scroll-check.mjs`): median 16.7 ms / max 16.8 ms per frame
- [x] README, KNOWN_GAPS.md

## Desktop app (Tauri v2, Windows) — second pass
Prerequisites checked: Rust 1.96.1 MSVC ✓, VS 2022 Build Tools ✓, WebView2 154 ✓, Node 22.16 ✓.
- [x] Audit existing `src-tauri` / `TauriFileAdapter` and finish it
- [x] Static export for all builds; `npm start` serves `out/`; web dev/build/test/smoke still pass
- [x] Window 1440×900, min 900×600; icons; scripts `tauri:dev`, `tauri:build`
- [x] TauriFileAdapter: native Open/Save/Save As, Ctrl+S in place (.xlsx/.csv/.tsv/.ods), .xls read-only → Save As
- [x] Recent files (settings.json in app dir) shown in File > Open, File backstage, native "Open Recent..."
- [x] Recovery snapshots every 30 s while dirty in app data dir (web: IndexedDB via adapter), restore offer after crash
- [x] Settings (theme, default zoom, default font, recent files) in app config dir; File > Options page
- [x] File associations (.xlsx/.xlsm, .xls, .csv, .ods) + opening the file passed at launch (scope granted in Rust)
- [x] Native Save/Don't Save/Cancel prompt on close
- [x] Strict CSP, least-privilege capabilities, persisted scope for picked files, no network plugins, fonts bundled, telemetry off
- [x] Native menu (File/Edit/View/Help) → same commands as ribbon; WebView2 browser shortcuts disabled
- [x] Native window title mirrors file name + save status; in-app title row kept
- [x] Native clipboard (CF_HTML + text) for ribbon/menu copy/paste; Ctrl+C/V via webview events
- [x] Release build + NSIS/MSI installers; `scripts/desktop-check.mjs` passes 15/15 against the packaged exe (launch-file open, in-place save, settings/recent, 30 s recovery + crash offer, CF_HTML clipboard, title, CSP, close prompt)

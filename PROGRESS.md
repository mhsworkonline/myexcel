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

## Phase 3 – Analysis & output
- [ ] pivot tables (field list, filters, grouping, calculated fields, refresh)
- [ ] charts (column/bar/line/pie/area/scatter/combo) + chart tools
- [ ] sparklines UI (renderer done), Goal Seek, Solver-lite, Scenario Manager, Data Tables
- [ ] page setup, print preview, page breaks, headers/footers, PDF export
- [ ] xlsx charts round-trip + constant names; .xls read; .ods read/write

## Phase 4 – Desktop
- [ ] Tauri wrapper, TauriFileAdapter, recent files, file associations, native menus

## Finish
- [ ] 200k-row scroll check script
- [ ] README, KNOWN_GAPS.md

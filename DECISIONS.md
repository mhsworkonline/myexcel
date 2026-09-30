# Design decisions

Decisions made autonomously while building MyExcel. Newest at the bottom.

1. **TypeScript 5.9 pinned** instead of TypeScript 7 (native port) because Next 15's build-time type checker targets the TS 5 API.
2. **HyperFormula license key `gpl-v3`.** HyperFormula is dual-licensed; using it under GPLv3 means MyExcel as distributed is GPLv3-compatible. A commercial HF license would be needed for closed-source distribution.
3. **Model is the source of truth for cell contents**; the engine mirrors it. Constant cells are read straight from the model, and only formula cells ask the engine. This keeps the UI fast and lets the engine be optional.
4. **Lazy engine attach.** Benchmarks show HyperFormula takes ~1.8 s and ~300 MB to load 1M cells. The engine is built only once the workbook contains a formula or a defined name, so large pure-data files (200k rows) open instantly.
5. **Structural edits (insert/delete rows/cols)**: the model shifts its sparse maps and HyperFormula adjusts formulas, then formula text is synced back from the engine. Undo of a delete = inverse insert + restore removed cells + restore every formula whose text changed.
6. **Styles** are flat objects deduplicated in a workbook-wide style table (`StyleTable`, id 0 = default). Whole-row/whole-column formatting is stored as row/column default styles, so formatting column A never materialises 1M cells.
7. **Default font** is Calibri 11 (with Carlito/Segoe UI fallbacks). No Microsoft fonts are bundled; Calibri is used only if installed on the system.
8. **Scroll extent** grows like Excel: used range plus a buffer, extended as the user scrolls, so the native scrollbar stays usable within browser element-size limits.
9. **INR / lakh-crore grouping** is triggered by the locale tag `[$₹-4009]` (en-IN) in a number format, matching how Excel stores Indian currency formats in .xlsx.
10. **Circular references** display 0 (as Excel does) and the status bar shows "Circular References: <cell>".
11. **Formula reference shifting** (copy/paste, fill, sort) is done by our own tokenizer in `model/formula.ts` rather than HyperFormula's clipboard, so the model stays the source of truth.
12. **`src/state/` application layer** sits between `model/` and `ui/`: the Zustand store plus all user actions. Actions build a `Tx` (transaction) which applies changes immediately and records inverse data, then commit one undoable entry. The UI calls actions and reads values through `state/values.ts`; only `engine/` imports HyperFormula.
13. Layout caches (column/row offsets) key on the identity of the size maps. Size maps are never mutated in place; every change replaces the map through the command layer.
14. **Extra dependencies for Phase 3**: SheetJS CE 0.20.3 (from cdn.sheetjs.com, the maintained build — the npm `xlsx@0.18.5` has known vulnerabilities) used *only* to read legacy `.xls`; JSZip for `.ods` read/write and post-processing ExcelJS output (charts, constant names, MyExcel metadata); jsPDF for PDF export.
15. **Pivot tables are computed snapshots** written into cells (like Excel's pivot cache): they update on Refresh or when their field layout changes, not on every source edit.
16. **Charts render as SVG overlays** positioned in sheet coordinates (px at 100% zoom). In `.xlsx` they are written as native DrawingML charts (column/bar/line/pie/doughnut/area/scatter/combo) so Excel opens them; MyExcel-only settings (style, colors) also go into a custom XML part for lossless round-trip.
17. **MyExcel metadata** that has no ExcelJS support (pivot definitions, sparklines, scenarios, full chart specs) is stored in `customXml/myexcelN.xml`, a standard custom XML part Excel preserves and ignores.
18. **Printing / PDF** reuses the canvas grid renderer: each page is rendered with a print viewport (print titles act like frozen panes), headers/footers are drawn in the margins, and jsPDF embeds the page images. Text in the PDF is therefore raster (not selectable).
19. **What-if Data Tables** are written as static values (recomputed when the dialog is re-run) rather than a live `{=TABLE()}` array.

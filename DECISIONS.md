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

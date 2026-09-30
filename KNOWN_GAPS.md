# Known gaps

Everything below works differently from Microsoft Excel or isn't implemented. Items are grouped by area.

## Calculation
- Formulas are evaluated by HyperFormula. Functions it doesn't implement (for example `LET`, `LAMBDA`, `XLOOKUP` match/search modes beyond the basics, `TEXTSPLIT`, cube and web functions) return `#NAME?`.
- Calculation is always automatic; the Manual calculation option is shown but disabled.
- Dynamic-array spill ranges are evaluated but not drawn with the spill border, and `#SPILL!` layout rules may differ.
- Circular references show 0 and a status-bar warning; iterative calculation isn't available.
- Dates use the 1900 system without Excel's fictitious 29 Feb 1900, so serials before March 1900 are off by one. The 1904 date system isn't supported.
- Structured table references (`Table1[Column]`) aren't resolved in formulas.

## Grid and editing
- Rich text (mixed formatting within one cell) is flattened to plain text on import.
- Shrink to fit, distributed alignment and justify are stored but drawn as normal alignment.
- Pattern fills (other than solid) are stored and saved but drawn as solid color.
- Text rotation is drawn for any angle, but row heights don't auto-grow for rotated text.
- Split view: the top/left panes scroll with the mouse wheel only (the scrollbars drive the bottom/right panes).
- Multi-sheet selection (grouped sheets) isn't supported.
- Outline/grouping of rows and columns isn't implemented.
- Shapes, pictures, text boxes and SmartArt aren't supported.

## Clipboard
- Pasting from Excel uses the HTML/text formats that Excel puts on the clipboard: values, number formats, fonts, fills, borders and merges come across; formulas from another Excel instance arrive as values. Within MyExcel, formulas paste with adjusted references.
- The ribbon's Paste button needs browser clipboard permission; Ctrl+V always works.

## Data tools
- AutoFilter: one filter per sheet (tables share it). Date grouping in the filter list and "Filter by icon" aren't available.
- Sorting doesn't move row heights or row-level formatting with the data.
- Flash Fill recognises token/case/initial patterns; it doesn't learn arbitrary substring positions.
- Conditional formatting icon sets use equal percentage bands in the quick gallery; custom per-icon thresholds are kept from files but can't be edited individually.
- Sheet protection uses a SHA-256 hash; Excel-protected passwords from files aren't verified (the sheet can be unprotected without the original password). Workbook-structure protection isn't implemented.
- Threaded comments are stored as notes with replies; @mentions and resolve aren't supported.

## PivotTables
- Output is a formatted snapshot in cells (Compact layout only); Tabular/Outline layouts, subtotal placement options, "Show Values As", slicers and timelines aren't available.
- Pivot definitions round-trip between MyExcel files (custom XML part). Excel sees the pivot output as plain cells, and PivotTables in files made by Excel open as plain cells.
- Calculated fields support arithmetic and functions over the sums of fields.

## Charts
- Types: column, bar (clustered and stacked), line (with and without markers), pie, doughnut, area (and stacked), scatter, combo with secondary axis. 3-D, radar, stock, surface, histogram, waterfall, treemap, sunburst and map charts aren't supported.
- Formatting is limited to title, axis titles, legend position, data labels, gridlines, color palettes and per-series colors. Trendlines, error bars and axis scaling options aren't available.
- Charts from Excel files are read from their DrawingML data; unsupported formatting is dropped.
- A chart on a frozen pane scrolls with the main pane.

## Printing and PDF
- Pages are rendered as images, so text in exported PDFs isn't selectable or searchable.
- Header/footer codes supported: page, pages, date, time, file name, sheet name. Pictures and per-section fonts in headers/footers aren't.
- Page Break Preview shows automatic and manual breaks but breaks can't be dragged; manual column breaks can only come from files.
- "Print Selection" prints the last selected range only.

## Files
- **Opening large files**: parsing is done by ExcelJS in the browser. A 200,000-row × 6-column workbook takes about 12 s to open (scrolling afterwards runs at 60 fps).
- `.xls` is read-only and imports values, formulas, number formats, merges, sizes, notes and links (no fonts/fills/borders). SheetJS CE is used only for this.
- `.ods` round-trips values, formulas, basic fonts/fills/alignment, merges, column widths, row heights, notes and links; conditional formats, validation, charts and number-format codes aren't written to ODS.
- Unsupported xlsx features are dropped on save: VBA/macros (`.xlsm` opens without macros), external links, Power Query, slicers, sparklines written by Excel (MyExcel's own sparklines round-trip via its metadata part), form controls, themes (theme colors are converted to RGB).
- Theme fonts and colors from files are resolved to the default Office theme.
- CSV export writes the displayed (formatted) values of the active sheet, like Excel.

## Fonts and look
- No Microsoft fonts are bundled. Calibri is used if installed (Windows); otherwise Carlito/Segoe UI/Arial, so column widths and text metrics can differ slightly from Excel.
- Icons are Lucide or drawn from scratch, so they resemble but don't match Excel's icons.
- The ribbon doesn't adapt (collapse groups) to narrow windows; it scrolls horizontally instead.
- Office Add-ins, Copilot, Analyze Data, co-authoring, version history and cloud sharing aren't part of a local-first single-user app. The Share button saves a copy.

## Desktop (Tauri)
- Native menu covers File/Edit/View/Help commands; the ribbon remains the primary UI.
- Opening a file via file association works on launch; opening a second file while the app is already running starts a second window/instance (no single-instance handoff).
- Recent files are kept per installation in local storage.

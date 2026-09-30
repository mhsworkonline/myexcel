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

## Desktop (Tauri, Windows)
- Only one workbook per window, and a single window: opening a second file (double-click while running) starts a second instance rather than a new window in the first.
- Native menu items show their shortcuts as labels; the shortcuts themselves are handled by the web UI (so they only act while the MyExcel window is focused, which is always the case for a menu).
- Files opened through Recent/launch are accessed with scopes granted at pick time (persisted across restarts). A file moved or renamed outside MyExcel must be reopened through File > Open.
- The recovery snapshot is one file per installation: if MyExcel crashes with two instances open, only the last snapshot written is offered.
- Installers are unsigned (Windows SmartScreen will warn). Code signing needs a certificate.
- The WebView2 bootstrapper in the installer downloads the runtime only on machines that don't have it (all current Windows 10/11 installs do).
- macOS/Linux builds aren't configured or tested; the native clipboard commands are Windows-only (other platforms fall back to the web clipboard).
- Default font from Options applies to new workbooks; files opened from disk keep their own Normal font (Calibri).

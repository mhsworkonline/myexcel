// Excel-like cursors as inline SVG data URIs (drawn from scratch, not Microsoft assets).
const svg = (s: string) => `url("data:image/svg+xml;utf8,${encodeURIComponent(s)}")`;

/** Thick white plus with black outline, used over cells. */
export const CELL_CURSOR =
  svg(`<svg xmlns='http://www.w3.org/2000/svg' width='20' height='20'><path d='M7 1h6v6h6v6h-6v6H7v-6H1V7h6z' fill='white' stroke='black' stroke-width='1'/></svg>`) + ' 10 10, cell';

/** Small black down arrow over column headers. */
export const COL_HEADER_CURSOR =
  svg(`<svg xmlns='http://www.w3.org/2000/svg' width='16' height='16'><path d='M6 1h4v8h3l-5 6-5-6h3z' fill='black' stroke='white' stroke-width='1'/></svg>`) + ' 8 15, s-resize';

/** Small black right arrow over row headers. */
export const ROW_HEADER_CURSOR =
  svg(`<svg xmlns='http://www.w3.org/2000/svg' width='16' height='16'><path d='M1 6v4h8v3l6-5-6-5v3z' fill='black' stroke='white' stroke-width='1'/></svg>`) + ' 15 8, e-resize';

/** Thin black cross for the fill handle. */
export const FILL_CURSOR =
  svg(`<svg xmlns='http://www.w3.org/2000/svg' width='16' height='16'><path d='M7 0h2v7h7v2H9v7H7V9H0V7h7z' fill='black'/></svg>`) + ' 8 8, crosshair';

export const MOVE_CURSOR = 'move';
export const COL_RESIZE_CURSOR = 'col-resize';
export const ROW_RESIZE_CURSOR = 'row-resize';

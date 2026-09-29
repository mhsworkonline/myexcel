// Tiny event bus so the grid, scrollbars and sheet-tab bar stay in sync without React re-renders.
type Fn = () => void;
const subs = new Set<Fn>();

export function onScrollChange(fn: Fn): () => void {
  subs.add(fn);
  return () => subs.delete(fn);
}

export function emitScroll(): void {
  for (const f of subs) f();
}

export interface GridApi {
  /** Scroll the main pane horizontally/vertically (zoomed px). */
  setScroll(axis: 'x' | 'y', value: number): void;
  extent(axis: 'x' | 'y'): { pos: number; view: number; total: number };
  redraw(): void;
}

let api: GridApi | null = null;
export function setGridApi(a: GridApi | null): void {
  api = a;
}
export function gridApi(): GridApi | null {
  return api;
}

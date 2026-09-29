/** Return keyboard focus to the grid (Excel does this after ribbon clicks and Name Box navigation). */
export function focusGrid(): void {
  const el = document.querySelector('[data-testid="grid-input"]') as HTMLElement | null;
  el?.focus({ preventScroll: true });
}

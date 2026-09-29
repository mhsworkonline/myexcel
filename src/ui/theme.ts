// Colour palettes for the canvas grid (CSS chrome uses the matching variables in globals.css).

export interface GridPalette {
  bg: string;
  grid: string;
  headerBg: string;
  headerText: string;
  headerLine: string;
  headerSelBg: string;
  headerSelText: string;
  headerFullBg: string;
  headerFullText: string;
  accent: string;
  selFill: string;
  text: string;
  cornerTriangle: string;
  frozenLine: string;
  splitBar: string;
  pageBreak: string;
  filterBtnBg: string;
  filterBtnBorder: string;
  noteMark: string;
  hyperlink: string;
}

export const LIGHT: GridPalette = {
  bg: '#FFFFFF',
  grid: '#D4D4D4',
  headerBg: '#F5F5F5',
  headerText: '#444444',
  headerLine: '#D4D4D4',
  headerSelBg: '#D3F0E0',
  headerSelText: '#0E5C2F',
  headerFullBg: '#9FD5B7',
  headerFullText: '#0B4D27',
  accent: '#107C41',
  selFill: 'rgba(20, 20, 20, 0.09)',
  text: '#000000',
  cornerTriangle: '#B4B4B4',
  frozenLine: '#9E9E9E',
  splitBar: '#C8C8C8',
  pageBreak: '#2F5597',
  filterBtnBg: '#FFFFFF',
  filterBtnBorder: '#A6A6A6',
  noteMark: '#E02020',
  hyperlink: '#0563C1',
};

export const DARK: GridPalette = {
  bg: '#1F1F1F',
  grid: '#3B3B3B',
  headerBg: '#2B2B2B',
  headerText: '#C8C8C8',
  headerLine: '#3F3F3F',
  headerSelBg: '#1D4A33',
  headerSelText: '#9FE3BD',
  headerFullBg: '#2E7D52',
  headerFullText: '#FFFFFF',
  accent: '#3FB871',
  selFill: 'rgba(255, 255, 255, 0.12)',
  text: '#E8E8E8',
  cornerTriangle: '#6A6A6A',
  frozenLine: '#7A7A7A',
  splitBar: '#4A4A4A',
  pageBreak: '#6F9BE8',
  filterBtnBg: '#2B2B2B',
  filterBtnBorder: '#6A6A6A',
  noteMark: '#FF5A5A',
  hyperlink: '#6CB4FF',
};

/** In dark mode, Excel shows "automatic" black text as light and white fills as dark. */
export function adaptColor(color: string | undefined, kind: 'text' | 'fill', dark: boolean): string | undefined {
  if (!dark || !color) return color;
  const c = color.toUpperCase();
  if (kind === 'text' && c === '#000000') return '#E8E8E8';
  if (kind === 'fill' && c === '#FFFFFF') return '#1F1F1F';
  return color;
}

// Excel table style rendering (Light / Medium / Dark families, approximated from the Office theme).
import type { Sheet } from '../model/sheet';
import type { BorderEdge } from '../model/styles';
import type { TableDef } from '../model/types';

export const ACCENTS = ['#000000', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];

export function tint(hex: string, t: number): string {
  const n = parseInt(hex.slice(1), 16);
  const ch = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => Math.round(t >= 0 ? v + (255 - v) * t : v * (1 + t)));
  return '#' + ch.map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
}

export interface TableCellLook {
  fill?: string;
  fontColor?: string;
  bold?: boolean;
  bTop?: BorderEdge;
  bBottom?: BorderEdge;
  bLeft?: BorderEdge;
  bRight?: BorderEdge;
}

export function parseTableStyle(name: string): { family: 'Light' | 'Medium' | 'Dark'; n: number } {
  const m = /TableStyle(Light|Medium|Dark)(\d+)/.exec(name);
  if (!m) return { family: 'Medium', n: 2 };
  return { family: m[1] as 'Light' | 'Medium' | 'Dark', n: parseInt(m[2], 10) };
}

export function tableAt(sheet: Sheet, r: number, c: number): TableDef | undefined {
  for (const t of sheet.tables) if (r >= t.range.r1 && r <= t.range.r2 && c >= t.range.c1 && c <= t.range.c2) return t;
  return undefined;
}

export function tableLook(t: TableDef, r: number, c: number): TableCellLook {
  const { family, n } = parseTableStyle(t.style);
  const accent = ACCENTS[(n - 1) % 7];
  const isGray = (n - 1) % 7 === 0;
  const header = t.headerRow && r === t.range.r1;
  const total = t.totalRow && r === t.range.r2;
  const dataIdx = r - t.range.r1 - (t.headerRow ? 1 : 0);
  const colIdx = c - t.range.c1;
  const bandRow = t.bandedRows && !header && !total && dataIdx % 2 === 0;
  const bandCol = t.bandedCols && !header && !total && colIdx % 2 === 0;
  const firstCol = t.firstCol && c === t.range.c1;
  const lastCol = t.lastCol && c === t.range.c2;
  const look: TableCellLook = {};
  const thin = (color: string): BorderEdge => ({ style: 'thin', color });
  if (family === 'Medium') {
    const light = isGray ? '#D9D9D9' : tint(accent, 0.8);
    if (n <= 7) {
      // Medium 1-7: banded with thin accent borders
      if (header) Object.assign(look, { fill: accent, fontColor: '#FFFFFF', bold: true });
      else if (bandRow || bandCol) look.fill = light;
      look.bTop = thin(tint(accent, 0.4));
      look.bBottom = thin(tint(accent, 0.4));
    } else if (n <= 14) {
      if (header) Object.assign(look, { fill: accent, fontColor: '#FFFFFF', bold: true });
      else look.fill = bandRow || bandCol ? light : tint(accent, 0.9);
      look.bBottom = thin('#FFFFFF');
      look.bRight = thin('#FFFFFF');
    } else {
      if (header) Object.assign(look, { fill: accent, fontColor: '#FFFFFF', bold: true });
      else if (bandRow || bandCol) look.fill = light;
      look.bBottom = thin(tint(accent, 0.4));
    }
    if (total) Object.assign(look, { bold: true, fill: n > 7 && n <= 14 ? accent : undefined, fontColor: n > 7 && n <= 14 ? '#FFFFFF' : undefined, bTop: { style: 'double', color: accent } });
  } else if (family === 'Light') {
    const light = isGray ? '#D9D9D9' : tint(accent, 0.8);
    if (n <= 7) {
      if (header) Object.assign(look, { bold: true, bBottom: thin(accent), bTop: thin(accent) });
      else if (bandRow || bandCol) look.fill = light;
      if (r === t.range.r2) look.bBottom = thin(accent);
    } else if (n <= 14) {
      if (header) Object.assign(look, { fill: accent, fontColor: '#FFFFFF', bold: true });
      look.bTop = thin(accent);
      look.bBottom = thin(accent);
      if (c === t.range.c1) look.bLeft = thin(accent);
      if (c === t.range.c2) look.bRight = thin(accent);
    } else {
      if (header) Object.assign(look, { bold: true, bBottom: thin(accent) });
      else if (bandRow || bandCol) look.fill = light;
      look.bTop = look.bTop ?? thin(accent);
      look.bBottom = look.bBottom ?? thin(accent);
      look.bLeft = thin(accent);
      look.bRight = thin(accent);
    }
    if (total) Object.assign(look, { bold: true, bTop: { style: 'double', color: accent } });
  } else {
    const dark = isGray ? '#404040' : tint(accent, -0.25);
    if (header) Object.assign(look, { fill: '#000000', fontColor: '#FFFFFF', bold: true });
    else Object.assign(look, { fill: bandRow || bandCol ? tint(dark, -0.2) : dark, fontColor: '#FFFFFF' });
    if (total) Object.assign(look, { fill: tint(dark, -0.5), fontColor: '#FFFFFF', bold: true, bTop: { style: 'double', color: '#FFFFFF' } });
  }
  if ((firstCol || lastCol) && !header) look.bold = true;
  return look;
}

export const TABLE_STYLE_NAMES = [
  ...Array.from({ length: 21 }, (_, i) => `TableStyleLight${i + 1}`),
  ...Array.from({ length: 28 }, (_, i) => `TableStyleMedium${i + 1}`),
  ...Array.from({ length: 11 }, (_, i) => `TableStyleDark${i + 1}`),
];

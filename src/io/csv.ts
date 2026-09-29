import Papa from 'papaparse';
import { parseInput } from '../model/parseInput';
import { Sheet } from '../model/sheet';
import type { Cell } from '../model/types';
import { Workbook } from '../model/workbook';

export function sheetNameFromFile(fileName: string): string {
  const base = fileName.replace(/\.[^.]+$/, '').replace(/[\\/?*[\]:]/g, '_').slice(0, 31);
  return base || 'Sheet1';
}

/** Parse CSV text into a workbook, converting numbers/dates like Excel does on open. */
export function readCsv(text: string, fileName = 'Book1.csv'): Workbook {
  const t = text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
  const res = Papa.parse<string[]>(t, { skipEmptyLines: false, delimiter: fileName.toLowerCase().endsWith('.tsv') ? '\t' : '' });
  const wb = new Workbook();
  const sheet = new Sheet(sheetNameFromFile(fileName));
  const rows = res.data;
  // drop the trailing empty line Papaparse yields for files ending with a newline
  while (rows.length && rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === '') rows.pop();
  const formatIds = new Map<string, number>();
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    for (let c = 0; c < row.length; c++) {
      const raw = row[c];
      if (raw === '' || raw === undefined) continue;
      const p = parseInput(raw);
      const cell: Cell = {};
      if (p.kind === 'formula') cell.f = p.formula;
      else if (p.kind === 'error') {
        cell.v = p.value;
        cell.e = true;
      } else if (p.kind !== 'empty') cell.v = p.value;
      if (p.autoFormat) {
        let id = formatIds.get(p.autoFormat);
        if (id === undefined) {
          id = wb.styles.intern({ numFmt: p.autoFormat });
          formatIds.set(p.autoFormat, id);
        }
        cell.s = id;
      }
      sheet.setCellRaw(r, c, cell);
    }
  }
  wb.sheets.push(sheet);
  wb.activeSheetId = sheet.id;
  return wb;
}

/** Serialise one sheet to CSV using the provided display-text function (Excel writes displayed values). */
export function writeCsv(sheet: Sheet, textAt: (r: number, c: number) => string, delimiter = ','): string {
  const used = sheet.dataRange();
  if (!used) return '';
  const lines: string[] = [];
  for (let r = 0; r <= used.r2; r++) {
    const vals: string[] = [];
    for (let c = 0; c <= used.c2; c++) vals.push(textAt(r, c));
    lines.push(Papa.unparse([vals], { delimiter, quotes: false, newline: '' }));
  }
  return lines.join('\r\n') + '\r\n';
}

// Post-processing of ExcelJS output for parts ExcelJS doesn't support (Phase 3: charts, constant names).
import type { Workbook } from '../model/workbook';
import { ValueReader, writeXlsx } from './xlsx';

export async function writeXlsxFull(wb: Workbook, getValue: ValueReader): Promise<ArrayBuffer> {
  return writeXlsx(wb, getValue);
}

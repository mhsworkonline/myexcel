import type { Sheet } from '../model/sheet';
import { Workbook } from '../model/workbook';

export async function readOds(_data: ArrayBuffer): Promise<Workbook> {
  throw new Error('OpenDocument import is not available yet.');
}

export async function writeOds(_wb: Workbook, _get: (s: Sheet, r: number, c: number) => unknown): Promise<Uint8Array> {
  throw new Error('OpenDocument export is not available yet.');
}

import { Workbook } from '../model/workbook';

export async function readXls(_data: ArrayBuffer): Promise<Workbook> {
  throw new Error('Excel 97-2003 import is not available yet.');
}

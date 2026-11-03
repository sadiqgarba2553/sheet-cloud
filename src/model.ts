export interface CellStyle { ff?: string; fs?: number; bl?: number; it?: number; ul?: { s: number }; cl?: { rgb: string }; bg?: { rgb: string }; ht?: number; vt?: number; tb?: number; n?: { pattern: string } }
export interface Cell { v?: string | number | boolean | null; f?: string; t?: number; s?: string | CellStyle; p?: any }
export interface Sheet { id: string; name: string; rowCount: number; columnCount: number; cellData: Record<number, Record<number, Cell>>; mergeData?: { startRow: number; endRow: number; startColumn: number; endColumn: number }[]; columnData?: Record<number, { w?: number; hd?: number }>; rowData?: Record<number, { h?: number; hd?: number }>; freeze?: { startRow: number; startColumn: number; xSplit: number; ySplit: number }; [key: string]: any }
export interface Workbook { id: string; name: string; appVersion: string; locale: string; sheetOrder: string[]; sheets: Record<string, Sheet>; styles: Record<string, CellStyle>; resources?: any[]; [key: string]: any }
export interface BookRecord { format: 'Sheet Cloud'; formatVersion: 1; id: string; title: string; createdAt: string; updatedAt: string; workbook: Workbook; versions: { id: string; name: string; at: string; workbook: Workbook }[] }

export function blankBook(title = 'Untitled workbook'): BookRecord {
  const id = crypto.randomUUID();
  const sheetId = crypto.randomUUID();
  const now = new Date().toISOString();
  return { format: 'Sheet Cloud', formatVersion: 1, id, title, createdAt: now, updatedAt: now, versions: [], workbook: { id, name: title, appVersion: '1.3.1', locale: 'enUS', styles: {}, sheetOrder: [sheetId], sheets: { [sheetId]: { id: sheetId, name: 'Sheet1', rowCount: 1000, columnCount: 26, cellData: {}, defaultColumnWidth: 110, defaultRowHeight: 26 } } } };
}

export function parseNative(text: string): BookRecord {
  if (text.length > 50 * 1024 * 1024) throw new Error('Workbook exceeds the 50 MB import limit.');
  const data = JSON.parse(text, (key, value) => { if (['__proto__', 'prototype', 'constructor'].includes(key)) throw new Error('Unsafe workbook property.'); return value; });
  if (data.format !== 'Sheet Cloud' || data.formatVersion !== 1) throw new Error('Unsupported Sheet Cloud file or version.');
  if (typeof data.id !== 'string' || typeof data.title !== 'string' || !Array.isArray(data.versions) || !Array.isArray(data.workbook?.sheetOrder) || !data.workbook.sheetOrder.length) throw new Error('Invalid workbook structure.');
  for (const id of data.workbook.sheetOrder) {
    const sheet = data.workbook.sheets?.[id];
    if (!sheet || typeof sheet.name !== 'string' || !Number.isInteger(sheet.rowCount) || !Number.isInteger(sheet.columnCount) || sheet.rowCount < 1 || sheet.rowCount > 1000000 || sheet.columnCount < 1 || sheet.columnCount > 16384 || typeof sheet.cellData !== 'object' || !sheet.cellData) throw new Error('Invalid worksheet.');
  }
  return data;
}

export function cellText(cell?: Cell): string {
  if (cell?.v != null) return String(cell.v);
  return cell?.p?.body?.dataStream?.replace(/\r?\n$/, '') || '';
}

export function usedExtent(sheet: Sheet) {
  let rows = 0, columns = 0;
  for (const [r, row] of Object.entries(sheet.cellData || {})) for (const [c, cell] of Object.entries(row || {})) if (cell && (cell.v != null || cell.f || cell.p || cell.s)) { rows = Math.max(rows, +r + 1); columns = Math.max(columns, +c + 1); }
  return { rows, columns };
}

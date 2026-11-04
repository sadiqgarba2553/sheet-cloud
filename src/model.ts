export interface CellStyle { ff?: string; fs?: number; bl?: number; it?: number; ul?: { s: number }; cl?: { rgb: string }; bg?: { rgb: string }; ht?: number; vt?: number; tb?: number; n?: { pattern: string } }
export interface Cell { v?: string | number | boolean | null; f?: string; t?: number; s?: string | CellStyle; p?: any }
export interface Sheet {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
  cellData: Record<number, Record<number, Cell>>;
  mergeData?: { startRow: number; endRow: number; startColumn: number; endColumn: number }[];
  columnData?: Record<number, { w?: number; hd?: number }>;
  rowData?: Record<number, { h?: number; hd?: number }>;
  freeze?: { startRow: number; startColumn: number; xSplit: number; ySplit: number };
  hidden?: number;
  charts?: SheetChart[];
  pivots?: SheetPivot[];
  [key: string]: any;
}

export type ChartType = 'bar' | 'column' | 'line' | 'pie' | 'scatter' | 'area';

export interface SheetChart {
  id: string;
  sheetId: string;
  title: string;
  chartType: ChartType;
  range: string;
  x: number;
  y: number;
  width: number;
  height: number;
  series?: { name: string; color?: string }[];
  showLegend?: boolean;
  showGrid?: boolean;
}

export interface SheetPivot {
  id: string;
  name: string;
  sourceSheetId: string;
  sourceRange: string;
  targetSheetId: string;
  targetCell: string;
  rowFields: string[];
  colFields: string[];
  valueFields: { field: string; agg: 'SUM' | 'COUNT' | 'AVERAGE' | 'MIN' | 'MAX' }[];
  filterFields?: { field: string; selectedValues?: string[] }[];
}

export interface Scenario {
  id: string;
  name: string;
  changingCells: { address: string; value: string | number }[];
  comment?: string;
}

export interface DataQuery {
  id: string;
  name: string;
  sourceType: 'range' | 'csv' | 'json';
  steps: { op: 'text-to-columns' | 'remove-duplicates' | 'filter' | 'sort'; params: any }[];
}

export interface Workbook {
  id: string;
  name: string;
  appVersion: string;
  locale: string;
  sheetOrder: string[];
  sheets: Record<string, Sheet>;
  styles: Record<string, CellStyle>;
  resources?: any[];
  charts?: SheetChart[];
  pivots?: SheetPivot[];
  scenarios?: Scenario[];
  dataQueries?: DataQuery[];
  [key: string]: any;
}

export interface BookRecord {
  format: 'Sheet Cloud';
  formatVersion: 1;
  id: string;
  title: string;
  createdAt: string;
  updatedAt: string;
  workbook: Workbook;
  versions: { id: string; name: string; at: string; workbook: Workbook }[];
}

export function blankBook(title = 'Untitled workbook'): BookRecord {
  const id = crypto.randomUUID();
  const sheetId = crypto.randomUUID();
  const now = new Date().toISOString();
  return {
    format: 'Sheet Cloud',
    formatVersion: 1,
    id,
    title,
    createdAt: now,
    updatedAt: now,
    versions: [],
    workbook: {
      id,
      name: title,
      appVersion: '1.3.1',
      locale: 'enUS',
      styles: {},
      sheetOrder: [sheetId],
      sheets: {
        [sheetId]: {
          id: sheetId,
          name: 'Sheet1',
          rowCount: 1000,
          columnCount: 26,
          cellData: {},
          defaultColumnWidth: 110,
          defaultRowHeight: 26,
          charts: [],
          pivots: [],
        },
      },
      charts: [],
      pivots: [],
      scenarios: [],
      dataQueries: [],
    },
  };
}

export function parseNative(text: string): BookRecord {
  if (text.length > 50 * 1024 * 1024) throw new Error('Workbook exceeds the 50 MB import limit.');
  const data = JSON.parse(text, (key, value) => {
    if (['__proto__', 'prototype', 'constructor'].includes(key)) throw new Error('Unsafe workbook property.');
    return value;
  });
  if (data.format !== 'Sheet Cloud' || data.formatVersion !== 1) throw new Error('Unsupported Sheet Cloud file or version.');
  if (typeof data.id !== 'string' || typeof data.title !== 'string' || !Array.isArray(data.versions) || !Array.isArray(data.workbook?.sheetOrder) || !data.workbook.sheetOrder.length) throw new Error('Invalid workbook structure.');
  for (const id of data.workbook.sheetOrder) {
    const sheet = data.workbook.sheets?.[id];
    if (!sheet || typeof sheet.name !== 'string' || !Number.isInteger(sheet.rowCount) || !Number.isInteger(sheet.columnCount) || sheet.rowCount < 1 || sheet.rowCount > 1000000 || sheet.columnCount < 1 || sheet.columnCount > 16384 || typeof sheet.cellData !== 'object' || !sheet.cellData) throw new Error('Invalid worksheet.');
    if (!sheet.charts) sheet.charts = [];
    if (!sheet.pivots) sheet.pivots = [];
  }
  if (!data.workbook.charts) data.workbook.charts = [];
  if (!data.workbook.pivots) data.workbook.pivots = [];
  if (!data.workbook.scenarios) data.workbook.scenarios = [];
  if (!data.workbook.dataQueries) data.workbook.dataQueries = [];
  return data;
}

export function cellText(cell?: Cell): string {
  if (cell?.v != null) return String(cell.v);
  return cell?.p?.body?.dataStream?.replace(/\r?\n$/, '') || '';
}

export function usedExtent(sheet: Sheet) {
  let rows = 0, columns = 0;
  for (const [r, row] of Object.entries(sheet.cellData || {})) {
    for (const [c, cell] of Object.entries(row || {})) {
      if (cell && (cell.v != null || cell.f || cell.p || cell.s)) {
        rows = Math.max(rows, +r + 1);
        columns = Math.max(columns, +c + 1);
      }
    }
  }
  return { rows, columns };
}

export function colToLetters(colIndex: number): string {
  let temp = colIndex;
  let letter = '';
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

export function lettersToCol(letters: string): number {
  let col = 0;
  const upper = letters.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    col = col * 26 + (upper.charCodeAt(i) - 64);
  }
  return col - 1;
}

export function rowColToA1(row: number, col: number): string {
  return `${colToLetters(col)}${row + 1}`;
}

export function parseA1(a1: string): { row: number; col: number } {
  const match = /^([A-Za-z]+)([1-9]\d*)$/.exec(a1.replace(/\$/g, '').trim());
  if (!match || +match[2] > 1000000 || lettersToCol(match[1]) > 16383) throw new Error('Enter a valid cell address (A1:XFD1000000).');
  return {
    row: parseInt(match[2], 10) - 1,
    col: lettersToCol(match[1]),
  };
}

export function parseRange(rangeStr: string): { startRow: number; startCol: number; endRow: number; endCol: number } {
  const cleaned = rangeStr.replace(/\$/g, '').trim();
  const parts = cleaned.split(':');
  if (parts.length > 2) throw new Error('Enter a single rectangular range.');
  if (parts.length === 1) {
    const pt = parseA1(parts[0]);
    return { startRow: pt.row, startCol: pt.col, endRow: pt.row, endCol: pt.col };
  }
  const p1 = parseA1(parts[0]);
  const p2 = parseA1(parts[1]);
  if ((Math.abs(p2.row - p1.row) + 1) * (Math.abs(p2.col - p1.col) + 1) > 200000) throw new Error('Choose a range of at most 200,000 cells.');
  return {
    startRow: Math.min(p1.row, p2.row),
    startCol: Math.min(p1.col, p2.col),
    endRow: Math.max(p1.row, p2.row),
    endCol: Math.max(p1.col, p2.col),
  };
}

import type ExcelJS from 'exceljs';
import { blankBook, BookRecord, Cell, CellStyle, cellText, Sheet, usedExtent } from './model';

const rgb = (color?: Partial<ExcelJS.Color>) => color?.argb ? `#${color.argb.slice(-6)}` : undefined;
const argb = (color?: string) => color && /^#[0-9a-f]{6}$/i.test(color) ? `FF${color.slice(1)}` : undefined;
const horizontal = ['', 'left', 'center', 'right'] as const;
const vertical = ['', 'top', 'middle', 'bottom'] as const;

export async function importXlsx(bytes: Uint8Array, filename: string): Promise<BookRecord> {
  if (bytes.byteLength > 25 * 1024 * 1024) throw new Error('Choose an XLSX file smaller than 25 MB.');
  const { default: Excel } = await import('exceljs');
  const source = new Excel.Workbook();
  await source.xlsx.load(bytes.slice().buffer as any);
  const book = blankBook(filename.replace(/\.xlsx$/i, ''));
  book.workbook.sheetOrder = []; book.workbook.sheets = {};
  let cells = 0;
  for (const worksheet of source.worksheets) {
    if (worksheet.rowCount > 100000 || worksheet.columnCount > 16384) throw new Error('This workbook exceeds the current import size limits.');
    const id = crypto.randomUUID();
    const sheet: Sheet = { id, name: worksheet.name, rowCount: Math.max(1000, worksheet.rowCount + 50), columnCount: Math.min(16384, Math.max(26, worksheet.columnCount + 5)), cellData: {}, columnData: {}, rowData: {}, mergeData: [], hidden: worksheet.state && worksheet.state !== 'visible' ? 1 : 0 };
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (row.height || row.hidden) sheet.rowData![rowNumber - 1] = { h: row.height ? row.height * 96 / 72 : undefined, hd: row.hidden ? 1 : 0 };
      row.eachCell({ includeEmpty: false }, (sourceCell, columnNumber) => {
        if (++cells > 200000) throw new Error('This workbook exceeds the 200,000 populated-cell import limit.');
        if (sourceCell.type === Excel.ValueType.Merge) return;
        const cell: Cell = {};
        const value = sourceCell.value;
        if (value == null) return;
        if (value instanceof Date) { cell.v = value.getTime() / 86400000 + 25569; cell.t = 2; }
        else if (typeof value === 'object' && ('formula' in value || 'sharedFormula' in value)) { cell.f = `=${sourceCell.formula}`; const result = sourceCell.result; cell.v = typeof result === 'object' && result && 'error' in result ? result.error : result instanceof Date ? result.getTime() / 86400000 + 25569 : result as any; }
        else if (typeof value === 'object' && 'richText' in value) { cell.v = value.richText.map(run => run.text).join(''); cell.t = 1; }
        else if (typeof value === 'object' && 'hyperlink' in value) { cell.v = value.text; cell.t = 1; }
        else if (typeof value === 'object' && 'error' in value) { cell.v = value.error; cell.t = 5; }
        else { cell.v = value as string | number | boolean; cell.t = typeof value === 'number' ? 2 : typeof value === 'boolean' ? 3 : 1; }
        const style: CellStyle = {};
        const font = sourceCell.font;
        if (font?.name) style.ff = font.name;
        if (font?.size) style.fs = font.size;
        if (font?.bold) style.bl = 1;
        if (font?.italic) style.it = 1;
        if (font?.underline) style.ul = { s: 1 };
        if (rgb(font?.color)) style.cl = { rgb: rgb(font.color)! };
        if (sourceCell.fill?.type === 'pattern' && rgb(sourceCell.fill.fgColor)) style.bg = { rgb: rgb(sourceCell.fill.fgColor)! };
        if (sourceCell.numFmt && sourceCell.numFmt !== 'General') style.n = { pattern: sourceCell.numFmt };
        if (value instanceof Date && !style.n) style.n = { pattern: 'yyyy-mm-dd' };
        const alignment = sourceCell.alignment;
        if (alignment?.horizontal && ['left', 'center', 'right'].includes(alignment.horizontal)) style.ht = horizontal.indexOf(alignment.horizontal as any);
        if (alignment?.vertical && ['top', 'middle', 'bottom'].includes(alignment.vertical)) style.vt = vertical.indexOf(alignment.vertical as any);
        if (alignment?.wrapText) style.tb = 3;
        if (Object.keys(style).length) cell.s = style;
        (sheet.cellData[rowNumber - 1] ||= {})[columnNumber - 1] = cell;
      });
    });
    (worksheet.columns || []).forEach((column, i) => { if (column.width || column.hidden) sheet.columnData![i] = { w: column.width ? column.width * 7 + 5 : undefined, hd: column.hidden ? 1 : 0 }; });
    for (const range of worksheet.model.merges || []) {
      const match = /([A-Z]+)(\d+):([A-Z]+)(\d+)/.exec(range);
      if (match) sheet.mergeData!.push({ startRow: +match[2] - 1, endRow: +match[4] - 1, startColumn: columnIndex(match[1]), endColumn: columnIndex(match[3]) });
    }
    const view = worksheet.views?.find(view => view.state === 'frozen');
    if (view?.state === 'frozen') sheet.freeze = { startRow: view.ySplit || -1, startColumn: view.xSplit || -1, xSplit: view.xSplit || 0, ySplit: view.ySplit || 0 };
    book.workbook.sheetOrder.push(id); book.workbook.sheets[id] = sheet;
  }
  if (!book.workbook.sheetOrder.length) throw new Error('This file contains no worksheets.');
  return book;
}

export async function exportXlsx(book: BookRecord): Promise<Uint8Array> {
  const { default: Excel } = await import('exceljs');
  const output = new Excel.Workbook();
  output.creator = 'BlueCloud AI Technologies'; output.title = book.title; output.calcProperties.fullCalcOnLoad = true;
  for (const id of book.workbook.sheetOrder) {
    const source = book.workbook.sheets[id];
    const worksheet = output.addWorksheet(source.name, { state: source.hidden ? 'hidden' : 'visible' });
    for (const [r, row] of Object.entries(source.cellData || {})) for (const [c, data] of Object.entries(row || {})) {
      if (!data) continue;
      const cell = worksheet.getCell(+r + 1, +c + 1);
      if (data.f) cell.value = { formula: data.f.replace(/^=/, ''), result: data.v ?? undefined } as any;
      else if (data.t === 5) cell.value = { error: String(data.v) } as any;
      else cell.value = data.p ? cellText(data) : data.v ?? null;
      const style = (typeof data.s === 'string' ? book.workbook.styles[data.s] : data.s) || {};
      cell.font = { name: style.ff || 'Arial', size: style.fs || 11, bold: style.bl === 1, italic: style.it === 1, underline: style.ul?.s === 1, ...(argb(style.cl?.rgb) ? { color: { argb: argb(style.cl?.rgb) } } : {}) };
      if (argb(style.bg?.rgb)) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(style.bg!.rgb) } };
      if (style.n?.pattern) cell.numFmt = style.n.pattern;
      cell.alignment = { ...(style.ht && horizontal[style.ht] ? { horizontal: horizontal[style.ht] as any } : {}), ...(style.vt && vertical[style.vt] ? { vertical: vertical[style.vt] as any } : {}), wrapText: style.tb === 3 };
    }
    for (const [i, data] of Object.entries(source.columnData || {})) { const column = worksheet.getColumn(+i + 1); if (data.w) column.width = Math.max(1, (data.w - 5) / 7); column.hidden = data.hd === 1; }
    for (const [i, data] of Object.entries(source.rowData || {})) { const row = worksheet.getRow(+i + 1); if (data.h) row.height = data.h * 72 / 96; row.hidden = data.hd === 1; }
    for (const merge of source.mergeData || []) worksheet.mergeCells(merge.startRow + 1, merge.startColumn + 1, merge.endRow + 1, merge.endColumn + 1);
    if (source.freeze?.xSplit || source.freeze?.ySplit) worksheet.views = [{ state: 'frozen', xSplit: source.freeze.xSplit, ySplit: source.freeze.ySplit }];
  }
  return new Uint8Array(await output.xlsx.writeBuffer());
}

export function columnIndex(label: string) { return [...label.toUpperCase()].reduce((n, char) => n * 26 + char.charCodeAt(0) - 64, 0) - 1; }

export function parseCsv(text: string): string[][] {
  const rows: string[][] = []; let row: string[] = [], field = '', quoted = false;
  text = text.replace(/^\uFEFF/, '');
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    if (char === '"') { if (quoted && text[i + 1] === '"') { field += '"'; i++; } else if (quoted || field === '') quoted = !quoted; else field += char; }
    else if (char === ',' && !quoted) { row.push(field); field = ''; }
    else if ((char === '\n' || char === '\r') && !quoted) { if (char === '\r' && text[i + 1] === '\n') i++; row.push(field); rows.push(row); row = []; field = ''; }
    else field += char;
  }
  if (quoted) throw new Error('CSV contains an unclosed quoted field.');
  if (field || row.length) { row.push(field); rows.push(row); }
  return rows;
}

export function importCsv(text: string, filename: string): BookRecord {
  const rows = parseCsv(text); if (rows.reduce((n, row) => n + row.length, 0) > 200000) throw new Error('CSV exceeds the 200,000-cell import limit.');
  const book = blankBook(filename.replace(/\.csv$/i, '')); const sheet = book.workbook.sheets[book.workbook.sheetOrder[0]];
  sheet.rowCount = Math.max(1000, rows.length + 50); sheet.columnCount = rows.reduce((maximum, row) => Math.max(maximum, row.length), 26);
  rows.forEach((row, r) => row.forEach((value, c) => { if (value === '') return; const numeric = /^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(value) && Number.isFinite(Number(value)) && value.replace(/\D/g, '').length < 16; (sheet.cellData[r] ||= {})[c] = numeric ? { v: Number(value), t: 2 } : { v: value, t: 1 }; }));
  return book;
}

export function exportCsv(sheet: Sheet): string {
  const { rows, columns } = usedExtent(sheet); if (rows * columns > 1000000) throw new Error('The used area is too large for CSV export.');
  return Array.from({ length: rows }, (_, r) => Array.from({ length: columns }, (_, c) => { const cell = sheet.cellData[r]?.[c]; let value = cellText(cell); if (typeof cell?.v === 'string' && /^[=+\-@\t\r]/.test(value)) value = `'${value}`; return /[",\r\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value; }).join(',')).join('\r\n');
}

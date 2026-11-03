import { describe, it, expect } from 'vitest';
import 'fake-indexeddb/auto';
import { blankBook, parseNative } from '../src/model';
import { fromTemplate } from '../src/templates';
import { importCsv, parseCsv, exportCsv, exportXlsx, importXlsx } from '../src/exchange';
import { store } from '../src/store';

describe('workbook integrity', () => {
  it('creates independent workbooks', () => { const a = blankBook(), b = blankBook(); expect(a.id).not.toBe(b.id); expect(a.workbook.sheetOrder[0]).not.toBe(b.workbook.sheetOrder[0]); });
  it('round trips a native workbook with formulas, styles and versions', () => { const a = fromTemplate('budget'); a.versions.push({ id: 'first', name: 'First', at: '2026-11-02', workbook: structuredClone(a.workbook) }); expect(parseNative(JSON.stringify(a))).toEqual(a); });
  it('rejects unknown versions', () => { const a = blankBook(); (a as any).formatVersion = 99; expect(() => parseNative(JSON.stringify(a))).toThrow('Unsupported'); });
  it('rejects unsafe object properties', () => { expect(() => parseNative('{"__proto__":{}}')).toThrow('Unsafe'); });
  it('rejects invalid worksheet dimensions', () => { const a = blankBook(); a.workbook.sheets[a.workbook.sheetOrder[0]].rowCount = -1; expect(() => parseNative(JSON.stringify(a))).toThrow('Invalid worksheet'); });
  it('persists a complete workbook and retains original bytes separately', async () => { const a = fromTemplate('sales'); await store.save(a); expect(await store.get(a.id)).toEqual(a); await store.original(a.id, 'test.xlsx', new Uint8Array([1, 2, 3])); expect(Array.from((await store.getOriginal(a.id))!.data)).toEqual([1, 2, 3]); await store.remove(a.id); expect(await store.get(a.id)).toBeUndefined(); expect(await store.getOriginal(a.id)).toBeUndefined(); });
});

describe('CSV', () => {
  it('parses quoted commas, line breaks and escaped quotes', () => { expect(parseCsv('name,note\r\n"Smith, Alex","A ""quote""\nand a line"\r\n')).toEqual([['name', 'note'], ['Smith, Alex', 'A "quote"\nand a line']]); });
  it('rejects unclosed quotes', () => { expect(() => parseCsv('a,"unfinished')).toThrow('unclosed'); });
  it('preserves leading zeros, large identifiers and formula-like text', () => { const a = importCsv('00123,12345678901234567890,=SUM(A1:A2),-20', 'data.csv'); const cells = a.workbook.sheets[a.workbook.sheetOrder[0]].cellData[0]; expect(cells[0].v).toBe('00123'); expect(cells[1].v).toBe('12345678901234567890'); expect(cells[2]).toEqual({ v: '=SUM(A1:A2)', t: 1 }); expect(cells[3].v).toBe(-20); });
  it('mitigates formula injection when exporting literal text', () => { const a = importCsv('=1+1,@SUM(A1),-2', 'data.csv'); expect(exportCsv(a.workbook.sheets[a.workbook.sheetOrder[0]])).toBe("'=1+1,'@SUM(A1),-2"); });
  it('round trips commas and newlines', () => { const a = importCsv('"a,b","line\nbreak"', 'data.csv'); expect(parseCsv(exportCsv(a.workbook.sheets[a.workbook.sheetOrder[0]]))).toEqual([['a,b', 'line\nbreak']]); });
});

describe('XLSX compatibility subset', () => {
  it('round trips formulas, merged headers, basic styles and sheet names', async () => {
    const book = fromTemplate('budget');
    const source = book.workbook.sheets[book.workbook.sheetOrder[0]];
    source.freeze = { startRow: 4, startColumn: -1, xSplit: 0, ySplit: 4 };
    const bytes = await exportXlsx(book);
    expect(bytes[0]).toBe(80); expect(bytes[1]).toBe(75);
    const imported = await importXlsx(bytes, 'Budget.xlsx');
    const sheet = imported.workbook.sheets[imported.workbook.sheetOrder[0]];
    expect(sheet.hidden).toBe(0); expect(sheet.name).toBe('Monthly budget'); expect(sheet.cellData[4][3].f).toBe('=B5-C5'); expect(sheet.cellData[4][1].v).toBe(1200); expect(sheet.mergeData).toEqual(source.mergeData); expect(sheet.cellData[0][0].s).toMatchObject({ bl: 1, fs: 23 }); expect(sheet.freeze?.ySplit).toBe(4);
  });
  it('exports and imports multiple worksheets', async () => { const a = fromTemplate('sales'), b = blankBook(); const id = b.workbook.sheetOrder[0]; a.workbook.sheetOrder.push(id); a.workbook.sheets[id] = b.workbook.sheets[id]; expect((await importXlsx(await exportXlsx(a), 'multiple.xlsx')).workbook.sheetOrder).toHaveLength(2); });
});

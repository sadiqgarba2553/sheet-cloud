import { describe, it, expect } from 'vitest';
import 'fake-indexeddb/auto';
import { blankBook, parseNative, SheetChart, SheetPivot, Sheet } from '../src/model';
import { fromTemplate } from '../src/templates';
import { importCsv, parseCsv, exportCsv, exportXlsx, importXlsx, importJson, exportJson } from '../src/exchange';
import { computePivotTable, writePivotToSheet, runGoalSeek, textToColumns, removeDuplicates } from '../src/analysis';
import { store } from '../src/store';

describe('workbook integrity', () => {
  it('creates independent workbooks', () => {
    const a = blankBook(), b = blankBook();
    expect(a.id).not.toBe(b.id);
    expect(a.workbook.sheetOrder[0]).not.toBe(b.workbook.sheetOrder[0]);
  });
  it('round trips a native workbook with formulas, styles and versions', () => {
    const a = fromTemplate('budget');
    a.versions.push({ id: 'first', name: 'First', at: '2026-11-02', workbook: structuredClone(a.workbook) });
    expect(parseNative(JSON.stringify(a))).toEqual(a);
  });
  it('rejects unknown versions', () => {
    const a = blankBook();
    (a as any).formatVersion = 99;
    expect(() => parseNative(JSON.stringify(a))).toThrow('Unsupported');
  });
  it('rejects unsafe object properties', () => {
    expect(() => parseNative('{"__proto__":{}}')).toThrow('Unsafe');
  });
  it('rejects invalid worksheet dimensions', () => {
    const a = blankBook();
    a.workbook.sheets[a.workbook.sheetOrder[0]].rowCount = -1;
    expect(() => parseNative(JSON.stringify(a))).toThrow('Invalid worksheet');
  });
  it('persists a complete workbook and retains original bytes separately', async () => {
    const a = fromTemplate('sales');
    await store.save(a);
    expect(await store.get(a.id)).toEqual(a);
    await store.original(a.id, 'test.xlsx', new Uint8Array([1, 2, 3]));
    expect(Array.from((await store.getOriginal(a.id))!.data)).toEqual([1, 2, 3]);
    await store.remove(a.id);
    expect(await store.get(a.id)).toBeUndefined();
    expect(await store.getOriginal(a.id)).toBeUndefined();
  });
  it('preserves embedded charts and pivots in native serialization', () => {
    const a = blankBook();
    const sheetId = a.workbook.sheetOrder[0];
    const chart: SheetChart = {
      id: 'c1',
      sheetId,
      title: 'Q1 Revenue',
      chartType: 'column',
      range: 'A1:B5',
      x: 10,
      y: 10,
      width: 400,
      height: 250,
      showLegend: true,
      showGrid: true,
    };
    a.workbook.charts = [chart];
    const json = JSON.stringify(a);
    const parsed = parseNative(json);
    expect(parsed.workbook.charts).toHaveLength(1);
    expect(parsed.workbook.charts![0].title).toBe('Q1 Revenue');
    expect(parsed.workbook.charts![0].chartType).toBe('column');
  });
});

describe('CSV & JSON Exchange', () => {
  it('parses quoted commas, line breaks and escaped quotes', () => {
    expect(parseCsv('name,note\r\n"Smith, Alex","A ""quote""\nand a line"\r\n')).toEqual([
      ['name', 'note'],
      ['Smith, Alex', 'A "quote"\nand a line'],
    ]);
  });
  it('rejects unclosed quotes', () => {
    expect(() => parseCsv('a,"unfinished')).toThrow('unclosed');
  });
  it('preserves leading zeros, large identifiers and formula-like text', () => {
    const a = importCsv('00123,12345678901234567890,=SUM(A1:A2),-20', 'data.csv');
    const cells = a.workbook.sheets[a.workbook.sheetOrder[0]].cellData[0];
    expect(cells[0].v).toBe('00123');
    expect(cells[1].v).toBe('12345678901234567890');
    expect(cells[2]).toEqual({ v: '=SUM(A1:A2)', t: 1 });
    expect(cells[3].v).toBe(-20);
  });
  it('mitigates formula injection when exporting literal text', () => {
    const a = importCsv('=1+1,@SUM(A1),-2', 'data.csv');
    expect(exportCsv(a.workbook.sheets[a.workbook.sheetOrder[0]])).toBe("'=1+1,'@SUM(A1),-2");
  });
  it('round trips commas and newlines', () => {
    const a = importCsv('"a,b","line\nbreak"', 'data.csv');
    expect(parseCsv(exportCsv(a.workbook.sheets[a.workbook.sheetOrder[0]]))).toEqual([['a,b', 'line\nbreak']]);
  });
  it('imports and exports JSON datasets', () => {
    const rawJson = JSON.stringify([
      { Department: 'Engineering', Headcount: 42, Budget: 500000 },
      { Department: 'Marketing', Headcount: 18, Budget: 250000 },
    ]);
    const book = importJson(rawJson, 'departments.json');
    const sheet = book.workbook.sheets[book.workbook.sheetOrder[0]];
    expect(sheet.cellData[0][0].v).toBe('Department');
    expect(sheet.cellData[1][0].v).toBe('Engineering');
    expect(sheet.cellData[1][1].v).toBe(42);

    const exported = JSON.parse(exportJson(sheet));
    expect(exported).toHaveLength(2);
    expect(exported[0].Department).toBe('Engineering');
    expect(exported[0].Headcount).toBe(42);
  });
});

describe('Analysis & Pivot Engine', () => {
  it('computes grouping and SUM / AVERAGE / COUNT aggregations', () => {
    const sheet: Sheet = {
      id: 's1',
      name: 'Sales',
      rowCount: 10,
      columnCount: 5,
      cellData: {
        0: { 0: { v: 'Region' }, 1: { v: 'Product' }, 2: { v: 'Sales' } },
        1: { 0: { v: 'East' }, 1: { v: 'Widget' }, 2: { v: 100 } },
        2: { 0: { v: 'West' }, 1: { v: 'Widget' }, 2: { v: 200 } },
        3: { 0: { v: 'East' }, 1: { v: 'Gadget' }, 2: { v: 300 } },
        4: { 0: { v: 'West' }, 1: { v: 'Gadget' }, 2: { v: 400 } },
      },
    };

    const pivotSpec: SheetPivot = {
      id: 'p1',
      name: 'Regional Sales',
      sourceSheetId: 's1',
      sourceRange: 'A1:C5',
      targetSheetId: 'pSheet',
      targetCell: 'A1',
      rowFields: ['Region'],
      colFields: [],
      valueFields: [{ field: 'Sales', agg: 'SUM' }],
    };

    const result = computePivotTable(sheet, pivotSpec);
    expect(result.headers).toEqual(['Region', 'SUM of Sales']);
    expect(result.rows).toHaveLength(2);
    // East: 100 + 300 = 400; West: 200 + 400 = 600
    const eastRow = result.rows.find(r => r[0] === 'East');
    const westRow = result.rows.find(r => r[0] === 'West');
    expect(eastRow![1]).toBe(400);
    expect(westRow![1]).toBe(600);
    expect(result.grandTotals[1]).toBe(1000);

    // Test writing to a sheet
    const targetSheet: Sheet = { id: 'pSheet', name: 'Pivot', rowCount: 20, columnCount: 10, cellData: {} };
    writePivotToSheet(targetSheet, result, 0, 0);
    expect(targetSheet.cellData[0][0].v).toBe('Region');
    expect(targetSheet.cellData[1][0].v).toBe('East');
  });

  it('converges in Goal Seek using secant method', () => {
    // Formula: y = 3 * x + 15. Target y = 75 => x should be 20.
    const formulaEvaluator = (x: number) => 3 * x + 15;
    const solution = runGoalSeek(formulaEvaluator, 75, 1);
    expect(solution.success).toBe(true);
    expect(Math.abs(solution.value - 20)).toBeLessThan(1e-4);
  });

  it('performs text-to-columns splitting by delimiter', () => {
    const sheet: Sheet = {
      id: 's1',
      name: 'Names',
      rowCount: 5,
      columnCount: 10,
      cellData: {
        0: { 0: { v: 'John,Doe,35' } },
        1: { 0: { v: 'Jane,Smith,28' } },
      },
    };

    const res = textToColumns(sheet, 0, ',');
    expect(res.rowsAffected).toBe(2);
    expect(res.colsAdded).toBe(2);
    expect(sheet.cellData[0][0].v).toBe('John');
    expect(sheet.cellData[0][1].v).toBe('Doe');
    expect(sheet.cellData[0][2].v).toBe(35);
  });

  it('removes duplicate rows based on key columns', () => {
    const sheet: Sheet = {
      id: 's1',
      name: 'Data',
      rowCount: 6,
      columnCount: 5,
      cellData: {
        0: { 0: { v: 'ID' }, 1: { v: 'Val' } },
        1: { 0: { v: 'A101' }, 1: { v: 'Alpha' } },
        2: { 0: { v: 'A102' }, 1: { v: 'Beta' } },
        3: { 0: { v: 'A101' }, 1: { v: 'Alpha' } }, // Duplicate of row 1
      },
    };

    const res = removeDuplicates(sheet, [0], 1, 3);
    expect(res.removedCount).toBe(1);
    expect(res.retainedCount).toBe(2);
    expect(sheet.cellData[3]).toBeUndefined();
  });
});

describe('XLSX compatibility subset', () => {
  it('round trips formulas, merged headers, basic styles and sheet names', async () => {
    const book = fromTemplate('budget');
    const source = book.workbook.sheets[book.workbook.sheetOrder[0]];
    source.freeze = { startRow: 4, startColumn: -1, xSplit: 0, ySplit: 4 };
    const bytes = await exportXlsx(book);
    expect(bytes[0]).toBe(80);
    expect(bytes[1]).toBe(75);
    const imported = await importXlsx(bytes, 'Budget.xlsx');
    const sheet = imported.workbook.sheets[imported.workbook.sheetOrder[0]];
    expect(sheet.hidden).toBe(0);
    expect(sheet.name).toBe('Monthly budget');
    expect(sheet.cellData[4][3].f).toBe('=B5-C5');
    expect(sheet.cellData[4][1].v).toBe(1200);
    expect(sheet.mergeData).toEqual(source.mergeData);
    expect(sheet.cellData[0][0].s).toMatchObject({ bl: 1, fs: 23 });
    expect(sheet.freeze?.ySplit).toBe(4);
  });
  it('exports and imports multiple worksheets', async () => {
    const a = fromTemplate('sales'), b = blankBook();
    const id = b.workbook.sheetOrder[0];
    a.workbook.sheetOrder.push(id);
    a.workbook.sheets[id] = b.workbook.sheets[id];
    expect((await importXlsx(await exportXlsx(a), 'multiple.xlsx')).workbook.sheetOrder).toHaveLength(2);
  });
});

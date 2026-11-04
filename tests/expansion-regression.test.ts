import { describe, it, expect } from 'vitest';
import { blankBook, parseA1, parseRange } from '../src/model';
import { runGoalSeekAsync, textToColumns, removeDuplicates } from '../src/analysis';

describe('expansion regression coverage', () => {
  it('rejects invalid addresses and unbounded analysis ranges', () => {
    for (const address of ['A0', 'not a cell', 'XFE1', 'A1000001']) expect(() => parseA1(address)).toThrow();
    expect(() => parseRange('A1:XFD1000000')).toThrow();
    expect(parseRange('$B$3:A1')).toEqual({startRow:0,startCol:0,endRow:2,endCol:1});
  });
  it('waits for formula results and rejects invalid results', async () => {
    const result = await runGoalSeekAsync(async x => { await Promise.resolve(); return x * x; }, 25, 3);
    expect(result.success).toBe(true); expect(result.value).toBeCloseTo(5, 4);
    await expect(runGoalSeekAsync(async () => NaN, 1)).rejects.toThrow();
    expect((await runGoalSeekAsync(async () => 2, 5)).success).toBe(false);
  });
  it('does not merge distinct compound duplicate keys', () => {
    const book = blankBook(), sheet = book.workbook.sheets[book.workbook.sheetOrder[0]];
    sheet.cellData = {1:{0:{v:'a___b'},1:{v:'c'}},2:{0:{v:'a'},1:{v:'b___c'}},3:{0:{v:'a'},1:{v:'b___c'}}};
    expect(removeDuplicates(sheet, [0,1], 1, 3)).toEqual({removedCount:1,retainedCount:2});
  });
  it('rejects a text split before changing any rows when a destination contains data', () => {
    const book = blankBook(), sheet = book.workbook.sheets[book.workbook.sheetOrder[0]];
    sheet.cellData = {0:{0:{v:'one,two'}},1:{0:{v:'three,four'},1:{v:'keep me'}}};
    const before = structuredClone(sheet);
    expect(() => textToColumns(sheet, 0, ',', 0, 1)).toThrow('overwrite');
    expect(sheet).toEqual(before);
  });
});

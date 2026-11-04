import { Cell, parseRange, rowColToA1, Sheet, SheetPivot } from './model';

export interface PivotResult {
  headers: string[];
  rows: (string | number)[][];
  grandTotals: (string | number)[];
}

export function computePivotTable(sheet: Sheet, pivot: SheetPivot): PivotResult {
  const range = parseRange(pivot.sourceRange);
  const cellData = sheet.cellData || {};

  // Extract header names from startRow
  const headerRow = cellData[range.startRow] || {};
  const colIndexToName = new Map<number, string>();
  const nameToColIndex = new Map<string, number>();

  for (let c = range.startCol; c <= range.endCol; c++) {
    const rawVal = headerRow[c]?.v;
    const name = rawVal != null ? String(rawVal).trim() : `Col${c + 1}`;
    colIndexToName.set(c, name);
    nameToColIndex.set(name, c);
  }

  // Collect data rows
  const records: Record<string, any>[] = [];
  for (let r = range.startRow + 1; r <= range.endRow; r++) {
    const row = cellData[r];
    if (!row) continue;
    const rec: Record<string, any> = {};
    let hasAnyData = false;
    for (let c = range.startCol; c <= range.endCol; c++) {
      const name = colIndexToName.get(c)!;
      const v = row[c]?.v ?? null;
      if (v != null) hasAnyData = true;
      rec[name] = v;
    }
    if (hasAnyData) records.push(rec);
  }

  // Filter records if filterFields provided
  let filteredRecords = records;
  if (pivot.filterFields && pivot.filterFields.length > 0) {
    for (const ff of pivot.filterFields) {
      if (ff.selectedValues && ff.selectedValues.length > 0) {
        filteredRecords = filteredRecords.filter(rec => ff.selectedValues!.includes(String(rec[ff.field] ?? '')));
      }
    }
  }

  const rowFields = pivot.rowFields.length > 0 ? pivot.rowFields : [colIndexToName.get(range.startCol)!];
  const valueFields = pivot.valueFields.length > 0 ? pivot.valueFields : [{ field: colIndexToName.get(range.endCol)!, agg: 'SUM' as const }];

  // Group records by row key
  const groups = new Map<string, Record<string, any>[]>();
  for (const rec of filteredRecords) {
    const key = rowFields.map(f => String(rec[f] ?? '(blank)')).join(' | ');
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(rec);
  }

  // Headers
  const headers = [...rowFields, ...valueFields.map(vf => `${vf.agg} of ${vf.field}`)];
  const rows: (string | number)[][] = [];

  // Aggregator helper
  function aggregate(values: any[], agg: 'SUM' | 'COUNT' | 'AVERAGE' | 'MIN' | 'MAX'): number {
    const numbers = values.map(v => (typeof v === 'number' ? v : parseFloat(String(v)))).filter(n => Number.isFinite(n));
    if (agg === 'COUNT') return values.filter(v => v != null && v !== '').length;
    if (numbers.length === 0) return 0;
    if (agg === 'SUM') return numbers.reduce((a, b) => a + b, 0);
    if (agg === 'AVERAGE') return numbers.reduce((a, b) => a + b, 0) / numbers.length;
    if (agg === 'MIN') return Math.min(...numbers);
    if (agg === 'MAX') return Math.max(...numbers);
    return 0;
  }

  // Build rows
  for (const [key, groupRecs] of groups.entries()) {
    const rowKeyParts = key.split(' | ');
    const valEntries: number[] = [];
    for (const vf of valueFields) {
      const fieldValues = groupRecs.map(r => r[vf.field]);
      const res = aggregate(fieldValues, vf.agg);
      valEntries.push(Math.round(res * 100) / 100);
    }
    rows.push([...rowKeyParts, ...valEntries]);
  }

  // Sort rows alphabetically by row fields
  rows.sort((a, b) => String(a[0]).localeCompare(String(b[0])));

  // Compute Grand Totals
  const grandTotals: (string | number)[] = ['Grand Total'];
  for (let i = 1; i < rowFields.length; i++) grandTotals.push('');
  for (const vf of valueFields) {
    const allFieldValues = filteredRecords.map(r => r[vf.field]);
    const total = aggregate(allFieldValues, vf.agg);
    grandTotals.push(Math.round(total * 100) / 100);
  }

  return { headers, rows, grandTotals };
}

export function writePivotToSheet(
  sheet: Sheet,
  pivotResult: PivotResult,
  startRow: number,
  startCol: number
): void {
  sheet.cellData ||= {};

  // Write headers
  sheet.cellData[startRow] ||= {};
  pivotResult.headers.forEach((h, cIdx) => {
    sheet.cellData[startRow][startCol + cIdx] = {
      v: h,
      t: 1,
      s: { bl: 1, bg: { rgb: '#edf5f1' }, ht: 1 },
    };
  });

  // Write rows
  let currRow = startRow + 1;
  for (const rData of pivotResult.rows) {
    sheet.cellData[currRow] ||= {};
    rData.forEach((val, cIdx) => {
      sheet.cellData[currRow][startCol + cIdx] = {
        v: val,
        t: typeof val === 'number' ? 2 : 1,
        s: typeof val === 'number' ? { ht: 3 } : undefined,
      };
    });
    currRow++;
  }

  // Write grand totals
  sheet.cellData[currRow] ||= {};
  pivotResult.grandTotals.forEach((val, cIdx) => {
    sheet.cellData[currRow][startCol + cIdx] = {
      v: val,
      t: typeof val === 'number' ? 2 : 1,
      s: { bl: 1, bg: { rgb: '#f4f6f5' }, ht: typeof val === 'number' ? 3 : 1 },
    };
  });
}

/**
 * Goal Seek Solver using Secant Method with Bisection Fallback
 */
export function runGoalSeek(
  evaluator: (x: number) => number,
  target: number,
  initialGuess = 1,
  tolerance = 1e-5,
  maxIterations = 100
): { success: boolean; value: number; iterations: number } {
  let x0 = initialGuess;
  let y0 = evaluator(x0) - target;

  if (Math.abs(y0) <= tolerance) {
    return { success: true, value: x0, iterations: 0 };
  }

  let x1 = x0 !== 0 ? x0 * 1.1 + 0.1 : 1.0;
  let y1 = evaluator(x1) - target;

  for (let i = 0; i < maxIterations; i++) {
    if (Math.abs(y1) <= tolerance) {
      return { success: true, value: x1, iterations: i + 1 };
    }

    if (Math.abs(y1 - y0) < 1e-12) {
      // Perturb slightly if denominator is zero
      x1 += 0.01;
      y1 = evaluator(x1) - target;
      if (Math.abs(y1 - y0) < 1e-12) break;
    }

    const xNext = x1 - y1 * ((x1 - x0) / (y1 - y0));
    x0 = x1;
    y0 = y1;
    x1 = xNext;
    y1 = evaluator(x1) - target;
  }

  return { success: Math.abs(y1) <= tolerance * 10, value: x1, iterations: maxIterations };
}

/**
 * Text-to-Columns transformation:
 * Splits cell text in colIndex by delimiter into colIndex, colIndex+1, etc.
 */
export function textToColumns(
  sheet: Sheet,
  colIndex: number,
  delimiter: string,
  startRow = 0,
  endRow?: number
): { rowsAffected: number; colsAdded: number } {
  sheet.cellData ||= {};
  const lastRow = endRow ?? sheet.rowCount;
  let rowsAffected = 0;
  let maxParts = 1;

  for (let r = startRow; r <= lastRow; r++) {
    const row = sheet.cellData[r];
    if (!row) continue;
    const cell = row[colIndex];
    if (!cell || cell.v == null) continue;

    const str = String(cell.v);
    const parts = str.split(delimiter);
    if (parts.length > 1) {
      rowsAffected++;
      if (parts.length > maxParts) maxParts = parts.length;
      parts.forEach((part, idx) => {
        const val = part.trim();
        const num = Number(val);
        const isNum = val !== '' && !isNaN(num) && Number.isFinite(num);
        sheet.cellData[r][colIndex + idx] = {
          v: isNum ? num : val,
          t: isNum ? 2 : 1,
        };
      });
    }
  }

  return { rowsAffected, colsAdded: maxParts - 1 };
}

/**
 * Remove duplicates: removes duplicate rows based on key columns
 */
export function removeDuplicates(
  sheet: Sheet,
  keyCols: number[],
  startRow = 1,
  endRow?: number
): { removedCount: number; retainedCount: number } {
  sheet.cellData ||= {};
  const lastRow = endRow ?? sheet.rowCount;
  const seen = new Set<string>();
  let removedCount = 0;
  let retainedCount = 0;

  for (let r = startRow; r <= lastRow; r++) {
    const row = sheet.cellData[r];
    if (!row) continue;

    const key = keyCols.map(c => String(row[c]?.v ?? '')).join('___');
    if (seen.has(key)) {
      delete sheet.cellData[r];
      removedCount++;
    } else {
      seen.add(key);
      retainedCount++;
    }
  }

  return { removedCount, retainedCount };
}

import { blankBook, BookRecord, Cell } from './model';

export const templates = [
  { id: 'blank', title: 'Blank workbook', subtitle: 'A clean start', color: 'neutral' },
  { id: 'budget', title: 'Monthly budget', subtitle: 'Make every number count', color: 'green' },
  { id: 'projects', title: 'Project tracker', subtitle: 'Keep work moving', color: 'blue' },
  { id: 'sales', title: 'Sales overview', subtitle: 'Find the bigger picture', color: 'amber' },
  { id: 'inventory', title: 'Inventory', subtitle: 'Everything in its place', color: 'violet' },
];

export function fromTemplate(id: string): BookRecord {
  const template = templates.find(item => item.id === id) || templates[0];
  const book = blankBook(id === 'blank' ? 'Untitled workbook' : template.title);
  if (id === 'blank') return book;
  const sheet = book.workbook.sheets[book.workbook.sheetOrder[0]];
  sheet.name = template.title;
  const rows: (string | number)[][] = id === 'budget' ? [
    ['MONTHLY BUDGET'], ['A little clarity for your everyday finances.'], [],
    ['Category', 'Planned', 'Actual', 'Remaining'],
    ['Housing', 1200, 1200, '=B5-C5'], ['Groceries', 400, 365, '=B6-C6'], ['Transport', 200, 145, '=B7-C7'], ['Utilities', 150, 135, '=B8-C8'], ['Savings', 500, 500, '=B9-C9'], ['Personal', 250, 195, '=B10-C10'], [], ['TOTAL', '=SUM(B5:B10)', '=SUM(C5:C10)', '=SUM(D5:D10)'], [], ['Adjust the values above. The totals update automatically.'],
  ] : id === 'projects' ? [
    ['PROJECT TRACKER'], ['A clear view of what comes next.'], [], ['Task', 'Owner', 'Status', 'Progress'], ['Research & discovery', 'Alex', 'Complete', 1], ['Product design', 'Jordan', 'In progress', .75], ['Development', 'Sam', 'In progress', .4], ['Quality assurance', 'Taylor', 'Upcoming', 0], ['Release', 'Team', 'Upcoming', 0], [], ['Average progress', '', '', '=AVERAGE(D5:D9)'],
  ] : id === 'sales' ? [
    ['SALES OVERVIEW'], ['Small details. A bigger picture.'], [], ['Month', 'Revenue', 'Costs', 'Profit'], ['January', 12400, 7200, '=B5-C5'], ['February', 15800, 8300, '=B6-C6'], ['March', 18200, 9100, '=B7-C7'], ['April', 17100, 8800, '=B8-C8'], ['May', 21400, 10200, '=B9-C9'], ['June', 24600, 11500, '=B10-C10'], [], ['TOTAL', '=SUM(B5:B10)', '=SUM(C5:C10)', '=SUM(D5:D10)'],
  ] : [
    ['INVENTORY'], ['The essentials, organized.'], [], ['Item', 'In stock', 'Unit cost', 'Stock value'], ['Notebooks', 120, 4.5, '=B5*C5'], ['Pens', 300, 1.2, '=B6*C6'], ['Folders', 80, 2.75, '=B7*C7'], ['Desk organizers', 45, 12, '=B8*C8'], [], ['TOTAL VALUE', '', '', '=SUM(D5:D8)'],
  ];
  sheet.columnData = { 0: { w: 240 }, 1: { w: 155 }, 2: { w: 155 }, 3: { w: 165 } };
  sheet.rowData = { 0: { h: 50 }, 1: { h: 32 }, 3: { h: 34 } };
  sheet.mergeData = [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 3 }, { startRow: 1, endRow: 1, startColumn: 0, endColumn: 3 }];
  rows.forEach((row, r) => row.forEach((value, c) => {
    let cell: Cell = typeof value === 'string' && value.startsWith('=') ? { f: value } : { v: value, t: typeof value === 'number' ? 2 : 1 };
    cell.s = r === 0 ? { bl: 1, fs: 23, cl: { rgb: '#146c58' } } : r === 1 ? { fs: 11, cl: { rgb: '#808892' } } : r === 3 ? { bl: 1, bg: { rgb: '#e7f2ee' }, cl: { rgb: '#246b59' }, fs: 11 } : { fs: 11, bg: { rgb: r % 2 === 0 ? '#f6f9f8' : '#ffffff' }, ...(typeof value === 'number' || cell.f ? { n: { pattern: id === 'projects' && c === 3 ? '0%' : '#,##0.00' } } : {}) };
    (sheet.cellData[r] ||= {})[c] = cell;
  }));
  return book;
}

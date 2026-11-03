import { test, expect, Page } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import Excel from 'exceljs';

async function go(page: Page, address: string) { const box = page.getByRole('textbox', { name: 'Cell address', exact: true }); await box.fill(address); await box.press('Enter'); }
async function set(page: Page, address: string, value: string) { await go(page, address); const box = page.getByRole('textbox', { name: 'Cell value or formula', exact: true }); await box.fill(value); await box.press('Enter'); }
async function ready(page: Page) { await expect(page.getByLabel('Spreadsheet workspace')).toBeVisible({ timeout: 30000 }); await expect(page.getByText('Preparing your workbook…')).toHaveCount(0, { timeout: 45000 }); await expect(page.locator('canvas:visible').first()).toBeVisible({ timeout: 15000 }); }

test('create, calculate SUM and IF, autosave and reopen', async ({ page }) => {
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await page.goto('/'); await page.screenshot({ path: 'output/playwright/home.png' });
  await page.getByRole('button', { name: 'New workbook', exact: true }).click(); await ready(page);
  await page.getByRole('textbox', { name: 'Workbook title' }).fill('Verified workbook');
  await set(page, 'A1', '10'); await set(page, 'A2', '20'); await set(page, 'A3', '=SUM(A1:A2)');
  await expect(page.getByLabel('Calculated cell value')).toHaveText('30');
  await set(page, 'B1', '=IF(A3>20,"High","Low")'); await expect(page.getByLabel('Calculated cell value')).toHaveText('High');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Back to workbooks' }).click();
  await page.reload(); await page.getByRole('button', { name: 'Verified workbook Local workbook' }).click(); await ready(page);
  await go(page, 'A3'); await expect(page.getByRole('textbox', { name: 'Cell value or formula', exact: true })).toHaveValue('=SUM(A1:A2)'); await expect(page.getByLabel('Calculated cell value')).toHaveText('30');
  expect(errors).toEqual([]);
});

test('template has working formulas and native export reopens', async ({ page }) => {
  await page.goto('/'); await page.getByRole('button', { name: 'Monthly budget Make every number count' }).click(); await ready(page);
  await go(page, 'D6'); await expect(page.getByLabel('Calculated cell value')).toHaveText('35.00');
  await set(page, 'C6', '300'); await go(page, 'D6'); await expect(page.getByLabel('Calculated cell value')).toHaveText('100.00');
  await page.screenshot({ path: 'output/playwright/workbook.png' });
  await page.getByRole('button', { name: 'File', exact: true }).click(); const downloading = page.waitForEvent('download'); await page.getByRole('button', { name: 'Sheet Cloud workbook Full supported workbook' }).click(); const download = await downloading; await download.saveAs('output/playwright/Budget.scloud');
  await page.getByRole('button', { name: 'Back to workbooks' }).click(); await page.locator('input[type=file]').setInputFiles('output/playwright/Budget.scloud'); await ready(page); await go(page, 'D6'); await expect(page.getByLabel('Calculated cell value')).toHaveText('100.00');
});

test('XLSX imports and exports real values and formulas', async ({ page }) => {
  const source = new Excel.Workbook(); const sheet = source.addWorksheet('Numbers'); sheet.getCell('A1').value = 7; sheet.getCell('A2').value = 8; sheet.getCell('A3').value = { formula: 'SUM(A1:A2)', result: 15 }; sheet.getCell('A1').font = { bold: true };
  await page.goto('/'); await page.locator('input[type=file]').setInputFiles({ name: 'Numbers.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(await source.xlsx.writeBuffer()) }); await ready(page); await go(page, 'A3'); await expect(page.getByLabel('Calculated cell value')).toHaveText('15');
  await page.getByRole('button', { name: 'File', exact: true }).click(); page.once('dialog', dialog => dialog.accept()); const downloading = page.waitForEvent('download'); await page.getByRole('button', { name: 'Excel workbook Basic cells' }).click(); const download = await downloading; await download.saveAs('output/playwright/Numbers.xlsx'); const result = new Excel.Workbook(); await result.xlsx.load(await readFile('output/playwright/Numbers.xlsx') as any); expect(result.worksheets[0].getCell('A3').formula).toBe('SUM(A1:A2)'); expect(result.worksheets[0].getCell('A1').value).toBe(7); expect(result.worksheets[0].getCell('A1').font.bold).toBe(true);
});

test('CSV formula-like strings stay literal and editing is local', async ({ page }) => {
  const external: string[] = []; page.on('request', request => { if (/^https?:/.test(request.url()) && !request.url().startsWith('http://127.0.0.1:1430')) external.push(request.url()); });
  await page.goto('/'); await page.locator('input[type=file]').setInputFiles({ name: 'data.csv', mimeType: 'text/csv', buffer: Buffer.from('Name,Amount\nAlice,50\nBob,75\n=2+2,100') }); await ready(page); await go(page, 'A4'); await expect(page.getByLabel('Calculated cell value')).toHaveText('=2+2'); await set(page, 'B5', '=SUM(B2:B4)'); await expect(page.getByLabel('Calculated cell value')).toHaveText('225'); expect(external).toEqual([]);
});

test('named history restores the workbook and keeps a backup', async ({ page }) => {
  await page.goto('/'); await page.getByRole('button', { name: 'New workbook', exact: true }).click(); await ready(page); await set(page, 'A1', 'Original'); await page.getByRole('button', { name: 'History', exact: true }).click(); await page.getByRole('textbox', { name: 'Version name' }).fill('First draft'); await page.getByRole('button', { name: 'Save version', exact: true }).click(); await page.getByRole('button', { name: 'Close dialog' }).click(); await set(page, 'A1', 'Later'); await page.getByRole('button', { name: 'History', exact: true }).click(); page.once('dialog', dialog => dialog.accept()); await page.getByRole('button', { name: 'Restore', exact: true }).click(); await ready(page); await go(page, 'A1'); await expect(page.getByLabel('Calculated cell value')).toHaveText('Original'); await page.getByRole('button', { name: 'History', exact: true }).click(); await expect(page.getByText('Before restore', { exact: true })).toBeVisible();
});

test('type directly into the grid, undo, and keep history separate between workbooks', async ({ page }) => {
  await page.goto('/'); await page.getByRole('button', { name: 'New workbook', exact: true }).click(); await ready(page);
  const canvas = page.locator('canvas:visible').first();
  await canvas.click({ position: { x: 85, y: 38 } });
  await page.keyboard.type('Direct grid entry'); await page.keyboard.press('Enter');
  await go(page, 'A1'); await expect(page.getByLabel('Calculated cell value')).toHaveText('Direct grid entry');
  await page.getByRole('button', { name: 'Undo', exact: true }).first().click();
  await go(page, 'A1'); await expect(page.getByLabel('Calculated cell value')).toHaveText('');
  await set(page, 'A1', 'First workbook'); await page.getByRole('button', { name: 'Back to workbooks' }).click(); await page.getByRole('button', { name: 'New workbook', exact: true }).click(); await ready(page);
  await page.getByRole('button', { name: 'Undo', exact: true }).first().click(); await go(page, 'A1'); await expect(page.getByLabel('Calculated cell value')).toHaveText('');
});

test('chart preview reflects selected numbers and dark appearance persists', async ({ page }) => {
  await page.goto('/'); await page.getByRole('button', { name: 'Settings & appearance' }).click(); await page.getByRole('button', { name: 'Dark', exact: true }).click(); await page.getByRole('button', { name: 'Close dialog' }).click(); await page.reload(); await expect(page.locator('html')).toHaveClass('dark');
  await page.getByRole('button', { name: 'Sales overview Find the bigger picture' }).click(); await ready(page); await go(page, 'A4:D10'); await page.getByRole('button', { name: 'Chart preview', exact: true }).click(); await expect(page.getByRole('img', { name: 'Selected data bar chart' })).toBeVisible(); await expect(page.locator('svg[aria-label="Selected data bar chart"] rect')).toHaveCount(6); await page.screenshot({ path: 'output/playwright/chart-dark.png' });
});

test('print exports the active sheet values to a real PDF layout', async ({ page, context }) => {
  await page.goto('/'); await page.getByRole('button', { name: 'Monthly budget Make every number count' }).click(); await ready(page);
  await go(page, 'D12'); await expect(page.getByLabel('Calculated cell value')).toHaveText('160.00');
  await page.evaluate(() => { Window.prototype.print = function () {}; });
  await page.getByRole('button', { name: 'Print active sheet' }).click();
  const frame = page.frameLocator('iframe[title="Sheet Cloud print layout"]'); await expect(frame.locator('table')).toContainText('Groceries'); await expect(frame.locator('table')).toContainText('160.00');
  const html = await page.locator('iframe[title="Sheet Cloud print layout"]').evaluate((element: HTMLIFrameElement) => element.contentDocument!.documentElement.outerHTML);
  const printPage = await context.newPage(); await printPage.setContent(html); await printPage.pdf({ path: 'output/playwright/budget.pdf', preferCSSPageSize: true, printBackground: true }); await printPage.close();
});

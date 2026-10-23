import { expect, test, type Page } from '@playwright/test';
import { gotoApp } from './helpers';

async function run<T = unknown>(page: Page, body: string): Promise<T> {
  return page.evaluate(`(async () => { const api = window.grillon.api; const wb = api.getActiveWorkbook(); const ws = wb.getActiveSheet(); ${body} })()`) as Promise<T>;
}
const hiddenRows = (page: Page) => run<number[]>(page, `const s = ws.getSheet(); return Array.from({ length: 8 }, (_, r) => r).filter((r) => !s.getRowVisible(r));`);

// 45292 = 01/01/2024
test.beforeEach(async ({ page }) => {
  await gotoApp(page);
  await run(page, `ws.getRange('A1:C7').setValues([
    ['Région', 'Date', 'Montant'],
    ['Nord', 45292, 10],
    ['Sud', 45323, 20],
    ['Nord', 45352, 30],
    ['Est', 45383, 40],
    ['Sud', 45413, 50],
    ['Nord', 45444, 60],
  ]);`);
});

async function insert(page: Page, type: 'slicer' | 'timeline', field: string) {
  await run(page, `wb.setActiveRange(ws.getRange('A2')); api.executeCommand('grillon.insert.${type}');`);
  const dialog = page.getByTestId('dialog-slicer');
  await dialog.getByTestId('slicer-field').selectOption(field);
  await dialog.getByRole('button', { name: 'OK' }).click();
}

test('Segment sur un tableau : filtre les lignes, Ctrl pour en ajouter, effacer le filtre', async ({ page }) => {
  await insert(page, 'slicer', '0');
  const slicer = page.getByTestId('slicer');
  await expect(slicer).toBeVisible();
  const items = slicer.getByTestId('slicer-item');
  await expect(items).toHaveText(['Est', 'Nord', 'Sud']);

  await items.filter({ hasText: 'Nord' }).click();
  await expect.poll(() => hiddenRows(page)).toEqual([2, 4, 5]);
  await items.filter({ hasText: 'Sud' }).click({ modifiers: ['ControlOrMeta'] });
  await expect.poll(() => hiddenRows(page)).toEqual([4]);

  await slicer.getByTestId('slicer-clear').click();
  await expect.poll(() => hiddenRows(page)).toEqual([]);
});

test('Chronologie : une période, puis une plage avec Maj', async ({ page }) => {
  await insert(page, 'timeline', '1');
  const timeline = page.getByTestId('timeline');
  await expect(timeline).toBeVisible();
  const period = (key: string) => timeline.locator(`[data-period="${key}"]`);
  await expect(period('2024-01')).toBeVisible();
  await period('2024-02').click();
  await expect(timeline.getByTestId('timeline-range')).toHaveText('févr. 2024');
  await expect.poll(() => hiddenRows(page)).toEqual([1, 3, 4, 5, 6]);
  await period('2024-04').click({ modifiers: ['Shift'] });
  await expect(timeline.getByTestId('timeline-range')).toHaveText('févr. 2024 – avr. 2024');
  await expect.poll(() => hiddenRows(page)).toEqual([1, 5, 6]);
  await timeline.getByTestId('timeline-level').selectOption('years');
  await expect(timeline.getByTestId('timeline-range')).toHaveText('Toutes les périodes');
  await expect.poll(() => hiddenRows(page)).toEqual([]);
});

test('Segments et chronologie combinés (ET)', async ({ page }) => {
  await insert(page, 'slicer', '0');
  await insert(page, 'timeline', '1');
  await page.getByTestId('slicer-item').filter({ hasText: 'Nord' }).click();
  await page.locator('[data-period="2024-03"]').click();
  await expect.poll(() => hiddenRows(page)).toEqual([1, 2, 4, 5, 6]);
});

test('Supprimer le segment rend les lignes masquées', async ({ page }) => {
  await insert(page, 'slicer', '0');
  await page.getByTestId('slicer-item').filter({ hasText: 'Est' }).click();
  await expect.poll(() => hiddenRows(page)).toEqual([1, 2, 3, 5, 6]);
  await run(page, `const d = ws.getAllFloatDoms()[0]; ws.removeFloatDom(d.id);`);
  await expect.poll(() => hiddenRows(page), { timeout: 5000 }).toEqual([]);
});

test('Segment sur un tableau croisé dynamique', async ({ page }) => {
  await page.getByRole('tab', { name: 'Insérer', exact: true }).click();
  await page.locator('[data-u-command="grillon.pivot.create"]').first().click();
  const dialog = page.getByTestId('dialog-pivot');
  const fields = dialog.locator('.g-pivot-field');
  await fields.nth(0).locator('select').first().selectOption('rows');
  await fields.nth(2).locator('select').first().selectOption('values');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect.poll(() => run<string[][]>(page, `return ws.getRange('A3:B7').getDisplayValues();`)).toEqual([
    ['Étiquettes de lignes', 'Somme de Montant'],
    ['Est', '40'],
    ['Nord', '100'],
    ['Sud', '70'],
    ['Total général', '210'],
  ]);
  await run(page, `wb.setActiveRange(ws.getRange('A4')); api.executeCommand('grillon.insert.slicer');`);
  const picker = page.getByTestId('dialog-slicer');
  await picker.getByTestId('slicer-field').selectOption('0');
  await picker.getByRole('button', { name: 'OK' }).click();
  await page.getByTestId('slicer-item').filter({ hasText: 'Sud' }).click();
  await expect.poll(() => run<string[][]>(page, `return ws.getRange('A3:B5').getDisplayValues();`)).toEqual([
    ['Étiquettes de lignes', 'Somme de Montant'],
    ['Sud', '70'],
    ['Total général', '70'],
  ]);
  await page.getByTestId('slicer-clear').click();
  await expect.poll(() => run<string>(page, `return ws.getRange('A7').getDisplayValue();`)).toBe('Total général');
  expect(await run<string>(page, `return ws.getRange('B7').getDisplayValue();`)).toBe('210');
});

test('Le segment et son filtre sont conservés après rechargement', async ({ page }) => {
  await insert(page, 'slicer', '0');
  await page.getByTestId('slicer-item').filter({ hasText: 'Est' }).click();
  await expect.poll(() => hiddenRows(page)).toEqual([1, 2, 3, 5, 6]);
  await page.evaluate(`window.grillon.doc.autosaveNow()`);
  await gotoApp(page, false);
  const items = page.getByTestId('slicer-item');
  await expect(items).toHaveText(['Est', 'Nord', 'Sud']);
  await expect(items.filter({ hasText: 'Est' })).toHaveClass(/selected/);
  await expect(items.filter({ hasText: 'Nord' })).not.toHaveClass(/selected/);
  expect(await hiddenRows(page)).toEqual([1, 2, 3, 5, 6]);
  // Effacer le filtre après rechargement rend bien les lignes (l'état « masqué par le segment » est enregistré).
  await page.getByTestId('slicer-clear').click();
  await expect.poll(() => hiddenRows(page)).toEqual([]);
});

test('Enregistrer en .xlsx avec un segment : les lignes masquées sont écrites, pas d’erreur', async ({ page }) => {
  await insert(page, 'slicer', '0');
  await page.getByTestId('slicer-item').filter({ hasText: 'Est' }).click();
  await expect.poll(() => hiddenRows(page)).toEqual([1, 2, 3, 5, 6]);
  const size = await page.evaluate(`window.grillon.doc.exportXlsxBytes().then((b) => b.length)`);
  expect(size).toBeGreaterThan(1000);
});

test('Segment et chronologie gardés dans le .xlsx : enregistrer puis rouvrir', async ({ page }) => {
  await insert(page, 'slicer', '0');
  await insert(page, 'timeline', '1');
  await page.getByTestId('slicer-item').filter({ hasText: 'Nord' }).click();
  await page.getByTestId('timeline').locator('[data-period="2024-03"]').click();
  await expect.poll(() => hiddenRows(page)).toEqual([1, 2, 4, 5, 6]);
  await page.evaluate(`(async () => { const d = window.grillon.doc; const b = await d.exportXlsxBytes(); await d.openBytes('segments.xlsx', b); })()`);
  const items = page.getByTestId('slicer-item');
  await expect(items).toHaveText(['Est', 'Nord', 'Sud']);
  await expect(items.filter({ hasText: 'Nord' })).toHaveClass(/selected/);
  await expect(page.getByTestId('timeline-range')).toHaveText('mars 2024');
  expect(await hiddenRows(page)).toEqual([1, 2, 4, 5, 6]);
  // La plage n'est pas devenue un tableau, et effacer les filtres rend toutes les lignes.
  expect(await run<number>(page, `return Object.keys(JSON.parse(wb.save().resources.find((r) => r.name === 'SHEET_TABLE_PLUGIN')?.data ?? '{}')).length;`)).toBe(0);
  await page.getByTestId('slicer').getByTestId('slicer-clear').click();
  await page.getByTestId('timeline').getByTestId('slicer-clear').click();
  await expect.poll(() => hiddenRows(page)).toEqual([]);
});

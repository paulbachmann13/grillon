import { expect, test, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { gotoApp } from './helpers';

async function run<T = unknown>(page: Page, body: string): Promise<T> {
  return page.evaluate(`(async () => { const api = window.grillon.api; const wb = api.getActiveWorkbook(); const ws = wb.getActiveSheet(); ${body} })()`) as Promise<T>;
}
const visibleRows = (page: Page) => run<string>(page, `return Array.from({ length: 8 }, (_, i) => ws.getSheet().getRowVisible(i) ? '1' : '0').join('');`);

test.beforeEach(async ({ page }) => {
  await gotoApp(page);
});

test('Plan : boutons de niveau et − / + dans la marge, comme Excel', async ({ page }) => {
  await run(page, `
    ws.getRange('A1:A8').setValues([['T'], ['a'], ['b'], ['c'], ['d'], ['e'], ['Total'], ['x']]);
    wb.setActiveRange(ws.getRange('A2:A6')); await api.executeCommand('grillon.data.group-rows');
    wb.setActiveRange(ws.getRange('A3:A4')); await api.executeCommand('grillon.data.group-rows');`);
  const gutter = page.getByTestId('outline-gutter');
  await expect(gutter.getByTestId('outline-row-level-3')).toBeVisible();
  await gutter.getByTestId('outline-row-level-1').click();
  await expect.poll(() => visibleRows(page)).toBe('10000011');
  await gutter.getByTestId('outline-row-toggle-6').click(); // + sur la ligne de synthèse (ligne 7)
  await expect.poll(() => visibleRows(page)).toBe('11001111');
  await gutter.getByTestId('outline-row-toggle-4').click(); // + du sous-groupe
  await expect.poll(() => visibleRows(page)).toBe('11111111');
  await gutter.getByTestId('outline-row-toggle-6').click(); // − : tout le groupe se replie
  await expect.poll(() => visibleRows(page)).toBe('10000011');
});

test('Protection : seules les cellules déverrouillées restent modifiables', async ({ page }) => {
  await run(page, `wb.setActiveRange(ws.getRange('B2:B3'));`);
  await page.locator('#univer').click({ position: { x: 700, y: 500 } });
  await run(page, `wb.setActiveRange(ws.getRange('B2:B3'));`);
  await page.keyboard.press('Control+Digit1');
  const dialog = page.getByTestId('dialog-format-cells');
  await dialog.getByRole('tab', { name: 'Protection' }).click();
  await expect(dialog.getByTestId('format-locked')).toBeChecked();
  await dialog.getByTestId('format-locked').uncheck();
  await dialog.getByRole('button', { name: 'OK' }).click();

  await page.locator('[data-u-command="grillon.format"]').click();
  await page.getByRole('menuitem', { name: 'Protéger la feuille…' }).click();
  await page.getByTestId('dialog-protect').getByRole('button', { name: 'OK' }).click();

  // Saisie dans une cellule déverrouillée : acceptée.
  await run(page, `wb.setActiveRange(ws.getRange('B3'));`);
  await page.keyboard.type('42');
  await page.keyboard.press('Enter');
  await expect.poll(() => run(page, `return ws.getRange('B3').getDisplayValue();`)).toBe('42');

  // Saisie dans une cellule verrouillée : refusée avec le message d'Excel.
  await run(page, `wb.setActiveRange(ws.getRange('D5'));`);
  await page.keyboard.type('9');
  const alert = page.getByRole('dialog').filter({ hasText: 'feuille protégée' });
  await expect(alert).toBeVisible();
  await alert.getByRole('button', { name: 'OK' }).click();
  await page.keyboard.press('Escape');
  expect(await run(page, `return ws.getRange('D5').getValue();`)).toBeNull();
  const ok = await run<boolean>(page, `wb.setActiveRange(ws.getRange('D5')); return await api.executeCommand('sheet.command.set-style', { style: { type: 'bl', value: 1 } });`);
  expect(ok).toBe(false);
});

test('Mettre sous forme de tableau : galerie des styles, changement de style et .xlsx', async ({ page }) => {
  await run(page, `ws.getRange('B2:D4').setValues([['Nom', 'Qté', 'Prix'], ['a', 1, 2], ['b', 3, 4]]); wb.setActiveRange(ws.getRange('C3'));`);
  await page.locator('[data-u-command="grillon.home.format-as-table"]').first().click();
  const gallery = page.getByTestId('dialog-table-styles');
  await expect(gallery.getByRole('button', { name: 'Moyen 2', exact: true })).toBeVisible();
  await gallery.getByRole('button', { name: 'Moyen 7', exact: true }).click();
  await expect(gallery).toHaveCount(0);
  await expect.poll(async () => (await run<any>(page, `return ws.getRange('B2').getCellStyleData();`))?.bg?.rgb).toBe('#70AD47');
  await run(page, `api.executeCommand('grillon.data.table.style');`);
  await page.getByTestId('dialog-table-styles').getByRole('button', { name: 'Foncé 2', exact: true }).click();
  await expect.poll(async () => (await run<any>(page, `return ws.getRange('B3').getCellStyleData();`))?.cl?.rgb).toBe('#FFFFFF');
  const bytes = await page.evaluate(async () => Array.from(await (window as any).grillon.doc.exportXlsxBytes()));
  const zip = await JSZip.loadAsync(Uint8Array.from(bytes as number[]));
  expect(await zip.file('xl/tables/table1.xml')!.async('string')).toContain('name="TableStyleDark2"');
});

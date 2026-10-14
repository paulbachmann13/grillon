import { expect, test, type Page } from '@playwright/test';
import { gotoApp } from './helpers';

async function run<T = unknown>(page: Page, body: string): Promise<T> {
  return page.evaluate(`(async () => { const api = window.grillon.api; const wb = api.getActiveWorkbook(); const ws = wb.getActiveSheet(); ${body} })()`) as Promise<T>;
}

test.beforeEach(async ({ page }) => {
  await gotoApp(page);
  await run(page, `ws.getRange('A1:C6').setValues([['Ville','Produit','Montant'],['Lyon','Café',10],['Paris','Thé',5],['Lyon','Thé',7],['Lyon','Café',3],['Nîmes','Café',8]]); wb.setActiveRange(ws.getRange('A1'));`);
});

test('insérer un graphique depuis le ruban, le modifier, et le retrouver après export .xlsx', async ({ page }) => {
  await page.getByRole('tab', { name: 'Insérer', exact: true }).click();
  await page.locator('[data-u-command="grillon.insert.chart"]').first().click();
  await page.getByRole('menuitem', { name: 'Histogramme (colonnes)', exact: true }).click();
  const chart = page.getByTestId('chart');
  await expect(chart).toBeVisible();
  await expect(chart.locator('canvas')).toHaveCount(1);
  await chart.hover();
  await chart.getByRole('button', { name: 'Modifier le graphique' }).click();
  const dialog = page.getByTestId('dialog-chart');
  await dialog.getByLabel('Titre du graphique').fill('Montants par ville');
  await dialog.getByLabel('Type').selectOption('line');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect.poll(() => run<string>(page, `return ws.getAllFloatDoms()[0].data.spec.title + '|' + ws.getAllFloatDoms()[0].data.spec.type;`)).toBe('Montants par ville|line');
  // Export puis réouverture : le graphique est toujours là.
  await page.evaluate(async () => {
    const g = (window as any).grillon;
    await g.doc.openBytes('g.xlsx', await g.doc.exportXlsxBytes());
  });
  await expect(page.getByTestId('chart')).toBeVisible();
  expect(await run<string>(page, `return ws.getAllFloatDoms()[0].data.spec.title;`)).toBe('Montants par ville');
});

test('tableau croisé dynamique : création puis actualisation', async ({ page }) => {
  await page.getByRole('tab', { name: 'Insérer', exact: true }).click();
  await page.locator('[data-u-command="grillon.pivot.create"]').first().click();
  const dialog = page.getByTestId('dialog-pivot');
  const fields = dialog.locator('.g-pivot-field');
  await fields.nth(0).locator('select').first().selectOption('rows');
  await fields.nth(2).locator('select').first().selectOption('values');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect.poll(() => run<string[][]>(page, `return ws.getRange('A3:B7').getDisplayValues();`)).toEqual([
    ['Étiquettes de lignes', 'Somme de Montant'],
    ['Lyon', '20'],
    ['Nîmes', '8'],
    ['Paris', '5'],
    ['Total général', '33'],
  ]);
  await run(page, `wb.getSheetByName('Feuil1').getRange('C3').setValue(50);`);
  await page.getByRole('tab', { name: 'Données', exact: true }).click();
  await page.locator('[data-u-command="grillon.pivot"]').first().click();
  await page.getByRole('menuitem', { name: 'Actualiser', exact: true }).click();
  await expect.poll(() => run<string>(page, `return ws.getRange('B6').getDisplayValue();`)).toBe('50');
});

test('valeur cible', async ({ page }) => {
  await run(page, `ws.getRange('E1').setValue(1); ws.getRange('F1').setValue('=E1*E1+3*E1'); wb.setActiveRange(ws.getRange('F1'));`);
  await page.getByRole('tab', { name: 'Données', exact: true }).click();
  await page.locator('[data-u-command="grillon.data.goal-seek"]').first().click();
  const dialog = page.getByTestId('dialog-goal-seek');
  await dialog.getByLabel('Valeur à atteindre :').fill('10');
  await dialog.getByLabel('Cellule à modifier :').fill('E1');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(page.getByTestId('dialog-goal-status')).toContainText('a été trouvée');
  await page.getByTestId('dialog-goal-status').getByRole('button', { name: 'OK' }).click();
  const x = await run<number>(page, `return Number(ws.getRange('E1').getCellDataGrid()[0][0].v);`);
  expect(x).toBeCloseTo(2, 5);
});

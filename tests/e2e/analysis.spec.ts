import { expect, test, type Page } from '@playwright/test';
import { gotoApp } from './helpers';

async function run<T = unknown>(page: Page, body: string): Promise<T> {
  return page.evaluate(`(async () => { const api = window.grillon.api; const wb = api.getActiveWorkbook(); const ws = wb.getActiveSheet(); ${body} })()`) as Promise<T>;
}

test.beforeEach(async ({ page }) => {
  await gotoApp(page);
  await run(page, `ws.getRange('A1:C6').setValues([['Ville','Produit','Montant'],['Lyon','Café',10],['Paris','Thé',5],['Lyon','Thé',7],['Lyon','Café',3],['Nîmes','Café',8]]); wb.setActiveRange(ws.getRange('A1'));`);
});

test('insérer un graphique depuis le ruban, régler ses options, et le retrouver après export .xlsx', async ({ page }) => {
  await run(page, `wb.setActiveRange(ws.getRange('B1:C6'));`);
  await page.getByRole('tab', { name: 'Insérer', exact: true }).click();
  await page.locator('[data-u-command="grillon.insert.chart"]').first().click();
  await page.getByRole('menuitem', { name: 'Histogramme groupé', exact: true }).click();
  const chart = page.getByTestId('chart');
  await expect(chart).toBeVisible();
  await expect(chart.locator('canvas')).toHaveCount(1);
  await chart.hover();
  await chart.getByRole('button', { name: 'Modifier le graphique' }).click();
  const dialog = page.getByTestId('dialog-chart');
  await expect(dialog.getByTestId('chart-preview').locator('canvas')).toHaveCount(1);
  // Éléments : titre, étiquettes de données.
  await dialog.getByRole('tab', { name: 'Éléments' }).click();
  await dialog.getByLabel('Titre du graphique', { exact: true }).fill('Montants par ville');
  await dialog.getByLabel('Étiquettes de données').check();
  // Axes : maximum fixé.
  await dialog.getByRole('tab', { name: 'Axes' }).click();
  await dialog.getByLabel('Maximum').first().fill('50');
  // Séries : courbe de tendance linéaire avec équation.
  await dialog.getByRole('tab', { name: 'Séries' }).click();
  await dialog.getByLabel('Courbe de tendance').first().selectOption('linear');
  await dialog.getByLabel('Afficher l’équation').check();
  // Type : courbes avec marques.
  await dialog.getByRole('tab', { name: 'Type' }).click();
  await dialog.locator('button.g-kind', { hasText: /^Courbes avec marques$/ }).click();
  await dialog.getByRole('button', { name: 'Appliquer' }).click();
  const spec = () => run<any>(page, 'return ws.getAllFloatDoms()[0].data.spec;');
  await expect.poll(async () => (await spec()).title).toBe('Montants par ville');
  const s = await spec();
  expect(s.type).toBe('line');
  expect(s.markers).toBe(true);
  expect(s.dataLabels).toMatchObject({ show: true });
  expect(s.axes.y.max).toBe(50);
  expect(s.series[0].trendlines[0]).toMatchObject({ type: 'linear', showEquation: true });
  // Export puis réouverture : le graphique et ses options sont toujours là.
  await page.evaluate(async () => {
    const g = (window as any).grillon;
    await g.doc.openBytes('g.xlsx', await g.doc.exportXlsxBytes());
  });
  await expect(page.getByTestId('chart')).toBeVisible();
  const back = await spec();
  expect(back.title).toBe('Montants par ville');
  expect(back.markers).toBe(true);
  expect(back.axes.y.max).toBe(50);
  expect(back.series[0].trendlines[0]).toMatchObject({ type: 'linear', showEquation: true });
});

test('« Tous les graphiques » : choisir un type dans la boîte puis insérer', async ({ page }) => {
  await run(page, `wb.setActiveRange(ws.getRange('B1:C6'));`);
  await page.getByRole('tab', { name: 'Insérer', exact: true }).click();
  await page.locator('[data-u-command="grillon.insert.chart"]').first().click();
  await page.getByRole('menuitem', { name: 'Tous les graphiques…', exact: true }).click();
  const dialog = page.getByTestId('dialog-chart');
  await dialog.locator('button.g-kind', { hasText: /^Anneau$/ }).click();
  await dialog.getByRole('button', { name: 'Insérer' }).click();
  await expect(page.getByTestId('chart')).toBeVisible();
  expect(await run<string>(page, 'return ws.getAllFloatDoms()[0].data.spec.type;')).toBe('doughnut');
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

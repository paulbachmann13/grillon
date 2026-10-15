import { expect, test, type Page } from '@playwright/test';
import { gotoApp } from './helpers';

async function run<T = unknown>(page: Page, body: string): Promise<T> {
  return page.evaluate(`(async () => { const api = window.grillon.api; const wb = api.getActiveWorkbook(); const ws = wb.getActiveSheet(); ${body} })()`) as Promise<T>;
}
const floatKeys = (page: Page) => run<string[]>(page, 'return ws.getAllFloatDoms().map((d) => d.componentKey);');

test.beforeEach(async ({ page }) => {
  await gotoApp(page);
  await run(page, `ws.getRange('A1:E4').setValues([['Produit','T1','T2','T3','T4'],['Café',10,14,9,20],['Thé',5,3,8,2],['Eau',7,7,6,9]]); wb.setActiveRange(ws.getRange('B2:E4'));`);
});

test('Insérer → Symbole : le caractère est ajouté à la cellule', async ({ page }) => {
  await run(page, `wb.setActiveRange(ws.getRange('A1'));`);
  await page.getByRole('tab', { name: 'Insérer', exact: true }).click();
  await page.locator('[data-u-command="grillon.insert.symbol"]').first().click();
  const dialog = page.getByTestId('dialog-symbols');
  await dialog.getByTestId('symbols').getByRole('button', { name: '€', exact: true }).click();
  await expect.poll(() => run<string>(page, `return ws.getRange('A1').getDisplayValue();`)).toBe('Produit€');
});

test('mini-graphiques : création, modification, export/réimport', async ({ page }) => {
  await page.getByRole('tab', { name: 'Insérer', exact: true }).click();
  await page.locator('[data-u-command="grillon.insert.sparkline"]').first().click();
  await page.getByRole('menuitem', { name: 'Histogramme', exact: true }).click();
  const dialog = page.getByTestId('dialog-sparkline');
  await expect(dialog.getByLabel('Plage d’emplacement (ex. G2:G10)')).toHaveValue('F2:F4');
  await dialog.getByLabel('Point haut').check();
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect.poll(async () => (await floatKeys(page)).filter((k) => k === 'GrillonSparkline').length).toBe(3);
  const spec = await run<any>(page, `return ws.getAllFloatDoms()[0].data.spec;`);
  expect(spec).toMatchObject({ type: 'column', high: true });
  // Export puis réouverture : les trois mini-graphiques sont toujours là.
  await page.evaluate(async () => {
    const g = (window as any).grillon;
    await g.doc.openBytes('s.xlsx', await g.doc.exportXlsxBytes());
  });
  await expect.poll(async () => (await floatKeys(page)).filter((k) => k === 'GrillonSparkline').length).toBe(3);
  expect(await run<any>(page, `return ws.getAllFloatDoms().map((d) => d.data.spec.dataRef).sort();`)).toEqual(['Feuil1!$B$2:$E$2', 'Feuil1!$B$3:$E$3', 'Feuil1!$B$4:$E$4']);
});

test('formes : insertion, modification du texte, export/réimport', async ({ page }) => {
  await run(page, `wb.setActiveRange(ws.getRange('G2'));`);
  await page.getByRole('tab', { name: 'Insérer', exact: true }).click();
  await page.locator('[data-u-command="grillon.insert.shape"]').first().click();
  await page.getByRole('menuitem', { name: 'Rectangle', exact: true }).click();
  const shape = page.getByTestId('shape');
  await expect(shape).toBeVisible();
  await shape.hover();
  await shape.getByRole('button', { name: 'Modifier la forme' }).click();
  const dialog = page.getByTestId('dialog-shape');
  await dialog.locator('textarea').fill('Total annuel');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect.poll(() => run<string>(page, `return ws.getAllFloatDoms()[0].data.spec.text;`)).toBe('Total annuel');
  await page.evaluate(async () => {
    const g = (window as any).grillon;
    await g.doc.openBytes('f.xlsx', await g.doc.exportXlsxBytes());
  });
  await expect.poll(() => run<string>(page, `return JSON.stringify(ws.getAllFloatDoms().map((d) => [d.componentKey, d.data.spec.kind, d.data.spec.text]));`)).toBe('[["GrillonShape","rect","Total annuel"]]');
});

test('Analyse rapide : le bouton apparaît et ajoute des totaux', async ({ page }) => {
  const button = page.locator('[data-testid="quick-analysis-button"]:visible');
  await expect(button).toBeVisible({ timeout: 10_000 });
  await button.click();
  const panel = page.getByTestId('quick-analysis');
  await panel.getByRole('tab', { name: 'Totaux' }).click();
  await panel.getByRole('menuitem', { name: 'Somme sous les colonnes' }).click();
  await expect.poll(() => run<string>(page, `return ws.getRange('B5').getFormula();`)).toBe('=SUM(B2:B4)');
});

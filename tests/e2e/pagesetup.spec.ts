import { expect, test } from '@playwright/test';
import { gotoApp } from './helpers';

test('Mise en page : réglages enregistrés dans la feuille, appliqués à l’impression et exportés en .xlsx', async ({ page }) => {
  await gotoApp(page);
  await page.evaluate(() => {
    const ws = (window as any).grillon.api.getActiveWorkbook().getActiveSheet();
    ws.getRange('A1:B4').setValues([['Nom', 'Montant'], ['Café', 1.5], ['Thé', 2], ['Eau', 1]]);
  });
  await page.getByTestId('file-menu').click();
  await page.getByTestId('file-page-setup').click();
  const dialog = page.getByTestId('dialog-page-setup');
  await dialog.getByLabel('Taille du papier').selectOption('A3');
  await dialog.getByLabel('Orientation :').selectOption('landscape');
  await dialog.getByLabel('Marges prédéfinies').selectOption('narrow');
  await dialog.getByLabel('Lignes à répéter en haut (ex. 1:2)').fill('1:1');
  await dialog.getByLabel('Imprimer le quadrillage').check();
  await dialog.getByLabel('Au centre').first().fill('Budget &A');
  await dialog.getByLabel('À droite').last().fill('Page &P / &N');
  await dialog.getByRole('button', { name: 'OK' }).click();
  const saved = await page.evaluate(() => (window as any).grillon.api.getActiveWorkbook().getActiveSheet().getCustomMetadata().grillon.page);
  expect(saved).toMatchObject({ paper: 'A3', orientation: 'landscape', titleRows: '1:1', gridlines: true, header: { center: 'Budget &A' }, footer: { right: 'Page &P / &N' } });
  expect(saved.margins.left).toBe(0.25);
  await expect(page.getByTestId('doc-title')).toHaveText(/•$/);

  // Impression : le document imprimable reprend papier, orientation, titres et en-têtes.
  await page.getByTestId('file-menu').click();
  await page.getByTestId('file-print').click();
  await page.getByTestId('dialog-print').getByRole('button', { name: 'Imprimer…' }).click();
  const html = await page.locator('iframe[aria-hidden="true"]').evaluate((f) => (f as HTMLIFrameElement).contentDocument!.documentElement.outerHTML);
  expect(html).toContain('size: 297mm 420mm landscape');
  expect(html).toContain('@top-center { content: "Budget Feuil1"');
  expect(html).toContain('@bottom-right { content: "Page " counter(page) " / " counter(pages)');
  expect(html).toContain('<thead><tr style');

  // Export / réimport : la mise en page est conservée.
  await page.evaluate(async () => {
    const g = (window as any).grillon;
    await g.doc.openBytes('p.xlsx', await g.doc.exportXlsxBytes());
  });
  await expect.poll(() => page.evaluate(() => (window as any).grillon.api.getActiveWorkbook().getActiveSheet().getCustomMetadata()?.grillon?.page?.paper)).toBe('A3');
});

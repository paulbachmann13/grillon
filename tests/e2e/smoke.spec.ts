import { expect, test } from '@playwright/test';
import { display, gotoApp } from './helpers';

test('l’application démarre en français avec ruban, barre de formule et onglet de feuille', async ({ page }) => {
  await gotoApp(page);
  for (const tab of ['Démarrer', 'Insérer', 'Formules', 'Données', 'Vue']) {
    await expect(page.getByText(tab, { exact: true }).first()).toBeVisible();
  }
  await expect(page.getByText('Feuil1', { exact: true })).toBeVisible();
  await expect(page.getByTestId('file-menu')).toHaveText('Fichier');
  await expect(page).toHaveTitle(/Classeur1 — Grillon/);
});

test('saisie d’une formule et recalcul', async ({ page }) => {
  await gotoApp(page);
  await page.evaluate(() => {
    const ws = (window as any).grillon.api.getActiveWorkbook().getActiveSheet();
    ws.getRange('A1').setValue(2);
    ws.getRange('A2').setValue(3);
    ws.getRange('A3').setValue('=SUM(A1:A2)*10');
  });
  await expect.poll(() => display(page, 'A3')).toBe('50');
  await page.evaluate(() => (window as any).grillon.api.getActiveWorkbook().getActiveSheet().getRange('A1').setValue(7));
  await expect.poll(() => display(page, 'A3')).toBe('100');
  await expect(page.getByTestId('doc-title')).toHaveText(/•$/);
});

test('mode sombre (Vue) et impression / PDF (Fichier)', async ({ page }) => {
  await gotoApp(page);
  await page.evaluate(() => {
    const ws = (window as any).grillon.api.getActiveWorkbook().getActiveSheet();
    ws.getRange('A1:B2').setValues([['Nom', 'Montant'], ['Café', 1.5]]);
  });
  await page.getByRole('tab', { name: 'Vue', exact: true }).click();
  await page.locator('[data-u-command="grillon.view.dark"]').first().click();
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');
  await page.locator('[data-u-command="grillon.view.dark"]').first().click();
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'light');

  await page.getByTestId('file-menu').click();
  await page.getByTestId('file-print').click();
  await page.getByTestId('dialog-print').getByRole('button', { name: 'Imprimer…' }).click();
  const frame = page.frameLocator('iframe[aria-hidden="true"]');
  await expect(frame.locator('td', { hasText: 'Café' })).toHaveCount(1);
  await expect(frame.locator('td', { hasText: '1,5' })).toHaveCount(1);
});

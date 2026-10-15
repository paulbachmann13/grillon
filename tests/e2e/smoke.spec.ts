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

test('la grille s’agrandit au-delà de la ligne 1000 et de la colonne Z', async ({ page }) => {
  await gotoApp(page);
  const size = () => page.evaluate(() => {
    const ws = (window as any).grillon.api.getActiveWorkbook().getActiveSheet();
    return [ws.getMaxRows(), ws.getMaxColumns()];
  });
  expect(await size()).toEqual([1000, 26]);
  await page.evaluate(() => {
    const wb = (window as any).grillon.api.getActiveWorkbook();
    wb.setActiveRange(wb.getActiveSheet().getRange('Z1000'));
  });
  await expect.poll(size).toEqual([2000, 52]);
  // Saisie en AA1001 (au-delà de l'ancienne grille), sans marquer le classeur modifié avant.
  await expect(page.getByTestId('doc-title')).not.toHaveText(/•$/);
  await page.evaluate(() => (window as any).grillon.api.getActiveWorkbook().getActiveSheet().getRange('AA1001').setValue(7));
  await expect.poll(() => display(page, 'AA1001')).toBe('7');
});

test('Vue → Texte qui déborde : réglage mémorisé', async ({ page }) => {
  await gotoApp(page);
  await page.getByRole('tab', { name: 'Vue', exact: true }).click();
  const button = page.locator('[data-u-command="grillon.view.text-overflow"]').first();
  await button.click();
  expect(await page.evaluate(() => localStorage.getItem('grillon-text-overflow'))).toBe('overflow');
  await button.click();
  expect(await page.evaluate(() => localStorage.getItem('grillon-text-overflow'))).toBeNull();
  await expect(page.getByTestId('doc-title')).not.toHaveText(/•$/);
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

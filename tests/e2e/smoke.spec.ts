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

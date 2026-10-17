import { expect, test, type Page } from '@playwright/test';
import { gotoApp } from './helpers';

async function run<T = unknown>(page: Page, body: string): Promise<T> {
  return page.evaluate(`(async () => { const api = window.grillon.api; const wb = api.getActiveWorkbook(); const ws = wb.getActiveSheet(); ${body} })()`) as Promise<T>;
}
const style = (page: Page, a1: string) => run<any>(page, `return ws.getRange('${a1}').getCellStyleData() || {};`);

test.beforeEach(async ({ page }) => {
  await gotoApp(page);
  await run(page, `ws.getRange('B2:D4').setValues([[1,2,3],[4,5,6],[7,8,9]]); wb.setActiveRange(ws.getRange('B2:D4'));`);
});

test('Styles de cellule : « Satisfaisant » colore la sélection en une étape annulable', async ({ page }) => {
  await page.locator('[data-u-command="grillon.cell-styles"]').first().click();
  await page.getByTestId('dialog-cell-styles').getByRole('button', { name: 'Satisfaisant', exact: true }).click();
  await expect(page.getByTestId('dialog-cell-styles')).toHaveCount(0);
  await expect.poll(async () => (await style(page, 'C3')).bg?.rgb).toBe('#C6EFCE');
  await run(page, `await api.executeCommand('univer.command.undo');`);
  await expect.poll(async () => (await style(page, 'C3')).bg?.rgb ?? null).toBeNull();
});

test('Format de cellule (Ctrl+1) : monétaire, contour et remplissage', async ({ page }) => {
  await page.locator('#univer').click({ position: { x: 700, y: 500 } });
  await run(page, `wb.setActiveRange(ws.getRange('B2:D4'));`);
  await page.keyboard.press('Control+Digit1');
  const dialog = page.getByTestId('dialog-format-cells');
  await expect(dialog).toBeVisible();
  await dialog.getByLabel('Catégorie').selectOption('currency');
  await expect(dialog.locator('.g-format-preview')).toContainText('€');
  await dialog.getByRole('tab', { name: 'Bordure' }).click();
  await dialog.getByRole('button', { name: 'Contour' }).click();
  await dialog.getByRole('tab', { name: 'Remplissage' }).click();
  await dialog.locator('.g-swatches button[title="#FFF2CC"]').click();
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect.poll(() => run<string>(page, `return ws.getRange('C3').getDisplayValue();`)).toBe('5,00 €');
  const b2 = await style(page, 'B2');
  expect(b2.bg.rgb).toBe('#FFF2CC');
  expect(b2.bd.t.s).toBe(1);
  expect(b2.bd.l.s).toBe(1);
  expect((await style(page, 'C3')).bd?.t).toBeUndefined();
});

test('Protéger la feuille avec mot de passe : plus de saisie, puis déprotection', async ({ page }) => {
  await page.locator('[data-u-command="grillon.format"]').click();
  await page.getByRole('menuitem', { name: 'Protéger la feuille…' }).click();
  const dialog = page.getByTestId('dialog-protect');
  await dialog.getByLabel('Mot de passe (facultatif)').fill('secret');
  await dialog.getByLabel('Confirmer le mot de passe').fill('secret');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog).toHaveCount(0);
  await expect(page.getByTestId('doc-title')).toHaveText(/•$/);
  const ok = await run<boolean>(page, `wb.setActiveRange(ws.getRange('F6')); return await api.executeCommand('sheet.command.set-range-values', { value: [[{ v: 9 }]] });`);
  expect(ok).toBe(false);
  // Univer prévient (comme Excel) que la feuille est protégée.
  const alert = page.getByRole('dialog').filter({ hasText: 'protégée' });
  await expect(alert).toBeVisible();
  await alert.getByRole('button', { name: 'OK' }).click();
  // Export / réimport : toujours protégée.
  await page.evaluate(async () => {
    const g = (window as any).grillon;
    await g.doc.openBytes('p.xlsx', await g.doc.exportXlsxBytes());
  });
  await expect.poll(() => run<boolean>(page, `return !!ws.getCustomMetadata()?.grillon?.protection?.hashValue;`)).toBe(true);
  // Le menu Format est désactivé sur une feuille protégée : on passe par le menu Fichier.
  await page.getByTestId('file-menu').click();
  await page.getByTestId('file-unprotect').click();
  const un = page.getByTestId('dialog-unprotect');
  await un.getByLabel('Mot de passe (facultatif)').fill('faux');
  await un.getByRole('button', { name: 'OK' }).click();
  await expect(un.getByRole('alert')).toHaveText('Mot de passe incorrect.');
  await un.getByLabel('Mot de passe (facultatif)').fill('secret');
  await un.getByRole('button', { name: 'OK' }).click();
  await expect(un).toHaveCount(0);
  const ok2 = await run<boolean>(page, `wb.setActiveRange(ws.getRange('F6')); return await api.executeCommand('sheet.command.set-range-values', { value: [[{ v: 9 }]] });`);
  expect(ok2).toBe(true);
});

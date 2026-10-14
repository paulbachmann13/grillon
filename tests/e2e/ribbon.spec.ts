import { expect, test, type Page } from '@playwright/test';
import { display, gotoApp } from './helpers';

/** Exécute du code avec la façade Univer (feuille active). */
async function withSheet<T>(page: Page, fn: string): Promise<T> {
  return page.evaluate(`(() => { const api = window.grillon.api; const wb = api.getActiveWorkbook(); const ws = wb.getActiveSheet(); ${fn} })()`) as Promise<T>;
}

async function select(page: Page, a1: string): Promise<void> {
  await withSheet(page, `wb.setActiveRange(ws.getRange('${a1}'));`);
}

/** Ouvre un menu déroulant du ruban (bouton identifié par son id de menu) et choisit un élément. */
async function ribbonMenu(page: Page, menuId: string, item: string): Promise<void> {
  await page.locator(`[data-u-command="${menuId}"]`).click();
  await page.getByRole('menuitem', { name: item, exact: true }).click();
}

test.beforeEach(async ({ page, context }) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await gotoApp(page);
  await withSheet(page, `
    ws.getRange('A1:A5').setValues([[1], [2], [3], [4], [5]]);
    ws.getRange('B1').setValue('=A1*2');
    ws.getRange('B1').setFontWeight('bold');
  `);
});

test('Insérer / Supprimer des lignes et des colonnes depuis le ruban', async ({ page }) => {
  await select(page, 'A2:A3');
  await ribbonMenu(page, 'grillon.insert', 'Insérer des lignes dans la feuille');
  await expect.poll(() => display(page, 'A2')).toBe('');
  expect(await display(page, 'A4')).toBe('2');

  await ribbonMenu(page, 'grillon.delete', 'Supprimer des lignes dans la feuille');
  await expect.poll(() => display(page, 'A2')).toBe('2');

  await select(page, 'A1');
  await ribbonMenu(page, 'grillon.insert', 'Insérer des colonnes dans la feuille');
  await expect.poll(() => display(page, 'B1')).toBe('1');
  await ribbonMenu(page, 'grillon.delete', 'Supprimer des colonnes dans la feuille');
  await expect.poll(() => display(page, 'A1')).toBe('1');

  // Annuler (Ctrl+Z) rétablit la colonne supprimée.
  await page.keyboard.press('ControlOrMeta+Z');
  await expect.poll(() => display(page, 'B1')).toBe('1');
});

test('Effacer le contenu et masquer des lignes', async ({ page }) => {
  await select(page, 'A1:A2');
  await ribbonMenu(page, 'grillon.clear', 'Effacer le contenu');
  await expect.poll(() => display(page, 'A1')).toBe('');
  expect(await display(page, 'A3')).toBe('3');

  await select(page, 'A4');
  await ribbonMenu(page, 'grillon.format', 'Masquer les lignes');
  await expect.poll(() => withSheet<boolean>(page, `return wb.save().sheets[ws.getSheetId()].rowData?.[3]?.hd === 1;`)).toBe(true);
});

test('Retrait : augmenter puis diminuer', async ({ page }) => {
  await select(page, 'A1');
  const indent = () => withSheet<number>(page, `return ws.getRange('A1').getCellStyleData()?.pd?.l ?? 0;`);
  await page.getByRole('button', { name: 'Augmenter le retrait' }).click();
  await page.getByRole('button', { name: 'Augmenter le retrait' }).click();
  await expect.poll(indent).toBe(18);
  await page.getByRole('button', { name: 'Diminuer le retrait' }).click();
  await expect.poll(indent).toBe(9);
});

test('Copier puis « Coller les formules » : formule adaptée, sans la mise en forme', async ({ page }) => {
  await select(page, 'B1');
  await page.getByRole('button', { name: 'Copier', exact: true }).first().click();
  await select(page, 'B2');
  await page.locator('[data-u-command="sheet.menu.paste"] [aria-haspopup="menu"], [aria-haspopup="menu"][data-u-command="sheet.menu.paste"]').first().click();
  await page.getByRole('menuitem', { name: 'Coller les formules' }).click();
  await expect.poll(() => withSheet<string>(page, `return ws.getRange('B2').getFormula();`)).toBe('=A2*2');
  await expect.poll(() => display(page, 'B2')).toBe('4');
  expect(await withSheet<number>(page, `return ws.getRange('B2').getCellStyleData()?.bl ?? 0;`)).toBe(0);
});

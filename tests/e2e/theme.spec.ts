import { expect, test, type Page } from '@playwright/test';
import { gotoApp } from './helpers';

async function run<T = unknown>(page: Page, body: string): Promise<T> {
  return page.evaluate(`(async () => { const api = window.grillon.api; const wb = api.getActiveWorkbook(); const ws = wb.getActiveSheet(); ${body} })()`) as Promise<T>;
}
const background = (page: Page, a1: string) => run<string>(page, `return api.getActiveWorkbook().getActiveSheet().getRange('${a1}').getBackground();`);
const fontOf = (page: Page, a1: string) => run<string>(page, `return api.getActiveWorkbook().getActiveSheet().getRange('${a1}').getFontFamily();`);

test.beforeEach(async ({ page }) => {
  await gotoApp(page);
  await run(page, `ws.getRange('A1').setValue('x'); ws.getRange('A1').setBackground('#4472C4'); ws.getRange('A2').setBackground('#123456'); ws.getRange('A1:A2').setFontFamily('Calibri');`);
});

test('Thèmes : un nouveau thème recolore les couleurs du thème et change les polices', async ({ page }) => {
  await run(page, `api.executeCommand('grillon.view.theme-theme');`);
  const dialog = page.getByTestId('dialog-theme');
  await dialog.getByTestId('theme-item').filter({ hasText: 'Océan' }).click();
  await dialog.getByRole('button', { name: 'Appliquer' }).click();
  await expect.poll(() => background(page, 'A1')).toBe('#0F6E8C');
  // Une couleur qui n'est pas dans le thème ne bouge pas.
  expect(await background(page, 'A2')).toBe('#123456');
  expect(await fontOf(page, 'A1')).toBe('Arial');
  expect(await run<string>(page, `return window.grillon.doc.snapshot().custom.grillon.theme.name;`)).toBe('Océan');
  expect(await run<string>(page, `return window.grillon.doc.dirty ? 'dirty' : 'clean';`)).toBe('dirty');

  // Retour au thème d'origine : les couleurs d'origine reviennent.
  await run(page, `api.executeCommand('grillon.view.theme-theme');`);
  await page.getByTestId('dialog-theme').getByTestId('theme-item').filter({ hasText: 'Office 2013 – 2022' }).click();
  await page.getByTestId('dialog-theme').getByRole('button', { name: 'Appliquer' }).click();
  await expect.poll(() => background(page, 'A1')).toBe('#4472C4');
  expect(await fontOf(page, 'A1')).toBe('Calibri');
});

test('Couleurs : couleur personnalisée du thème', async ({ page }) => {
  await run(page, `api.executeCommand('grillon.view.theme-colors');`);
  const dialog = page.getByTestId('dialog-theme');
  await dialog.getByTestId('theme-color-accent1').fill('#ff0000');
  await dialog.getByRole('button', { name: 'Appliquer' }).click();
  await expect.poll(() => background(page, 'A1')).toBe('#FF0000');
  expect(await run<string>(page, `return window.grillon.doc.snapshot().custom.grillon.theme.colors.name;`)).toBe('Personnalisé');
});

test('Polices : police des titres et du corps', async ({ page }) => {
  await run(page, `api.executeCommand('grillon.view.theme-fonts');`);
  const dialog = page.getByTestId('dialog-theme');
  await dialog.getByTestId('font-list').getByTestId('theme-item').filter({ hasText: 'Georgia – Verdana' }).click();
  await dialog.getByRole('button', { name: 'Appliquer' }).click();
  await expect.poll(() => fontOf(page, 'A1')).toBe('Verdana');
  expect(await background(page, 'A1')).toBe('#4472C4');
});

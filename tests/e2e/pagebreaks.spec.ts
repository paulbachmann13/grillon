import { expect, test, type Page } from '@playwright/test';
import { gotoApp } from './helpers';

async function run<T = unknown>(page: Page, body: string): Promise<T> {
  return page.evaluate(`(async () => { const api = window.grillon.api; const wb = api.getActiveWorkbook(); const ws = wb.getActiveSheet(); ${body} })()`) as Promise<T>;
}
const select = (page: Page, a1: string) => run(page, `wb.setActiveRange(ws.getRange('${a1}'));`);
const pageCount = async (page: Page) => {
  await run(page, `api.executeCommand('grillon.view.print');`);
  // Par défaut la boîte ajuste à la largeur de la page, ce qui ignore les sauts manuels.
  await page.getByTestId('dialog-print').getByLabel('Ajuster à la largeur de la page').uncheck();
  const label = page.getByTestId('dialog-print').getByTestId('print-page-count');
  await expect(label).toContainText('sur');
  const n = Number((await label.textContent())!.split('sur')[1]);
  await page.keyboard.press('Escape');
  await expect(page.getByTestId('dialog-print')).toHaveCount(0);
  return n;
};
const breaks = (page: Page) => run<{ rowBreaks?: number[]; colBreaks?: number[] }>(page, `return window.grillon.doc.snapshot().sheets[ws.getSheetId()].custom?.grillon?.page ?? {};`);

test.beforeEach(async ({ page }) => {
  await gotoApp(page);
  await run(page, `ws.getRange('A1:D120').setValues(Array.from({ length: 120 }, (_, i) => [i + 1, 'a', 'b', 'c']));`);
});

test('Insérer et supprimer un saut de page manuel', async ({ page }) => {
  const before = await pageCount(page);
  await select(page, 'C31');
  await run(page, `api.executeCommand('grillon.view.break-insert');`);
  expect(await breaks(page)).toMatchObject({ rowBreaks: [30], colBreaks: [2] });
  expect(await run<string>(page, `return window.grillon.doc.dirty ? 'dirty' : 'clean';`)).toBe('dirty');
  expect(await pageCount(page)).toBeGreaterThan(before);

  await run(page, `api.executeCommand('grillon.view.break-remove');`);
  expect(await breaks(page)).toEqual({});
  expect(await pageCount(page)).toBe(before);
});

test('Ligne entière : saut de ligne seulement ; Rétablir tous les sauts', async ({ page }) => {
  await run(page, `wb.setActiveRange(ws.getRange(40, 0, 1, ws.getMaxColumns()));`);
  await run(page, `api.executeCommand('grillon.view.break-insert');`);
  expect(await breaks(page)).toEqual({ rowBreaks: [40] });
  await select(page, 'B61');
  await run(page, `api.executeCommand('grillon.view.break-insert');`);
  expect(await breaks(page)).toEqual({ rowBreaks: [40, 60], colBreaks: [1] });
  await run(page, `api.executeCommand('grillon.view.break-reset');`);
  expect(await breaks(page)).toEqual({});
});

test('Aperçu des sauts de page : lignes manuelles et automatiques, pages numérotées', async ({ page }) => {
  await select(page, 'A21');
  await run(page, `api.executeCommand('grillon.view.break-insert');`);
  await run(page, `api.executeCommand('grillon.view.break-preview');`);
  const preview = page.getByTestId('page-break-preview');
  await expect(preview.getByTestId('pb-row-manual-20')).toBeVisible();
  await expect(preview.getByTestId('pb-page-1')).toBeVisible();
  await expect(preview.getByTestId('pb-page-1')).toHaveText('Page 1');
  await preview.getByTestId('pb-normal').click();
  await expect(preview).toHaveCount(0);
});

test('Imprimer : les sauts manuels sont ignorés avec « Ajuster à la largeur », et la note le dit', async ({ page }) => {
  await select(page, 'A31');
  await run(page, `api.executeCommand('grillon.view.break-insert');`);
  await run(page, `api.executeCommand('grillon.view.print');`);
  const dialog = page.getByTestId('dialog-print');
  await expect(dialog.getByTestId('print-fit-note')).toBeVisible();
  await dialog.getByLabel('Ajuster à la largeur de la page').uncheck();
  await expect(dialog.getByTestId('print-fit-note')).toBeHidden();
  // Page 1 : 30 lignes (jusqu'au saut), pas 40.
  await expect(dialog.locator('iframe').contentFrame().locator('tbody tr')).toHaveCount(30);
});

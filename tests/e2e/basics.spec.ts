import { expect, test, type Page } from '@playwright/test';
import { gotoApp } from './helpers';

// La toute première frappe après la prise de focus est parfois perdue en headless.
test.describe.configure({ retries: 2 });

async function run<T = unknown>(page: Page, body: string): Promise<T> {
  return page.evaluate(`(async () => { const api = window.grillon.api; const wb = api.getActiveWorkbook(); const ws = wb.getActiveSheet(); ${body} })()`) as Promise<T>;
}
const select = (page: Page, a1: string) => run(page, `wb.setActiveRange(ws.getRange('${a1}'));`);
const active = (page: Page) => run<string>(page, `return wb.getActiveRange().getA1Notation();`);
const display = (page: Page, a1: string) => run<string>(page, `return ws.getRange('${a1}').getDisplayValue();`);
const formula = (page: Page, a1: string) => run<string>(page, `return ws.getRange('${a1}').getFormula();`);

test.beforeEach(async ({ page }) => {
  await gotoApp(page);
  await run(page, `ws.getRange('A1:B3').setValues([[1,2],[3,4],[5,6]]);`);
  await page.locator('#univer').click({ position: { x: 600, y: 500 } });
  await page.waitForTimeout(400);
  await select(page, 'A1');
  await expect.poll(() => active(page)).toBe('A1');
});

test('Alt+= : somme automatique au-dessus, puis sur une sélection', async ({ page }) => {
  await select(page, 'A4');
  await page.keyboard.press('Alt+Equal');
  await expect.poll(() => formula(page, 'A4')).toBe('=SUM(A1:A3)');
  await expect.poll(() => display(page, 'A4')).toBe('9');
  await select(page, 'A1:B3');
  await page.keyboard.press('Alt+Equal');
  await expect.poll(() => formula(page, 'B4')).toBe('=SUM(B1:B3)');
  await expect.poll(() => display(page, 'B4')).toBe('12');
});

test('Σ du ruban : moyenne', async ({ page }) => {
  await select(page, 'A4');
  await page.getByRole('tab', { name: 'Formules', exact: true }).click();
  await page.locator('[data-u-command="grillon.autosum.formulas"]').first().click();
  await page.getByRole('menuitem', { name: 'Moyenne', exact: true }).click();
  await expect.poll(() => formula(page, 'A4')).toBe('=AVERAGE(A1:A3)');
  await expect.poll(() => display(page, 'A4')).toBe('3');
});

test('Remplir vers le haut et vers la gauche', async ({ page }) => {
  await select(page, 'A1:A3');
  await run(page, `await ws.getRange('A1:A3').getCellData(); ws.getRange('A3').setValue(9);`);
  await page.getByRole('tab', { name: 'Démarrer', exact: true }).click();
  await page.locator('[data-u-command="grillon.fill"]').first().click();
  await page.getByRole('menuitem', { name: 'Vers le haut', exact: true }).click();
  await expect.poll(() => display(page, 'A1')).toBe('9');
  await expect.poll(() => display(page, 'A2')).toBe('9');
  await select(page, 'A1:B1');
  await page.locator('[data-u-command="grillon.fill"]').first().click();
  await page.getByRole('menuitem', { name: 'Vers la gauche', exact: true }).click();
  await expect.poll(() => display(page, 'A1')).toBe('2');
});

test('Ctrl+` affiche les formules ; Ctrl+Maj+5 met en pourcentage', async ({ page }) => {
  await run(page, `ws.getRange('C1').setValue('=A1+B1');`);
  await expect.poll(() => display(page, 'C1')).toBe('3');
  await select(page, 'C1');
  await page.keyboard.press('Control+Backquote');
  await expect.poll(() => run<string>(page, `return JSON.stringify(ws.getRange('C1').getCellDataGrid()[0][0]);`)).toContain('A1+B1');
  await page.keyboard.press('Control+Backquote');
  await select(page, 'B2');
  await page.keyboard.press('Control+Shift+Digit5');
  await expect.poll(() => run<string>(page, `return ws.getRange('B2').getNumberFormat();`)).toBe('0%');
});

test('Ctrl+Espace, Maj+Espace, Ctrl+Fin, Ctrl+Origine', async ({ page }) => {
  await select(page, 'B2');
  await page.keyboard.press('Control+Space');
  await expect.poll(() => active(page)).toBe('B1:B1000');
  await select(page, 'B2');
  await page.keyboard.press('Shift+Space');
  await expect.poll(() => active(page)).toMatch(/^A2:[A-Z]+2$/);
  await select(page, 'A1');
  await page.keyboard.press('Control+End');
  await expect.poll(() => active(page)).toBe('B3');
  await page.keyboard.press('Control+Home');
  await expect.poll(() => active(page)).toBe('A1');
});

test('Ctrl+- : supprimer une ligne entière via la boîte', async ({ page }) => {
  await select(page, 'A2');
  await page.keyboard.press('Control+Minus');
  const dialog = page.getByTestId('dialog-cells');
  await expect(dialog).toBeVisible();
  await dialog.getByLabel('Ligne entière').check();
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect.poll(() => display(page, 'A2')).toBe('5');
});

import { expect, test, type Page } from '@playwright/test';
import { gotoApp } from './helpers';

async function run<T = unknown>(page: Page, body: string): Promise<T> {
  return page.evaluate(`(async () => { const api = window.grillon.api; const wb = api.getActiveWorkbook(); const ws = wb.getActiveSheet(); ${body} })()`) as Promise<T>;
}
const formulaOf = (page: Page, a1: string) => run<string>(page, `return ws.getRange('${a1}').getFormula();`);
const shown = (page: Page, a1: string) => run<string>(page, `return ws.getRange('${a1}').getDisplayValue();`);

test.beforeEach(async ({ page }) => {
  await gotoApp(page);
  await run(page, `ws.getRange('A1:A3').setValues([[1], [2], [3]]);`);
  // Donne le focus à la grille (la toute première frappe peut sinon être perdue).
  const box = (await page.locator('canvas[id^="univer-sheet-main-canvas"]').boundingBox())!;
  await page.mouse.click(box.x + 600, box.y + 400);
  await page.waitForTimeout(400);
});

async function typeIn(page: Page, a1: string, text: string) {
  await run(page, `wb.setActiveRange(ws.getRange('${a1}'));`);
  await page.keyboard.press('F2');
  await page.waitForTimeout(250);
  await page.keyboard.type(text, { delay: 25 });
  await page.keyboard.press('Enter');
}

test('une formule saisie en français est enregistrée en anglais et calculée', async ({ page }) => {
  await typeIn(page, 'C1', '=SOMME(A1:A3;10)');
  await expect.poll(() => formulaOf(page, 'C1')).toBe('=SUM(A1:A3,10)');
  await expect.poll(() => shown(page, 'C1')).toBe('16');
});

test('virgule décimale, SI, VRAI et RECHERCHEV', async ({ page }) => {
  await typeIn(page, 'C2', '=SI(SOMME(A1:A3)>5;ARRONDI(2,5;0);"non")');
  await expect.poll(() => formulaOf(page, 'C2')).toBe('=IF(SUM(A1:A3)>5,ROUND(2.5,0),"non")');
  await expect.poll(() => shown(page, 'C2')).toBe('3');
  await typeIn(page, 'C3', '=RECHERCHEV(2;A1:A3;1;FAUX)');
  await expect.poll(() => formulaOf(page, 'C3')).toBe('=VLOOKUP(2,A1:A3,1,FALSE)');
  await expect.poll(() => shown(page, 'C3')).toBe('2');
});

test('une formule saisie en anglais fonctionne toujours', async ({ page }) => {
  await typeIn(page, 'C1', '=SUM(A1:A3,10)');
  await expect.poll(() => formulaOf(page, 'C1')).toBe('=SUM(A1:A3,10)');
  await expect.poll(() => shown(page, 'C1')).toBe('16');
});

test('ouvrir une cellule puis valider sans rien changer laisse la formule intacte', async ({ page }) => {
  const f = '=IF(SUM(A1:A3)>5,ROUND(2.5,0),"non")';
  await run(page, `ws.getRange('C1').setFormula('${f}');`);
  await run(page, `wb.setActiveRange(ws.getRange('C1'));`);
  await page.keyboard.press('F2');
  await page.waitForTimeout(300);
  await page.keyboard.press('Enter');
  await page.waitForTimeout(500);
  expect(await formulaOf(page, 'C1')).toBe(f);
  await expect.poll(() => shown(page, 'C1')).toBe('3');
});

test('autocomplétion : les noms de fonctions sont proposés en français', async ({ page }) => {
  await run(page, `wb.setActiveRange(ws.getRange('B3'));`);
  await page.keyboard.press('F2');
  await page.waitForTimeout(250);
  await page.keyboard.type('=SOM', { delay: 60 });
  await expect(page.getByText('SOMME.SI', { exact: true }).first()).toBeVisible();
  await expect(page.getByText('SUMIF', { exact: true })).toHaveCount(0);
});

test('la fenêtre Espion montre la formule en français', async ({ page }) => {
  await run(page, `ws.getRange('C1').setFormula('=IF(SUM(A1:A3)>5,0.5,"non")'); wb.setActiveRange(ws.getRange('C1')); api.executeCommand('grillon.formulas.watch');`);
  const dialog = page.getByTestId('dialog-watch');
  await dialog.getByTestId('watch-add').click();
  await expect(dialog.getByTestId('watch-row')).toContainText('=SI(SOMME(A1:A3)>5;0,5;"non")');
});

// Univer traitait la barre de formule comme une saisie de références : une flèche changeait « A1:A3 » en « Y1000 ».
test('barre de formule : les flèches déplacent le curseur dans le texte, pas la plage', async ({ page }) => {
  await run(page, `ws.getRange('C1').setFormula('=SUM(A1:A3)+1'); wb.setActiveRange(ws.getRange('C1'));`);
  await page.waitForTimeout(500);
  // Clic dans la barre de formule, à droite du texte : le curseur est à la fin.
  await page.mouse.click(600, 155);
  await page.waitForTimeout(400);
  await page.keyboard.press('ArrowLeft');
  await page.waitForTimeout(150);
  await page.keyboard.type('2', { delay: 40 });
  await page.keyboard.press('Enter');
  await expect.poll(() => formulaOf(page, 'C1')).toBe('=SUM(A1:A3)+21');
  await expect.poll(() => shown(page, 'C1')).toBe('27');
});

test('modification avec F2 : les flèches déplacent aussi le curseur', async ({ page }) => {
  await run(page, `ws.getRange('C1').setFormula('=SUM(A1:A3)+1'); wb.setActiveRange(ws.getRange('C1'));`);
  await page.waitForTimeout(400);
  await page.keyboard.press('F2');
  await page.waitForTimeout(300);
  await page.keyboard.press('ArrowLeft');
  await page.keyboard.type('2', { delay: 40 });
  await page.keyboard.press('Enter');
  await expect.poll(() => formulaOf(page, 'C1')).toBe('=SUM(A1:A3)+21');
});

import { expect, test, type Page } from '@playwright/test';
import { gotoApp } from './helpers';

async function run<T = unknown>(page: Page, body: string): Promise<T> {
  return page.evaluate(`(async () => { const api = window.grillon.api; const wb = api.getActiveWorkbook(); const ws = wb.getActiveSheet(); ${body} })()`) as Promise<T>;
}

test.beforeEach(async ({ page }) => {
  await gotoApp(page);
  await run(page, `ws.getRange('F1:F3').setValues([[2], [3], [4]]);
    ws.getRange('H1').setFormula('=SUM(F1:F3)*2+IF(F1>1,MAX(F2,F3),0)');
    ws.getRange('H2').setFormula('=1/0');
    ws.getRange('H3').setFormula('=FOO(1)');
    ws.getRange('H4').setValue({ v: '123', t: 1 });`);
});

test('Évaluer la formule : pas à pas jusqu’au résultat', async ({ page }) => {
  await run(page, `wb.setActiveRange(ws.getRange('H1'));`);
  await run(page, `api.executeCommand('grillon.formulas.evaluate');`);
  const dialog = page.getByTestId('dialog-evaluate');
  const expr = dialog.getByTestId('evaluate-expression');
  await expect(expr.locator('u')).toHaveText('F1:F3');
  await dialog.getByTestId('evaluate-step').click();
  await expect(expr.locator('u')).toHaveText('SUM({2;3;4})');
  for (let i = 0; i < 12 && (await expr.textContent()) !== '=22'; i++) {
    const before = await expr.textContent();
    await expect(dialog.getByTestId('evaluate-step')).toBeEnabled();
    await dialog.getByTestId('evaluate-step').click();
    await expect(expr).not.toHaveText(before!);
  }
  await expect(expr).toHaveText('=22');
  await expect(dialog.getByTestId('evaluate-step')).toBeDisabled();
});

test('Vérification des erreurs : erreurs en français, conversion d’un nombre stocké en texte', async ({ page }) => {
  await expect.poll(() => run(page, `return ws.getRange('H3').getDisplayValue();`)).toBe('#NOM?');
  await run(page, `wb.setActiveRange(ws.getRange('A1')); api.executeCommand('grillon.formulas.error-checking');`);
  const dialog = page.getByTestId('dialog-error-checking');
  await expect(dialog.getByTestId('error-cell')).toHaveText('Erreur dans la cellule H2');
  await expect(dialog.getByTestId('error-title')).toContainText('division par zéro');
  await dialog.getByTestId('error-next').click();
  await expect(dialog.getByTestId('error-title')).toContainText('#NOM?');
  await dialog.getByTestId('error-next').click();
  await expect(dialog.getByTestId('error-title')).toHaveText('Nombre stocké sous forme de texte');
  await dialog.getByTestId('error-convert').click();
  await expect.poll(() => run(page, `return ws.getSheet().getCellRaw(3, 7)?.v;`)).toBe(123);
  await dialog.getByTestId('error-ignore').click();
  await dialog.getByTestId('error-ignore').click();
  await expect(page.getByRole('dialog').filter({ hasText: 'La vérification des erreurs est terminée' })).toBeVisible();
});

test('Options de calcul : manuel puis F9, comme Excel', async ({ page }) => {
  await run(page, `ws.getRange('M1').setValue(5); ws.getRange('M2').setFormula('=M1*2');`);
  await expect.poll(() => run(page, `return ws.getRange('M2').getDisplayValue();`)).toBe('10');
  await run(page, `await api.executeCommand('grillon.formulas.calc-manual');`);
  await expect(page.getByTestId('calc-manual')).toBeVisible();
  await run(page, `ws.getRange('M1').setValue(7);`);
  await page.waitForTimeout(800);
  expect(await run(page, `return ws.getRange('M2').getDisplayValue();`)).toBe('10');
  await page.locator('#univer').click({ position: { x: 700, y: 500 } });
  await page.keyboard.press('F9');
  await expect.poll(() => run(page, `return ws.getRange('M2').getDisplayValue();`)).toBe('14');
  await run(page, `await api.executeCommand('grillon.formulas.calc-auto');`);
  await expect(page.getByTestId('calc-manual')).toHaveCount(0);
  await run(page, `ws.getRange('M1').setValue(1);`);
  await expect.poll(() => run(page, `return ws.getRange('M2').getDisplayValue();`)).toBe('2');
});

test('Imprimer : aperçu des pages à côté des réglages', async ({ page }) => {
  await run(page, `ws.getRange('A1:B150').setValues(Array.from({ length: 150 }, (_, i) => ['Ligne ' + (i + 1), i]));`);
  await page.evaluate(() => void (window as any).grillon.doc.print());
  const dialog = page.getByTestId('dialog-print');
  const count = dialog.getByTestId('print-page-count');
  await expect(count).toHaveText(/^Page 1 sur [2-9]$/);
  const before = Number((await count.textContent())!.match(/sur (\d+)/)![1]);
  await dialog.getByRole('button', { name: 'Page suivante' }).click();
  await expect(count).toHaveText(`Page 2 sur ${before}`);
  await dialog.getByLabel('Orientation :').selectOption('landscape');
  await expect.poll(async () => Number((await count.textContent())!.match(/sur (\d+)/)![1])).toBeGreaterThan(before);
  await dialog.getByRole('button', { name: 'Annuler' }).click();
});

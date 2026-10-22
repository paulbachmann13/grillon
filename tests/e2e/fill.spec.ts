import { expect, test, type Page } from '@playwright/test';
import { gotoApp } from './helpers';

async function run<T = unknown>(page: Page, body: string): Promise<T> {
  return page.evaluate(`(async () => { const api = window.grillon.api; const wb = api.getActiveWorkbook(); const ws = wb.getActiveSheet(); ${body} })()`) as Promise<T>;
}

// Une formule étendue (poignée de recopie, Remplir) est « partagée » dans Univer : seule la première cellule garde le texte
// de la formule. Elle doit quand même être écrite dans chaque cellule du .xlsx (sinon la réouverture ne montre que des nombres).
test('formule étendue vers le bas et vers la droite : enregistrée, puis rouverte avec ses formules', async ({ page }) => {
  await gotoApp(page);
  await run(page, `
    ws.getRange('A1:A6').setValues([[1], [2], [3], [4], [5], [6]]);
    ws.getRange('B1:C1').setValues([[10], [20]]);
    ws.getRange('B2').setFormula('=A2*$B$1');
    ws.getRange('B2').autoFill(ws.getRange('B2:B6'), 'COPY');
    ws.getRange('D2').setFormula('=SUM(A2:B2)');
    ws.getRange('D2').autoFill(ws.getRange('D2:F2'), 'COPY');
  `);
  await expect.poll(() => run<string>(page, `return ws.getRange('B6').getDisplayValue();`)).toBe('60');

  // Enregistre (en mémoire) puis rouvre comme un fichier.
  await page.evaluate(`window.grillon.doc.exportXlsxBytes().then((bytes) => window.__saved = bytes)`);
  await page.evaluate(`window.grillon.doc.openBytes('rouvert.xlsx', window.__saved)`);
  await expect(page.getByTestId('doc-title')).toHaveText(/^rouvert/);
  await expect.poll(() => run<string>(page, `return ws.getRange('B6').getDisplayValue();`)).toBe('60');

  const formulas = await run<string[]>(page, `return ['B2', 'B3', 'B4', 'B5', 'B6', 'D2', 'E2', 'F2'].map((a) => ws.getRange(a).getFormula());`);
  expect(formulas).toEqual(['=A2*$B$1', '=A3*$B$1', '=A4*$B$1', '=A5*$B$1', '=A6*$B$1', '=SUM(A2:B2)', '=SUM(B2:C2)', '=SUM(C2:D2)']);
  // Et elles se recalculent : changer la source change le résultat.
  await run(page, `ws.getRange('A5').setValue(100);`);
  await expect.poll(() => run<string>(page, `return ws.getRange('B5').getDisplayValue();`)).toBe('1000');
});

test('copier-coller de formules : elles sont aussi enregistrées', async ({ page }) => {
  await gotoApp(page);
  await run(page, `
    ws.getRange('A1:A4').setValues([[1], [2], [3], [4]]);
    ws.getRange('B1').setFormula('=A1+100');
    ws.getRange('B1').autoFill(ws.getRange('B1:B4'), 'COPY');
    ws.getRange('C1:C4').setValues([[0], [0], [0], [0]]);
  `);
  await page.waitForTimeout(500);
  await page.evaluate(`window.grillon.doc.exportXlsxBytes().then((bytes) => window.__saved = bytes)`);
  await page.evaluate(`window.grillon.doc.openBytes('copie.xlsx', window.__saved)`);
  await expect(page.getByTestId('doc-title')).toHaveText(/^copie/);
  const formulas = await run<string[]>(page, `return ['B1', 'B2', 'B3', 'B4'].map((a) => ws.getRange(a).getFormula());`);
  expect(formulas).toEqual(['=A1+100', '=A2+100', '=A3+100', '=A4+100']);
});

test('« Afficher les formules » montre aussi la formule des cellules étendues', async ({ page }) => {
  await gotoApp(page);
  await run(page, `
    ws.getRange('A1:A4').setValues([[1], [2], [3], [4]]);
    ws.getRange('B1').setFormula('=A1+100');
    ws.getRange('B1').autoFill(ws.getRange('B1:B4'), 'COPY');
  `);
  await page.waitForTimeout(400);
  await run(page, `api.executeCommand('grillon.formulas.show.formulas');`);
  await expect.poll(() => run<string>(page, `return ws.getRange('B4').getDisplayValue();`)).toBe('=A4+100');
  expect(await run<string>(page, `return ws.getRange('B2').getDisplayValue();`)).toBe('=A2+100');
});

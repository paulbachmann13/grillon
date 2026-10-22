import { expect, test, type Page } from '@playwright/test';
import { gotoApp } from './helpers';

async function run<T = unknown>(page: Page, body: string): Promise<T> {
  return page.evaluate(`(async () => { const api = window.grillon.api; const wb = api.getActiveWorkbook(); const ws = wb.getActiveSheet(); ${body} })()`) as Promise<T>;
}
const rows = (page: Page) => run<string[]>(page, `return [0, 1, 2, 3].map((r) => ws.getRange(r, 0).getDisplayValue() + ' | ' + (ws.getRange(r, 1).getFormula() || '') + ' = ' + ws.getRange(r, 1).getDisplayValue());`);
const sortDesc = (page: Page) =>
  run(page, `await api.executeCommand('sheet.command.sort-range', { unitId: wb.getId(), subUnitId: ws.getSheetId(), range: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 }, orderRules: [{ colIndex: 0, type: 'desc' }], hasTitle: false });`);

// Univer déplaçait les lignes sans adapter les formules : après tri, « 8 » gardait « =A3*10 » et affichait 30.
test('tri : chaque formule suit sa ligne, comme dans Excel', async ({ page }) => {
  await gotoApp(page);
  await run(page, `ws.getRange('A1:A4').setValues([[5], [3], [8], [1]]); for (let r = 1; r <= 4; r++) ws.getRange('B' + r).setFormula('=A' + r + '*10');`);
  await sortDesc(page);
  await expect.poll(() => rows(page)).toEqual(['8 | =A1*10 = 80', '5 | =A2*10 = 50', '3 | =A3*10 = 30', '1 | =A4*10 = 10']);
});

test('tri de formules étendues, puis annulation en une fois', async ({ page }) => {
  await gotoApp(page);
  await run(page, `ws.getRange('A1:A4').setValues([[5], [3], [8], [1]]); ws.getRange('B1').setFormula('=A1*10'); ws.getRange('B1').autoFill(ws.getRange('B1:B4'), 'COPY');`);
  await expect.poll(() => rows(page)).toEqual(['5 | =A1*10 = 50', '3 | =A2*10 = 30', '8 | =A3*10 = 80', '1 | =A4*10 = 10']);
  await sortDesc(page);
  await expect.poll(() => rows(page)).toEqual(['8 | =A1*10 = 80', '5 | =A2*10 = 50', '3 | =A3*10 = 30', '1 | =A4*10 = 10']);
  // Enregistré puis rouvert : les formules sont bien là.
  await page.evaluate(`window.grillon.doc.exportXlsxBytes().then((b) => window.__b = b)`);
  await run(page, `api.executeCommand('univer.command.undo');`);
  await expect.poll(() => rows(page)).toEqual(['5 | =A1*10 = 50', '3 | =A2*10 = 30', '8 | =A3*10 = 80', '1 | =A4*10 = 10']);
  await page.evaluate(`window.grillon.doc.openBytes('trie.xlsx', window.__b)`);
  await expect(page.getByTestId('doc-title')).toHaveText(/^trie/);
  await expect.poll(() => rows(page)).toEqual(['8 | =A1*10 = 80', '5 | =A2*10 = 50', '3 | =A3*10 = 30', '1 | =A4*10 = 10']);
});

test('tri d’une partie d’une colonne de formules étendues : les cellules hors du tri gardent leur formule', async ({ page }) => {
  await gotoApp(page);
  await run(page, `ws.getRange('A1:A6').setValues([[5], [3], [8], [1], [6], [2]]); ws.getRange('B1').setFormula('=A1*10'); ws.getRange('B1').autoFill(ws.getRange('B1:B6'), 'COPY');`);
  await page.waitForTimeout(400);
  await sortDesc(page);
  await expect.poll(() => run<string[]>(page, `return ['B5', 'B6'].map((a) => ws.getRange(a).getFormula() + ' = ' + ws.getRange(a).getDisplayValue());`)).toEqual(['=A5*10 = 60', '=A6*10 = 20']);
});

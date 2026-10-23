import { expect, test, type Page } from '@playwright/test';
import { gotoApp } from './helpers';

async function run<T = unknown>(page: Page, body: string): Promise<T> {
  return page.evaluate(`(async () => { const api = window.grillon.api; const wb = api.getActiveWorkbook(); const ws = wb.getActiveSheet(); ${body} })()`) as Promise<T>;
}

// Saisie au clavier « à la française », comme Excel en français. [ce qu'on tape, valeur rangée, type, affichage]
const CASES: [string, unknown, number, string][] = [
  ['1,5', 1.5, 2, '1,5'],
  ['1 234,5', 1234.5, 2, '1 234,5'],
  ['1 000 000', 1000000, 2, '1 000 000'],
  ['-1 234', -1234, 2, '-1 234'],
  ['-12,50 €', -12.5, 2, '-12,50 €'],
  ['1.5', 1.5, 2, '1,5'],
  ['15%', 0.15, 2, '15%'],
  ['12/03/2024', 45363, 2, '12/03/2024'],
  ['31/12', null, 2, '31/12'],
  ['VRAI', 1, 3, 'VRAI'],
  ['faux', 0, 3, 'FAUX'],
  ["'1.5", '1.5', 4, '1.5'],
];

test('saisie au clavier : nombres, montants, dates et VRAI / FAUX à la française, conservés après réouverture', async ({ page }) => {
  await gotoApp(page);
  const box = (await page.locator('canvas[id^="univer-sheet-main-canvas"]').boundingBox())!;
  await page.mouse.click(box.x + 600, box.y + 400);
  await page.waitForTimeout(400);
  for (let i = 0; i < CASES.length; i++) {
    await run(page, `wb.setActiveRange(ws.getRange(${i}, 0));`);
    await page.keyboard.type(CASES[i][0], { delay: 15 });
    await page.keyboard.press('Enter');
    await page.waitForTimeout(80);
  }
  const read = () => run<[unknown, number, string][]>(page, `return Array.from({ length: ${CASES.length} }, (_, i) => { const c = ws.getSheet().getCellRaw(i, 0) || {}; return [c.v, c.t, ws.getRange(i, 0).getDisplayValue().replace(/[\\u00a0\\u202f]/g, ' ')]; });`);
  const year = new Date().getFullYear();
  const expected = CASES.map(([text, v, t, shown]) => [text === '31/12' ? Math.round((Date.UTC(year, 11, 31) - Date.UTC(1899, 11, 30)) / 86400000) : v, t, shown]);
  await expect.poll(read).toEqual(expected);
  // Une somme les compte bien comme des nombres.
  await run(page, `ws.getRange('C1').setFormula('=SUM(A1:A6)');`);
  await expect.poll(() => run<string>(page, `return Number(ws.getRange('C1').getValue());`)).toBe(1.5 + 1234.5 + 1000000 - 1234 - 12.5 + 1.5);
  // Enregistré puis rouvert : mêmes affichages (le texte forcé reste du texte).
  await page.evaluate(`window.grillon.doc.exportXlsxBytes().then((b) => window.__b = b)`);
  await page.evaluate(`window.grillon.doc.openBytes('saisie.xlsx', window.__b)`);
  await expect(page.getByTestId('doc-title')).toHaveText(/^saisie/);
  await expect.poll(async () => (await read()).map((x) => x[2])).toEqual(CASES.map((c) => c[3]));
});

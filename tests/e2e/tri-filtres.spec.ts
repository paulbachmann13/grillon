import { expect, test, type Page } from '@playwright/test';
import { gotoApp, waitSortFilterReady } from './helpers';

async function run<T = unknown>(page: Page, body: string): Promise<T> {
  return page.evaluate(`(async () => { const api = window.grillon.api; const wb = api.getActiveWorkbook(); const ws = wb.getActiveSheet(); ${body} })()`) as Promise<T>;
}
const table = (page: Page, ref: string) => run<string[]>(page, `return ws.getRange('${ref}').getDisplayValues().map((r) => r.join(' / '));`);
const visibleNames = (page: Page, rows: number) => run<string[]>(page, `const s = ws.getSheet(); return Array.from({ length: ${rows} }, (_, r) => r).filter((r) => s.getRowVisible(r)).map((r) => ws.getRange(r, 0).getDisplayValue());`);

// Liste type Parcoursup : noms accentués, ex-aequo, une moyenne vide, en-tête en gras.
const ELEVES = `[['Nom', 'Prénom', 'Moyenne'], ['Zoé', 'Lina', 14.5], ['Émile', 'Paul', 12], ['adam', 'Léa', 16], ['Éloïse', 'Marc', 9.25], ['Ethan', 'Noé', 12], ['Bérénice', 'Inès', null], ['Jean-Marc', 'Tom', 18]]`;

async function eleves(page: Page) {
  await gotoApp(page);
  await waitSortFilterReady(page);
  await run(page, `ws.getRange('A1:C8').setValues(${ELEVES}); ws.getRange('A1:C1').setFontWeight('bold');`);
}

test.describe('tri comme Excel', () => {
  test('une cellule sélectionnée, A→Z : tout le tableau, en-tête en haut, ordre alphabétique français', async ({ page }) => {
    await eleves(page);
    await run(page, `wb.setActiveRange(ws.getRange('A3')); await api.executeCommand('sheet.command.sort-range-asc');`);
    await expect.poll(() => table(page, 'A1:C8')).toEqual([
      'Nom / Prénom / Moyenne',
      'adam / Léa / 16',
      'Bérénice / Inès / ',
      'Éloïse / Marc / 9,25',
      'Émile / Paul / 12',
      'Ethan / Noé / 12',
      'Jean-Marc / Tom / 18',
      'Zoé / Lina / 14,5',
    ]);
  });

  test('une colonne sélectionnée, Z→A : les lignes restent entières (les notes suivent les élèves), vides en dernier, ex-aequo dans l’ordre', async ({ page }) => {
    await eleves(page);
    await run(page, `wb.setActiveRange(ws.getRange('C1:C8')); await api.executeCommand('sheet.command.sort-range-desc');`);
    await expect.poll(() => table(page, 'A1:C8')).toEqual([
      'Nom / Prénom / Moyenne',
      'Jean-Marc / Tom / 18',
      'adam / Léa / 16',
      'Zoé / Lina / 14,5',
      'Émile / Paul / 12',
      'Ethan / Noé / 12',
      'Éloïse / Marc / 9,25',
      'Bérénice / Inès / ',
    ]);
  });

  test('menu contextuel (Trier) : même comportement', async ({ page }) => {
    await eleves(page);
    await run(page, `wb.setActiveRange(ws.getRange('C4')); await api.executeCommand('sheet.command.sort-range-desc-ctx');`);
    await expect.poll(async () => (await table(page, 'A1:A3'))).toEqual(['Nom', 'Jean-Marc', 'adam']);
  });

  test('colonne Rang en formule : les rangs restent justes après le tri', async ({ page }) => {
    await eleves(page);
    await run(page, `ws.getRange('D1').setValue('Rang'); ws.getRange('D2').setFormula('=RANK(C2,$C$2:$C$8)'); ws.getRange('D2').autoFill(ws.getRange('D2:D8'), 'COPY');`);
    await run(page, `wb.setActiveRange(ws.getRange('C2')); await api.executeCommand('sheet.command.sort-range-desc');`);
    await expect.poll(() => table(page, 'A2:D5')).toEqual(['Jean-Marc / Tom / 18 / 1', 'adam / Léa / 16 / 2', 'Zoé / Lina / 14,5 / 3', 'Émile / Paul / 12 / 4']);
  });

  test('liste sans en-tête (que des noms, sans mise en forme) : la première ligne est triée aussi', async ({ page }) => {
    await gotoApp(page);
    await waitSortFilterReady(page);
    await run(page, `ws.getRange('A1:A4').setValues([['Zoé'], ['Émile'], ['adam'], ['Ethan']]); wb.setActiveRange(ws.getRange('A1')); await api.executeCommand('sheet.command.sort-range-asc');`);
    await expect.poll(() => table(page, 'A1:A4')).toEqual(['adam', 'Émile', 'Ethan', 'Zoé']);
  });
});

test('tri personnalisé à deux clés (Spécialité A→Z puis Moyenne Z→A), avec en-tête', async ({ page }) => {
  await gotoApp(page);
  await waitSortFilterReady(page);
  await run(page, `ws.getRange('A1:C6').setValues([['Nom', 'Spécialité', 'Moyenne'], ['Zoé', 'Physique', 14], ['Émile', 'maths', 12], ['adam', 'Maths', 16], ['Éloïse', 'Économie', 9], ['Ethan', 'Physique', 15]]);`);
  // C'est ce qu'envoie la boîte « Tri personnalisé ».
  await run(page, `await api.executeCommand('sheet.command.sort-range', { unitId: wb.getId(), subUnitId: ws.getSheetId(), range: { startRow: 0, endRow: 5, startColumn: 0, endColumn: 2 }, orderRules: [{ colIndex: 1, type: 'asc' }, { colIndex: 2, type: 'desc' }], hasTitle: true });`);
  await expect.poll(() => table(page, 'A1:C6')).toEqual(['Nom / Spécialité / Moyenne', 'Éloïse / Économie / 9', 'adam / Maths / 16', 'Émile / maths / 12', 'Ethan / Physique / 15', 'Zoé / Physique / 14']);
});

test.describe('filtres', () => {
  const SPE = `[['Nom', 'Spécialité', 'Moyenne'], ['Zoé', 'Maths', 14.5], ['Émile', 'Physique', 12], ['adam', 'Maths', 16], ['Éloïse', 'SES', 9.25], ['Ethan', 'Maths', 12], ['Bérénice', 'Physique', null], ['Jean-Marc', 'SES', 18], ['Inès', 'Maths', 8], ['Noé', 'Physique', 100]]`;
  async function withFilter(page: Page) {
    await gotoApp(page);
    await waitSortFilterReady(page);
    await run(page, `ws.getRange('A1:C10').setValues(${SPE}); ws.getRange('A1:C1').setFontWeight('bold'); ws.getRange('E1').setFormula('=SUBTOTAL(109,C2:C10)'); wb.setActiveRange(ws.getRange('B3'));`);
    await run(page, `await api.executeCommand('sheet.command.smart-toggle-filter');`);
    await expect.poll(() => run<string>(page, `return ws.getFilter()?.getRange().getA1Notation();`)).toBe('A1:C10');
  }

  test('par valeurs, puis tri pendant le filtre, puis enregistrement et réouverture', async ({ page }) => {
    await withFilter(page);
    await run(page, `ws.getFilter().setColumnFilterCriteria(1, { colId: 1, filters: { filters: ['Maths'] } });`);
    await expect.poll(() => visibleNames(page, 10)).toEqual(['Nom', 'Zoé', 'adam', 'Ethan', 'Inès']);
    await expect.poll(() => run<string>(page, `return ws.getRange('E1').getDisplayValue();`)).toBe('50,5');
    await run(page, `wb.setActiveRange(ws.getRange('C2')); await api.executeCommand('sheet.command.sort-range-desc');`);
    await expect.poll(() => visibleNames(page, 10)).toEqual(['Nom', 'adam', 'Zoé', 'Ethan', 'Inès']);
    await page.evaluate(`window.grillon.doc.exportXlsxBytes().then((b) => window.__b = b)`);
    await page.evaluate(`window.grillon.doc.openBytes('filtre.xlsx', window.__b)`);
    await expect(page.getByTestId('doc-title')).toHaveText(/^filtre/);
    await expect.poll(() => visibleNames(page, 10)).toEqual(['Nom', 'adam', 'Zoé', 'Ethan', 'Inès']);
  });

  test('condition numérique (> 12) : 100 est bien plus grand que 12', async ({ page }) => {
    await withFilter(page);
    await run(page, `ws.getFilter().setColumnFilterCriteria(2, { colId: 2, customFilters: { customFilters: [{ operator: 'greaterThan', val: '12' }] } });`);
    await expect.poll(() => visibleNames(page, 10)).toEqual(['Nom', 'Zoé', 'adam', 'Jean-Marc', 'Noé']);
  });

  test('cellules vides, texte « contient », deux filtres combinés, puis effacer', async ({ page }) => {
    await withFilter(page);
    await run(page, `ws.getFilter().setColumnFilterCriteria(2, { colId: 2, filters: { blank: true } });`);
    await expect.poll(() => visibleNames(page, 10)).toEqual(['Nom', 'Bérénice']);
    await run(page, `ws.getFilter().removeColumnFilterCriteria(2); ws.getFilter().setColumnFilterCriteria(0, { colId: 0, customFilters: { customFilters: [{ val: '*an*' }] } });`);
    await expect.poll(() => visibleNames(page, 10)).toEqual(['Nom', 'Ethan', 'Jean-Marc']);
    await run(page, `ws.getFilter().removeColumnFilterCriteria(0); ws.getFilter().setColumnFilterCriteria(1, { colId: 1, filters: { filters: ['Maths'] } }); ws.getFilter().setColumnFilterCriteria(2, { colId: 2, customFilters: { customFilters: [{ operator: 'greaterThanOrEqual', val: '12' }] } });`);
    await expect.poll(() => visibleNames(page, 10)).toEqual(['Nom', 'Zoé', 'adam', 'Ethan']);
    await run(page, `[0, 1, 2].forEach((c) => ws.getFilter().removeColumnFilterCriteria(c));`);
    await expect.poll(async () => (await visibleNames(page, 10)).length).toBe(10);
  });

  test('un élève ajouté sous la liste est pris dans le filtre', async ({ page }) => {
    await withFilter(page);
    await run(page, `ws.getRange('A11:C11').setValues([['Zara', 'Maths', 11]]); ws.getFilter().setColumnFilterCriteria(1, { colId: 1, filters: { filters: ['Maths'] } });`);
    await expect.poll(() => visibleNames(page, 11)).toEqual(['Nom', 'Zoé', 'adam', 'Ethan', 'Inès', 'Zara']);
  });
});

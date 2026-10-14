import { expect, test, type Page } from '@playwright/test';
import ExcelJS from 'exceljs';
import { display, gotoApp } from './helpers';

async function run<T = unknown>(page: Page, body: string): Promise<T> {
  return page.evaluate(`(async () => { const api = window.grillon.api; const wb = api.getActiveWorkbook(); const ws = wb.getActiveSheet(); ${body} })()`) as Promise<T>;
}

async function select(page: Page, a1: string): Promise<void> {
  await run(page, `wb.setActiveRange(ws.getRange('${a1}'));`);
}

async function dataTab(page: Page): Promise<void> {
  await page.getByRole('tab', { name: 'Données', exact: true }).click();
}

async function ribbon(page: Page, id: string, item?: string): Promise<void> {
  await page.locator(`[data-u-command="${id}"]`).first().click();
  if (item) await page.getByRole('menuitem', { name: item, exact: true }).click();
}

async function values(page: Page, a1: string): Promise<string[][]> {
  return run(page, `return ws.getRange('${a1}').getDisplayValues();`);
}

test.beforeEach(async ({ page }) => {
  await gotoApp(page);
});

test('Supprimer les doublons', async ({ page }) => {
  await run(page, `ws.getRange('A1:B6').setValues([['Nom','Ville'],['Alice','Lyon'],['alice','lyon'],['Bob','Paris'],['Alice','Paris'],['Bob','Paris']]);`);
  await select(page, 'A1');
  await dataTab(page);
  await ribbon(page, 'grillon.data.remove-duplicates');
  const dialog = page.getByTestId('dialog-remove-duplicates');
  await expect(dialog.getByLabel('Mes données ont des en-têtes')).toBeChecked();
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(page.getByText('2 valeur(s) en double trouvée(s) et supprimée(s) ; 3 valeur(s) unique(s) restante(s).')).toBeVisible();
  await page.getByRole('button', { name: 'OK' }).click();
  expect(await values(page, 'A1:B6')).toEqual([['Nom', 'Ville'], ['Alice', 'Lyon'], ['Bob', 'Paris'], ['Alice', 'Paris'], ['', ''], ['', '']]);
});

test('Convertir : texte en colonnes avec point-virgule', async ({ page }) => {
  await run(page, `ws.getRange('A1:A3').setValues([['Lyon;1,5;14/10/2026'],['Paris;2;01/01/2026'],['Nîmes;3,25;31/12/2025']]);`);
  await select(page, 'A1:A3');
  await dataTab(page);
  await ribbon(page, 'grillon.data.text-to-columns');
  await page.getByTestId('dialog-text-to-columns').getByRole('button', { name: 'OK' }).click();
  await expect.poll(() => values(page, 'A1:C3')).toEqual([
    ['Lyon', '1,5', '14/10/2026'],
    ['Paris', '2', '01/01/2026'],
    ['Nîmes', '3,25', '31/12/2025'],
  ]);
});

test('Remplissage instantané avec Ctrl+E', async ({ page }) => {
  await run(page, `ws.getRange('A1:B5').setValues([['Nom complet','Prénom'],['Jean Dupont','Jean'],['Marie Curie',''],['Ada Lovelace',''],['Alan Turing','']]);`);
  // Donne le focus à la grille (clic dans une cellule vide), puis sélectionne B3.
  await page.locator('#univer').click({ position: { x: 600, y: 500 } });
  await select(page, 'B3');
  await page.keyboard.press('Control+E');
  await expect.poll(() => values(page, 'B2:B5')).toEqual([['Jean'], ['Marie'], ['Ada'], ['Alan']]);
});

test('Sous-total et niveaux de plan', async ({ page }) => {
  await run(page, `ws.getRange('A1:B5').setValues([['Ville','Montant'],['Lyon',10],['Lyon',20],['Paris',5],['Paris',7]]);`);
  await select(page, 'A1');
  await dataTab(page);
  await ribbon(page, 'grillon.data.subtotal');
  await page.getByTestId('dialog-subtotal').getByRole('button', { name: 'OK' }).click();
  await expect.poll(() => values(page, 'A1:B8')).toEqual([
    ['Ville', 'Montant'],
    ['Lyon', '10'],
    ['Lyon', '20'],
    ['Total Lyon', '30'],
    ['Paris', '5'],
    ['Paris', '7'],
    ['Total Paris', '12'],
    ['Total général', '42'],
  ]);
  // Les niveaux de plan sont enregistrés dans le .xlsx (outlineLevel).
  const bytes = await page.evaluate(async () => Array.from(await (window as any).grillon.doc.exportXlsxBytes()));
  const xl = new ExcelJS.Workbook();
  await xl.xlsx.load(new Uint8Array(bytes as number[]) as unknown as ArrayBuffer);
  const ws = xl.worksheets[0];
  expect([2, 3, 4, 5, 6, 7, 8].map((r) => ws.getRow(r).outlineLevel ?? 0)).toEqual([2, 2, 1, 2, 2, 1, 0]);
});

test('Grouper puis masquer et afficher le détail', async ({ page }) => {
  await run(page, `ws.getRange('A1:A6').setValues([['T'],[1],[2],[3],[4],['Total']]);`);
  await select(page, 'A2:A5');
  await dataTab(page);
  await ribbon(page, 'grillon.data.outline', 'Grouper les lignes');
  await select(page, 'A3');
  await ribbon(page, 'grillon.data.outline', 'Masquer le détail');
  const hidden = () => run<number[]>(page, `const rd = wb.save().sheets[ws.getSheetId()].rowData || {}; return Object.keys(rd).filter(k => rd[k].hd).map(Number);`);
  await expect.poll(hidden).toEqual([1, 2, 3, 4]);
  await select(page, 'A6');
  await ribbon(page, 'grillon.data.outline', 'Afficher le détail');
  await expect.poll(hidden).toEqual([]);
});

test('Filtre avancé sur place avec zone de critères', async ({ page }) => {
  await run(page, `ws.getRange('A1:C6').setValues([['Nom','Ville','Montant'],['Alice','Lyon',120],['Bob','Paris',80],['Chloé','Lyon',40],['David','Lille',200],['Eve','Paris',150]]);
    ws.getRange('E1:F2').setValues([['Ville','Montant'],['Lyon','>100']]);
    ws.getRange('E3').setValue('Paris');`);
  await select(page, 'A1');
  await dataTab(page);
  await ribbon(page, 'grillon.data.advanced', 'Filtre avancé…');
  const dialog = page.getByTestId('dialog-advanced-filter');
  await dialog.getByLabel('Zone de critères :').fill('$E$1:$F$3');
  await dialog.getByRole('button', { name: 'OK' }).click();
  const hidden = () => run<number[]>(page, `const rd = wb.save().sheets[ws.getSheetId()].rowData || {}; return Object.keys(rd).filter(k => rd[k].hd).map(Number);`);
  await expect.poll(hidden).toEqual([3, 4]);
  await ribbon(page, 'grillon.data.advanced', 'Afficher tout');
  await expect.poll(hidden).toEqual([]);
});

test('Mettre sous forme de tableau, ligne de totaux et référence structurée', async ({ page }) => {
  await run(page, `ws.getRange('A1:B4').setValues([['Produit','Qte'],['Café',2],['Thé',5],['Jus',1]]);`);
  await select(page, 'A2');
  await dataTab(page);
  await ribbon(page, 'grillon.data.table', 'Mettre sous forme de tableau');
  await expect.poll(() => run<string[]>(page, `return wb.getTableList().map(t => t.name);`)).toEqual(['Tableau1']);
  await ribbon(page, 'grillon.data.table', 'Ligne de totaux');
  await expect.poll(() => display(page, 'B5')).toBe('8');
  expect(await display(page, 'A5')).toBe('Total');
  expect(await run<string>(page, `return ws.getRange('B5').getFormula();`)).toBe('=SUBTOTAL(109,Tableau1[Qte])');
});

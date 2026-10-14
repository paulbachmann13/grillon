import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { importXlsx } from '../../src/io/xlsx/import';
import { definedNames, normalize } from '../unit/helpers';
import { display, FIXTURE_DIR, gotoApp, openViaMenu, useFallbackFilePickers } from './helpers';

test.beforeEach(async ({ page }) => {
  await useFallbackFilePickers(page);
  await gotoApp(page);
});

async function saveViaMenu(page: import('@playwright/test').Page): Promise<Uint8Array> {
  const download = page.waitForEvent('download');
  await page.getByTestId('file-menu').click();
  await page.getByTestId('file-save').click();
  const d = await download;
  return new Uint8Array(readFileSync(await d.path()));
}

for (const name of ['styles', 'formats', 'formulas', 'layout', 'large']) {
  test(`round-trip par l’interface : ouvrir puis enregistrer ${name}.xlsx`, async ({ page }) => {
    await openViaMenu(page, `${name}.xlsx`);
    const saved = await saveViaMenu(page);
    const original = await importXlsx(readFileSync(join(FIXTURE_DIR, `${name}.xlsx`)));
    const roundTripped = await importXlsx(saved);
    const a = normalize(original);
    const b = normalize(roundTripped);
    expect(b.map((s) => s.name)).toEqual(a.map((s) => s.name));
    // Univer recalcule les formules : un résultat absent du fichier d'origine peut apparaître.
    const withoutCachedResults = (cells: Record<string, unknown>) =>
      Object.fromEntries(
        Object.entries(cells).map(([k, c]) => {
          const cell = c as { f?: string; v?: unknown; t?: unknown };
          return [k, cell.f ? { ...cell, v: undefined, t: undefined } : cell];
        }),
      );
    for (let i = 0; i < a.length; i++) {
      expect(withoutCachedResults(b[i].cells), `cellules de « ${a[i].name} »`).toEqual(withoutCachedResults(a[i].cells));
      expect(b[i].merges).toEqual(a[i].merges);
      expect(b[i].freeze).toEqual(a[i].freeze);
      expect(b[i].hidden).toEqual(a[i].hidden);
      expect(b[i].columns).toEqual(a[i].columns);
      expect(b[i].rows).toEqual(a[i].rows);
    }
    expect(definedNames(roundTripped)).toEqual(definedNames(original));
  });
}

test('modifier un classeur puis l’enregistrer conserve la modification', async ({ page }) => {
  await openViaMenu(page, 'formulas.xlsx');
  await page.evaluate(() => {
    const wb = (window as any).grillon.api.getActiveWorkbook();
    wb.getSheetByName('Données').getRange('B2').setValue(2);
  });
  await expect.poll(() => display(page, 'D2', 'Données')).toBe('20');
  await expect.poll(() => display(page, 'B1', 'Synthèse')).toBe('33');
  const saved = await importXlsx(await saveViaMenu(page));
  const data = Object.values(saved.sheets).find((s) => s.name === 'Données')!;
  expect(data.cellData?.[1]?.[1]?.v).toBe(2);
  expect(data.cellData?.[1]?.[3]).toMatchObject({ f: '=B2*C2', v: 20 });
  await expect(page.getByTestId('doc-title')).not.toHaveText(/•$/);
});

test('import et export CSV (séparateur ; et virgule décimale)', async ({ page }) => {
  const csv = '﻿Nom;Prix;Date\nCafé;1,5;14/10/2026\nThé;2;01/01/2026\n';
  await page.evaluate(async (text) => {
    await (window as any).grillon.doc.openBytes('achats.csv', new TextEncoder().encode(text));
  }, csv);
  await expect(page.getByTestId('doc-title')).toHaveText(/^achats/);
  expect(await display(page, 'A2')).toBe('Café');
  expect(await display(page, 'B2')).toBe('1,5');
  expect(await display(page, 'C2')).toBe('14/10/2026');

  const download = page.waitForEvent('download');
  await page.getByTestId('file-menu').click();
  await page.getByTestId('file-export-csv').click();
  const d = await download;
  expect(d.suggestedFilename()).toMatch(/\.csv$/);
  const text = readFileSync(await d.path(), 'utf8');
  expect(text.replace(/^﻿/, '').split('\r\n')).toEqual(['Nom;Prix;Date', 'Café;1,5;14/10/2026', 'Thé;2;01/01/2026']);
});

test('sauvegarde automatique : le classeur est restauré après rechargement', async ({ page }) => {
  await openViaMenu(page, 'formats.xlsx');
  await page.evaluate(() => (window as any).grillon.api.getActiveWorkbook().getActiveSheet().getRange('C1').setValue('Modifié'));
  // Délai de la sauvegarde automatique (1,5 s) + marge.
  await page.waitForTimeout(2500);
  await page.reload();
  await expect(page.getByTestId('doc-title')).toHaveText(/^formats/);
  await expect.poll(() => display(page, 'C1')).toBe('Modifié');
  await expect.poll(() => display(page, 'B5')).toBe('14/10/2026');
});

test('onglets de feuille : ajouter, renommer, réordonner, supprimer', async ({ page }) => {
  const names = () => page.evaluate(() => (window as any).grillon.api.getActiveWorkbook().getSheets().map((s: any) => s.getSheetName()));
  await page.evaluate(() => {
    const wb = (window as any).grillon.api.getActiveWorkbook();
    wb.insertSheet('Budget');
    wb.getSheetByName('Budget').setName('Budget 2026');
    wb.insertSheet('Notes');
    wb.getSheetByName('Notes').moveSheet?.(0) ?? wb.moveSheet(wb.getSheetByName('Notes'), 0);
  });
  await expect.poll(names).toEqual(['Notes', 'Feuil1', 'Budget 2026']);
  await expect(page.getByText('Budget 2026', { exact: true })).toBeVisible();
  await page.evaluate(() => {
    const wb = (window as any).grillon.api.getActiveWorkbook();
    wb.deleteSheet(wb.getSheetByName('Notes'));
  });
  await expect.poll(names).toEqual(['Feuil1', 'Budget 2026']);
});

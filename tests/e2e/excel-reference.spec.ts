/**
 * Comparaison avec le vrai Excel : mêmes formules et mêmes formats que ceux relevés dans Excel (tests/fixtures/excel-reference.json).
 * Protège aussi les corrections apportées au moteur d'Univer (custom-functions.ts) : à relancer à chaque mise à jour d'Univer.
 */
import { expect, test } from '@playwright/test';
import { exportXlsx } from '../../src/io/xlsx/export';
import { FORMULAS, KNOWN_DIFFERENCES, NUMBER_FORMATS, referenceWorkbook } from '../fixtures/excel-reference';
import reference from '../fixtures/excel-reference.json' with { type: 'json' };
import { gotoApp } from './helpers';

const ERRORS: Record<string, string> = { '#VALEUR!': '#VALUE!', '#NOM?': '#NAME?', '#NOMBRE!': '#NUM!', '#NUL!': '#NULL!', '#EPARS!': '#SPILL!' };
const norm = (s: unknown) => String(s ?? '').replace(/[  ]/g, ' ').trim();
const num = (s: unknown) => {
  if (typeof s === 'number') return s;
  const t = norm(s).replace(/ /g, '').replace(',', '.');
  return t !== '' && Number.isFinite(Number(t)) ? Number(t) : null;
};

test('formules et formats : résultats identiques à ceux d’Excel', async ({ page }) => {
  await gotoApp(page);
  const bytes = [...(await exportXlsx(referenceWorkbook()))];
  await page.evaluate(async (d) => {
    await (window as never as { grillon: { doc: { openBytes(n: string, b: Uint8Array): Promise<void> } } }).grillon.doc.openBytes('reference.xlsx', new Uint8Array(d));
  }, bytes);
  const read = () =>
    page.evaluate(([nf, nn]) => {
      const wb = (window as never as { grillon: { api: { getActiveWorkbook(): { getSheetByName(n: string): { getRange(a: string): { getCellDataGrid(): { v?: unknown }[][]; getDisplayValues(): string[][] } } } } } }).grillon.api.getActiveWorkbook();
      const f = wb.getSheetByName('F').getRange(`B1:B${nf}`);
      return { raw: f.getCellDataGrid().map((r) => r[0]?.v ?? null), shown: f.getDisplayValues().map((r) => r[0]), formats: wb.getSheetByName('N').getRange(`C1:C${nn}`).getDisplayValues().map((r) => r[0]) };
    }, [FORMULAS.length, NUMBER_FORMATS.length]);
  await expect.poll(async () => (await read()).raw[0], { timeout: 30_000 }).toBe(77);
  await page.waitForTimeout(3000);
  const got = await read();
  const diffs: string[] = [];
  FORMULAS.forEach((formula, i) => {
    const excel = reference.formules[i]!;
    if (KNOWN_DIFFERENCES.has(formula)) return;
    // Fonction absente de la version d'Excel du relevé : rien à comparer.
    if (excel === '#NOM?' && got.shown[i] !== '#NOM?') return;
    const a = num(excel);
    const b = num(got.raw[i]);
    if (a !== null && b !== null && Math.abs(a - b) <= Math.max(1e-8, Math.abs(a) * 2e-8)) return;
    if (norm(ERRORS[excel] ?? excel) === norm(ERRORS[got.shown[i]!] ?? got.shown[i])) return;
    // Complexes : Excel en français écrit la virgule.
    if (/^IM[A-Z0-9]+\(/.test(formula) && norm(excel).replace(/,/g, '.').toLowerCase() === norm(got.shown[i]).toLowerCase()) return;
    diffs.push(`${formula} : Excel ${excel}, Grillon ${got.shown[i]}`);
  });
  NUMBER_FORMATS.forEach(([value, format], i) => {
    const excel = norm(reference.formats[i]);
    const mine = norm(got.formats[i]);
    if (/^#+$/.test(excel) || excel === mine || mine.replace(/\./g, '') === excel) return;
    if (format === 'h:mm:ss') return; // format n° 21 : Excel en français l'affiche hh:mm:ss.
    diffs.push(`${value} au format ${format} : Excel ${excel}, Grillon ${mine}`);
  });
  expect(diffs).toEqual([]);
});

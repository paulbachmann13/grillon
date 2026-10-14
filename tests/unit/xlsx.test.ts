import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { exportXlsx } from '../../src/io/xlsx/export';
import { importXlsx } from '../../src/io/xlsx/import';
import { BooleanNumber, BorderStyleTypes, CellValueType, HorizontalAlign, VerticalAlign, WrapStrategy } from '../../src/model/univer-types';
import { buildFixture, FIXTURES } from '../fixtures/build';
import { cellAt, definedNames, normalize, sheetByName, styleOf } from './helpers';

async function load(name: Parameters<typeof buildFixture>[0]) {
  return importXlsx(await buildFixture(name));
}

describe('import xlsx — styles', () => {
  it('lit police, couleurs, bordures, alignements', async () => {
    const wb = await load('styles');
    const s = (a1: string) => styleOf(wb, cellAt(wb, 'Styles', a1));
    expect(s('A1')).toMatchObject({ ff: 'Calibri', fs: 14, bl: BooleanNumber.TRUE });
    expect(s('A2')).toMatchObject({ ff: 'Arial', fs: 10, it: BooleanNumber.TRUE, ul: { s: 1 } });
    expect(s('A3')).toMatchObject({ st: { s: 1 }, cl: { rgb: '#FF0000' } });
    expect(s('A4')).toMatchObject({ bg: { rgb: '#FFFF00' } });
    expect(s('A5').bd).toMatchObject({
      t: { s: BorderStyleTypes.THIN, cl: { rgb: '#000000' } },
      b: { s: BorderStyleTypes.DOUBLE, cl: { rgb: '#0000FF' } },
      r: { s: BorderStyleTypes.MEDIUM },
    });
    expect(s('B1')).toMatchObject({ ht: HorizontalAlign.CENTER, vt: VerticalAlign.MIDDLE });
    expect(s('B2')).toMatchObject({ tb: WrapStrategy.WRAP, vt: VerticalAlign.TOP });
    expect(s('B3')).toMatchObject({ ht: HorizontalAlign.RIGHT });
    // ExcelJS écrit le thème Office 2007 (accent1 = 4F81BD).
    expect(s('B4')).toMatchObject({ cl: { rgb: '#4F81BD' } });
    expect(s('C1').tr).toEqual({ a: 45 });
    expect(s('C3')).toMatchObject({ bg: { rgb: '#00B050' } });
    expect(sheetByName(wb, 'Styles').columnData?.[1]?.w).toBe(215);
  });
});

describe('import xlsx — types et formats', () => {
  it('lit nombres, dates, booléens, texte accentué et formats', async () => {
    const wb = await load('formats');
    const c = (a1: string) => cellAt(wb, 'Formats', a1)!;
    expect(c('B1')).toMatchObject({ v: 42, t: CellValueType.NUMBER });
    expect(styleOf(wb, c('B2')).n).toEqual({ pattern: '0.00' });
    expect(styleOf(wb, c('B3')).n).toEqual({ pattern: '0.0%' });
    expect(styleOf(wb, c('B4')).n?.pattern).toBe('#,##0.00 "€"');
    // 14/10/2026 = série 46309
    expect(c('B5').v).toBe(46309);
    expect(styleOf(wb, c('B5')).n).toEqual({ pattern: 'dd/mm/yyyy' });
    expect(c('B6').v).toBeCloseTo(46053 + (13 * 60 + 45) / 1440, 8);
    expect(c('B9')).toMatchObject({ v: 1, t: CellValueType.BOOLEAN });
    expect(c('B10')).toMatchObject({ v: 0, t: CellValueType.BOOLEAN });
    expect(c('B11')).toMatchObject({ v: 'Élève à l’école, garçon, œuvre, ÇA', t: CellValueType.STRING });
    expect(c('B12')).toMatchObject({ v: '00123', t: CellValueType.STRING });
  });
});

describe('import xlsx — formules et noms', () => {
  it('lit formules, résultats en cache, erreurs, références inter-feuilles', async () => {
    const wb = await load('formulas');
    expect(cellAt(wb, 'Données', 'D2')).toMatchObject({ f: '=B2*C2', v: 5 });
    expect(cellAt(wb, 'Données', 'D7')).toMatchObject({ f: '=SUM(D2:D5)', v: 18 });
    expect(cellAt(wb, 'Données', 'D9')).toMatchObject({ f: '=IF(D7>10,"Gros","Petit")', v: 'Gros' });
    expect(cellAt(wb, 'Données', 'D11')).toMatchObject({ f: '=1/0', v: '#DIV/0!' });
    expect(cellAt(wb, 'Synthèse', 'B1')?.f).toBe("=SUM('Données'!D2:D5)");
    expect(definedNames(wb)).toEqual([{ name: 'Totaux', ref: "'Données'!$D$2:$D$5" }]);
  });
});

describe('import xlsx — mise en page', () => {
  it('lit fusions, volets figés, tailles, feuilles masquées', async () => {
    const wb = await load('layout');
    const a = sheetByName(wb, 'Mise en page');
    expect(a.mergeData).toEqual(
      expect.arrayContaining([
        { startRow: 0, endRow: 0, startColumn: 0, endColumn: 3 },
        { startRow: 2, endRow: 4, startColumn: 1, endColumn: 2 },
      ]),
    );
    expect(a.freeze).toEqual({ xSplit: 1, ySplit: 2, startRow: 2, startColumn: 1 });
    expect(a.tabColor).toBe('#4472C4');
    expect(a.columnData?.[0]?.w).toBe(180);
    expect(a.columnData?.[4]?.hd).toBe(1);
    expect(a.rowData?.[1]?.h).toBe(53);
    expect(a.rowData?.[5]?.hd).toBe(1);
    expect(wb.sheetOrder.map((id) => wb.sheets[id]!.name)).toEqual(['Mise en page', 'Deuxième feuille', 'Masquée', 'Sans quadrillage']);
    expect(sheetByName(wb, 'Masquée').hidden).toBe(1);
    expect(sheetByName(wb, 'Sans quadrillage').showGridlines).toBe(0);
  });
});

describe('round-trip import → export → import', () => {
  for (const name of FIXTURES) {
    it(`conserve tout : ${name}`, async () => {
      const first = await load(name);
      const exported = await exportXlsx(first);
      const second = await importXlsx(exported);
      expect(normalize(second)).toEqual(normalize(first));
      expect(definedNames(second)).toEqual(definedNames(first));
    });
  }

  it('produit un fichier relisible par ExcelJS avec les mêmes valeurs (gros volume)', async () => {
    const first = await load('large');
    const out = await exportXlsx(first);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(out as unknown as ArrayBuffer);
    const ws = wb.getWorksheet('Gros volume')!;
    expect(ws.rowCount).toBe(50_001);
    expect(ws.getCell('B50001').value).toBe('Élément 50000');
    expect(ws.getCell('G2').formula).toBe('SUM(D2:D50001)');
    expect(ws.getCell('F4').value).toBe(true);
  });
});

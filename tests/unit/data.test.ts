import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { exportXlsx } from '../../src/io/xlsx/export';
import { importXlsx } from '../../src/io/xlsx/import';
import type { IWorkbookData } from '../../src/model/univer-types';
import { buildFixture } from '../fixtures/build';
import { cellAt, normalize, sheetByName } from './helpers';

function resource<T>(wb: IWorkbookData, name: string): T {
  const r = wb.resources?.find((x) => x.name === name);
  return (r ? JSON.parse(r.data) : {}) as T;
}

/** Ressources par nom de feuille, sans identifiants générés. */
function dataResources(wb: IWorkbookData) {
  const byName = (id: string) => wb.sheets[id]?.name ?? id;
  const dv = resource<Record<string, Record<string, unknown>[]>>(wb, 'SHEET_DATA_VALIDATION_PLUGIN');
  const filters = resource<Record<string, unknown>>(wb, 'SHEET_FILTER_PLUGIN');
  const tables = resource<Record<string, { tables: Record<string, unknown>[] }>>(wb, 'SHEET_TABLE_PLUGIN');
  const strip = (o: Record<string, unknown>) => {
    const { uid: _u, id: _i, ...rest } = o;
    return rest;
  };
  return {
    dv: Object.fromEntries(
      Object.entries(dv).map(([k, rules]) => [
        byName(k),
        rules.map((r) => ({ ...strip(r), ranges: (r.ranges as Record<string, number>[]).map(({ startRow, endRow, startColumn, endColumn }) => ({ startRow, endRow, startColumn, endColumn })) })),
      ]),
    ),
    filters: Object.fromEntries(Object.entries(filters).map(([k, f]) => [byName(k), f])),
    tables: Object.fromEntries(
      Object.entries(tables).map(([k, t]) => [
        byName(k),
        t.tables.map((tb) => ({ ...strip(tb), columns: (tb.columns as Record<string, unknown>[]).map(strip) })),
      ]),
    ),
  };
}

describe('validation de données, filtres, tableaux — import', () => {
  it('lit toutes les validations, y compris sur une colonne entière, sans les déplier', async () => {
    const t0 = Date.now();
    const wb = await importXlsx(await buildFixture('data'));
    expect(Date.now() - t0).toBeLessThan(5000);
    const { dv } = dataResources(wb);
    const rules = dv['Données'] as Record<string, unknown>[];
    const byCol = (c: number) => rules.find((r) => (r.ranges as { startColumn: number }[])[0].startColumn === c)!;
    expect(byCol(1)).toMatchObject({ type: 'list', formula1: '["Lyon","Paris","Nîmes"]', showDropDown: true, allowBlank: true, showErrorMessage: true, error: 'Choisir une ville de la liste' });
    expect(byCol(2)).toMatchObject({ type: 'whole', operator: 'between', formula1: '0', formula2: '1000', showInputMessage: true, prompt: 'Entier entre 0 et 1000', errorStyle: 2 });
    expect(byCol(3)).toMatchObject({ type: 'date', operator: 'greaterThanOrEqual', formula1: '2026-01-01' });
    expect(byCol(4)).toMatchObject({ type: 'textLength', operator: 'lessThanOrEqual', formula1: '5' });
    expect(byCol(5)).toMatchObject({ type: 'list', formula1: '=$H$1:$H$3', formula2: '1' });
    expect(byCol(0)).toMatchObject({ type: 'custom', formula1: '=LEN(A2)>3' });
    expect(byCol(9)).toMatchObject({ type: 'decimal', ranges: [{ startRow: 0, endRow: 1048575, startColumn: 9, endColumn: 9 }] });
  });

  it('lit le filtre automatique avec ses critères et ne garde pas les lignes filtrées comme masquées', async () => {
    const wb = await importXlsx(await buildFixture('data'));
    const { filters } = dataResources(wb);
    expect(filters['Données']).toEqual({
      ref: { startRow: 0, endRow: 10, startColumn: 0, endColumn: 5 },
      filterColumns: [
        { colId: 1, filters: { filters: ['Lyon', 'Paris'] } },
        { colId: 2, customFilters: { customFilters: [{ val: 20, operator: 'greaterThan' }] } },
      ],
    });
    const rows = sheetByName(wb, 'Données').rowData ?? {};
    expect(Object.values(rows).some((r) => r?.hd)).toBe(false);
  });

  it('lit le tableau structuré, sa ligne de totaux et les références structurées', async () => {
    const wb = await importXlsx(await buildFixture('data'));
    const { tables } = dataResources(wb);
    const [t] = tables['Tableau'] as Record<string, unknown>[];
    expect(t).toMatchObject({
      name: 'Ventes',
      range: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 },
      options: { showHeader: true, showAutoFilter: true },
      meta: { excel: { totalsRowCount: 1, styleName: 'TableStyleMedium9', showRowStripes: true } },
    });
    expect((t.columns as Record<string, unknown>[]).map((c) => [c.displayName, c.totalsRowFunction, c.totalsRowLabel])).toEqual([
      ['Produit', undefined, 'Total'],
      ['Qté', 'sum', undefined],
      ['Prix', 'average', undefined],
    ]);
    expect(cellAt(wb, 'Tableau', 'B5')?.f).toBe('=SUBTOTAL(109,Ventes[Qté])');
    expect(cellAt(wb, 'Tableau', 'F1')?.f).toBe('=SUM(Ventes[Prix])');
  });
});

describe('tableau créé dans Grillon', () => {
  it('une ligne de SUBTOTAL sous le tableau devient sa ligne de totaux Excel', async () => {
    const sheetId = 's1';
    const wb: IWorkbookData = {
      id: 'w',
      name: 'w',
      appVersion: '1',
      locale: 'frFR' as IWorkbookData['locale'],
      styles: {},
      sheetOrder: [sheetId],
      sheets: {
        [sheetId]: {
          id: sheetId,
          name: 'F',
          cellData: {
            0: { 0: { v: 'Produit', t: 1 }, 1: { v: 'Qte', t: 1 } },
            1: { 0: { v: 'Café', t: 1 }, 1: { v: 2, t: 2 } },
            2: { 0: { v: 'Thé', t: 1 }, 1: { v: 5, t: 2 } },
            3: { 0: { v: 'Total', t: 1 }, 1: { f: '=SUBTOTAL(109,Tableau1[Qte])', v: 7, t: 2 } },
          },
        },
      },
      resources: [
        {
          name: 'SHEET_TABLE_PLUGIN',
          data: JSON.stringify({
            [sheetId]: {
              tables: [
                {
                  id: 't1',
                  name: 'Tableau1',
                  range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
                  options: { showHeader: true, showAutoFilter: true },
                  filters: { tableColumnFilterList: [] },
                  columns: [
                    { id: 'a', displayName: 'Produit', dataType: 'string', formula: '', showFilterButton: true, meta: {}, style: {} },
                    { id: 'b', displayName: 'Qte', dataType: 'string', formula: '', showFilterButton: true, meta: {}, style: {} },
                  ],
                },
              ],
              tableFilteredOutRows: [],
            },
          }),
        },
      ],
    };
    const zip = await JSZip.loadAsync(await exportXlsx(wb));
    const table = await zip.file('xl/tables/table1.xml')!.async('string');
    expect(table).toContain('ref="A1:B4"');
    expect(table).toContain('totalsRowCount="1"');
    expect(table).toContain('<tableColumn id="1" name="Produit" totalsRowLabel="Total"/>');
    expect(table).toContain('<tableColumn id="2" name="Qte" totalsRowFunction="sum"/>');
    const back = await importXlsx(await zip.generateAsync({ type: 'uint8array' }));
    const t = dataResources(back).tables.F[0] as unknown as { range: object; meta: { excel: { totalsRowCount: number } } };
    expect(t.range).toEqual({ startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 });
    expect(t.meta.excel.totalsRowCount).toBe(1);
  });
});

describe('validation de données, filtres, tableaux — round-trip', () => {
  it('conserve validations, critères de filtre, tableaux et cellules', async () => {
    const first = await importXlsx(await buildFixture('data'));
    const out = await exportXlsx(first);
    const second = await importXlsx(out);
    expect(dataResources(second)).toEqual(dataResources(first));
    expect(normalize(second)).toEqual(normalize(first));
  });

  it('produit un XML valide : ordre des éléments, parties tableau, lignes filtrées masquées', async () => {
    const first = await importXlsx(await buildFixture('data'));
    const filter = resource<Record<string, { cachedFilteredOut?: number[] }>>(first, 'SHEET_FILTER_PLUGIN');
    const sheetId = Object.keys(filter)[0];
    filter[sheetId].cachedFilteredOut = [1, 2, 3, 6, 9];
    first.resources = first.resources!.map((r) => (r.name === 'SHEET_FILTER_PLUGIN' ? { ...r, data: JSON.stringify(filter) } : r));
    const zip = await JSZip.loadAsync(await exportXlsx(first));
    const sheet1 = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    const order = ['<sheetData', '<autoFilter', '<mergeCells', '<dataValidations', '<pageMargins'].map((t) => sheet1.indexOf(t)).filter((i) => i >= 0);
    expect(order).toEqual([...order].sort((a, b) => a - b));
    expect(sheet1).toMatch(/<autoFilter ref="A1:F11"><filterColumn colId="1"><filters><filter val="Lyon"\/><filter val="Paris"\/><\/filters><\/filterColumn>/);
    expect([...sheet1.matchAll(/<row r="(\d+)"[^>]*hidden="1"/g)].map((m) => Number(m[1]))).toEqual([2, 3, 4, 7, 10]);
    expect(sheet1).toMatch(/sqref="J1:J1048576"/);
    const sheet2 = await zip.file('xl/worksheets/sheet2.xml')!.async('string');
    expect(sheet2).toMatch(/<tableParts count="1"><tablePart r:id="rId\d+"\/><\/tableParts>/);
    const table = await zip.file('xl/tables/table1.xml')!.async('string');
    expect(table).toContain('ref="A1:C5"');
    expect(table).toContain('totalsRowCount="1"');
    expect(table).toContain('<tableColumn id="2" name="Qté" totalsRowFunction="sum"/>');
    const types = await zip.file('[Content_Types].xml')!.async('string');
    expect(types).toContain('/xl/tables/table1.xml');
    // ExcelJS relit le fichier sans erreur et retrouve le tableau.
    const reread = new ExcelJS.Workbook();
    await reread.xlsx.load((await zip.generateAsync({ type: 'uint8array' })) as unknown as ArrayBuffer);
    expect(reread.getWorksheet('Tableau')!.getCell('B5').formula).toBe('SUBTOTAL(109,Ventes[Qté])');
  });
});

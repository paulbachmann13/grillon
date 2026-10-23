import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import type { PivotDefinition } from '../../src/data/pivot';
import { exportXlsx } from '../../src/io/xlsx/export';
import { importXlsx } from '../../src/io/xlsx/import';
import { PIVOT_DEF, pivotWorkbook } from '../fixtures/pivot';

const pivotOf = (wb: Awaited<ReturnType<typeof importXlsx>>) =>
  Object.values(wb.sheets)
    .map((s) => (s as { custom?: { grillon?: { pivot?: PivotDefinition } } }).custom?.grillon?.pivot)
    .find(Boolean);

describe('tableaux croisés dynamiques — .xlsx', () => {
  it('écrit un vrai tableau croisé Excel (cache, enregistrements, définition, relations)', async () => {
    const zip = await JSZip.loadAsync(await exportXlsx(pivotWorkbook()));
    const workbook = await zip.file('xl/workbook.xml')!.async('string');
    expect(workbook).toMatch(/<pivotCaches><pivotCache cacheId="1" r:id="rId\d+"\/><\/pivotCaches>/);
    const cache = await zip.file('xl/pivotCache/pivotCacheDefinition1.xml')!.async('string');
    expect(cache).toContain('refreshOnLoad="1"');
    expect(cache).toContain('<worksheetSource ref="A1:D7" sheet="Ventes"/>');
    expect(cache).toContain('<cacheField name="Ville" numFmtId="0"><sharedItems count="3"><s v="Lyon"/><s v="Nîmes"/><s v="Paris"/></sharedItems></cacheField>');
    expect(cache).toMatch(/<cacheField name="Montant" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="3" maxValue="12"\/>/);
    const records = await zip.file('xl/pivotCache/pivotCacheRecords1.xml')!.async('string');
    expect(records).toContain('count="6"');
    expect(records).toContain('<r><x v="0"/><x v="0"/><n v="10"/><n v="1"/></r>');
    const table = await zip.file('xl/pivotTables/pivotTable1.xml')!.async('string');
    expect(table).toContain('<location ref="A3:F8"');
    expect(table).toContain('<rowFields count="1"><field x="0"/></rowFields>');
    expect(table).toContain('<colFields count="2"><field x="1"/><field x="-2"/></colFields>');
    expect(table).toContain('<dataField name="Moyenne de Quantité" fld="3" subtotal="average"');
    const sheetRels = await zip.file('xl/worksheets/_rels/sheet2.xml.rels')!.async('string');
    expect(sheetRels).toContain('relationships/pivotTable');
    const types = await zip.file('[Content_Types].xml')!.async('string');
    expect(types).toContain('/xl/pivotTables/pivotTable1.xml');
    expect(types).toContain('/xl/pivotCache/pivotCacheRecords1.xml');
  });

  it('round-trip : la définition est relue (lignes, colonnes, valeurs, filtres, emplacement)', async () => {
    const def: PivotDefinition = { ...PIVOT_DEF, config: { ...PIVOT_DEF.config, filters: [{ col: 1, include: ['Café'] }], columns: [] } };
    const back = pivotOf(await importXlsx(await exportXlsx(pivotWorkbook(def))));
    expect(back).toMatchObject({
      sourceSheet: 'Ventes',
      sourceRange: 'A1:D7',
      config: { rows: [0], columns: [], values: [{ col: 2, fn: 'sum' }, { col: 3, fn: 'average' }], filters: [{ col: 1, include: ['Café'] }] },
      outRow: 2,
      outCol: 0,
      name: 'TCD1',
    });
  });

  it('relit un filtre sur des éléments qui ne sont pas les premiers du cache', async () => {
    const def: PivotDefinition = { ...PIVOT_DEF, config: { ...PIVOT_DEF.config, columns: [1], filters: [{ col: 0, include: ['Nîmes', 'Paris'] }] } };
    const back = pivotOf(await importXlsx(await exportXlsx(pivotWorkbook(def))));
    expect(back?.config.filters).toEqual([{ col: 0, include: ['Nîmes', 'Paris'] }]);
  });

  it('un classeur sans tableau croisé n’a pas de parties pivot', async () => {
    const wb = pivotWorkbook();
    delete (wb.sheets.s2 as { custom?: unknown }).custom;
    const zip = await JSZip.loadAsync(await exportXlsx(wb));
    expect(Object.keys(zip.files).some((f) => f.includes('pivot'))).toBe(false);
  });
});

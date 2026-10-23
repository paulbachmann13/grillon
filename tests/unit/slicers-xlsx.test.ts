import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import type { SlicerSpec } from '../../src/data/slicer';
import { exportXlsx } from '../../src/io/xlsx/export';
import { importXlsx } from '../../src/io/xlsx/import';
import { PIVOT_SLICER, slicerWorkbook, TABLE_SLICER, TIMELINE } from '../fixtures/slicers';

type Wb = Awaited<ReturnType<typeof importXlsx>>;
const resource = (wb: Wb, name: string) => JSON.parse(wb.resources!.find((r) => r.name === name)?.data ?? '{}');
const slicersOf = (wb: Wb) => {
  const drawings = resource(wb, 'SHEET_DRAWING_PLUGIN') as Record<string, { data: Record<string, { componentKey: string; data: SlicerSpec; sheetTransform: { from: { column: number } } }> }>;
  return Object.entries(drawings).flatMap(([sheetId, d]) =>
    Object.values(d.data)
      .filter((x) => x.componentKey === 'GrillonSlicer')
      .map((x) => ({ sheet: wb.sheets[sheetId]!.name, spec: x.data, column: x.sheetTransform.from.column })),
  );
};

describe('segments et chronologies — .xlsx', () => {
  it('écrit de vrais segments d’Excel (tableau et tableau croisé)', async () => {
    const zip = await JSZip.loadAsync(await exportXlsx(slicerWorkbook()));
    const file = (p: string) => zip.file(p)!.async('string');
    // La plage filtrée devient un tableau sans style, filtré comme le segment.
    const table = await file('xl/tables/table1.xml');
    expect(table).toContain('ref="A1:D7"');
    expect(table).toContain('<filterColumn colId="0"><filters><filter val="Lyon"/><filter val="Paris"/></filters></filterColumn>');
    expect(table).not.toMatch(/tableStyleInfo[^>]*name=/);
    const caches = Object.keys(zip.files).filter((p) => p.startsWith('xl/slicerCaches/') && p.endsWith('.xml'));
    expect(caches).toHaveLength(2);
    const all = (await Promise.all(caches.map(file))).join('\n');
    expect(all).toContain('name="Segment_Ville" sourceName="Ville"');
    expect(all).toContain('<x15:tableSlicerCache tableId="1" column="1"/>');
    expect(all).toMatch(/name="Segment_Produit" sourceName="Produit"><pivotTables><pivotTable tabId="2" name="TCD1"\/><\/pivotTables><data><tabular pivotCacheId="1"><items count="2"><i x="0" s="1"\/><i x="1"\/><\/items>/);
    expect(await file('xl/pivotCache/pivotCacheDefinition1.xml')).toContain('<x14:pivotCacheDefinition pivotCacheId="1"/>');
    const workbook = await file('xl/workbook.xml');
    expect(workbook).toContain('<definedName name="Segment_Ville">#N/A</definedName>');
    expect(workbook).toMatch(/<x15:slicerCaches[^>]*><x14:slicerCache r:id="rId\d+"\/><\/x15:slicerCaches>/);
    expect(workbook).toMatch(/<x14:slicerCaches><x14:slicerCache r:id="rId\d+"\/><\/x14:slicerCaches>/);
    expect(workbook.indexOf('<extLst>')).toBeGreaterThan(workbook.indexOf('</definedNames>'));
    const sheet1 = await file('xl/worksheets/sheet1.xml');
    expect(sheet1).toMatch(/<extLst><ext uri="\{3A4CF648-6AED-40f4-86FF-DC5316D8AED3\}"[^>]*><x14:slicerList[^>]*><x14:slicer r:id="rId\d+"\/><\/x14:slicerList><\/ext><\/extLst><\/worksheet>$/);
    const slicers1 = Object.keys(zip.files).filter((p) => p.startsWith('xl/slicers/') && p.endsWith('.xml'));
    expect(slicers1).toHaveLength(2);
    const drawing = await file('xl/drawings/drawing1.xml');
    expect(drawing).toContain('Requires="sle15"');
    expect(drawing).toContain('<sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="Ville"/>');
    const types = await file('[Content_Types].xml');
    expect(types).toContain('application/vnd.ms-excel.slicerCache+xml');
    expect(types).toContain('application/vnd.ms-excel.slicer+xml');
    // La chronologie (pas de chronologie sur un tableau dans Excel) est gardée pour Grillon.
    expect(await file('customXml/item1.xml')).toContain('urn:grillon:segments:1');
  });

  it('relit segments et chronologie à l’identique, sans le tableau créé pour l’occasion', async () => {
    const back = await importXlsx(await exportXlsx(slicerWorkbook()));
    const slicers = slicersOf(back);
    expect(slicers.map((s) => s.spec.id).sort()).toEqual([PIVOT_SLICER.id, TABLE_SLICER.id, TIMELINE.id].sort());
    expect(slicers.find((s) => s.spec.id === TABLE_SLICER.id)).toMatchObject({ sheet: 'Ventes', column: 6, spec: TABLE_SLICER });
    expect(slicers.find((s) => s.spec.id === PIVOT_SLICER.id)).toMatchObject({ sheet: 'TCD', spec: PIVOT_SLICER });
    expect(slicers.find((s) => s.spec.id === TIMELINE.id)?.spec).toEqual(TIMELINE);
    expect(resource(back, 'SHEET_TABLE_PLUGIN')).toEqual({});
    // Les noms « Segment_… » d'Excel ne deviennent pas des noms du classeur.
    expect(JSON.stringify(back.resources)).not.toContain('Segment_Ville');
  });

  it('relit les segments d’un fichier d’Excel (sans l’état de Grillon)', async () => {
    const zip = await JSZip.loadAsync(await exportXlsx(slicerWorkbook()));
    zip.remove('customXml/item1.xml');
    const rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string');
    zip.file('xl/_rels/workbook.xml.rels', rels.replace(/<Relationship[^>]*customXml"[^>]*\/>/, ''));
    const back = await importXlsx(await zip.generateAsync({ type: 'uint8array' }));
    const slicers = slicersOf(back);
    expect(slicers).toHaveLength(2);
    expect(slicers.find((s) => s.spec.title === 'Ville')?.spec).toMatchObject({ type: 'slicer', target: { kind: 'table', sheet: 'Ventes', ref: 'A1:D7', col: 0 }, selected: ['Lyon', 'Paris'] });
    expect(slicers.find((s) => s.spec.title === 'Produit')?.spec).toMatchObject({ target: { kind: 'pivot', sheet: 'TCD', col: 1 }, selected: ['Café'] });
    // Le tableau d'Excel reste un tableau.
    expect(resource(back, 'SHEET_TABLE_PLUGIN')).toHaveProperty(Object.keys(back.sheets)[0]!);
  });

  it('n’écrit pas de segment Excel pour une plage déjà munie d’un filtre automatique', async () => {
    const wb = slicerWorkbook([{ sheet: 's1', spec: TABLE_SLICER }]);
    wb.resources!.push({ name: 'SHEET_FILTER_PLUGIN', data: JSON.stringify({ s1: { ref: { startRow: 0, endRow: 6, startColumn: 0, endColumn: 3 } } }) });
    const zip = await JSZip.loadAsync(await exportXlsx(wb));
    expect(Object.keys(zip.files).some((p) => p.startsWith('xl/slicerCaches/'))).toBe(false);
    expect(zip.file('xl/tables/table1.xml')).toBeNull();
    const back = slicersOf(await importXlsx(await zip.generateAsync({ type: 'uint8array' })));
    expect(back.map((s) => s.spec)).toEqual([TABLE_SLICER]);
  });
});

/**
 * Fichiers écrits par d'autres logiciels qu'Excel de bureau (ou par Excel, sous des formes rares) : défauts trouvés sur les
 * fichiers de test d'Apache POI.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { exportXlsx } from '../../src/io/xlsx/export';
import { importXlsx } from '../../src/io/xlsx/import';
import { numberRowsAndCells, unprefixRoot } from '../../src/io/xlsx/normalize';
import type { IWorkbookData } from '../../src/model/univer-types';
import { pivotWorkbook, PIVOT_DEF } from '../fixtures/pivot';

const book = (sheet: object, resources: IWorkbookData['resources'] = []): IWorkbookData =>
  ({ id: 'w', name: 'w', appVersion: '1', locale: 'frFR', styles: {}, sheetOrder: ['s'], sheets: { s: { id: 's', name: 'Feuil1', ...sheet } }, resources }) as never;
const first = (wb: IWorkbookData) => Object.values(wb.sheets)[0]!;
const zipOf = async (wb: IWorkbookData) => JSZip.loadAsync(await exportXlsx(wb));
const bytes = (zip: JSZip) => zip.generateAsync({ type: 'uint8array' });
const SIMPLE = book({ cellData: { 0: { 0: { v: 'Nom', t: 1 }, 1: { v: 'Note', t: 1 } }, 1: { 0: { v: 'Léa', t: 1 }, 1: { v: 15, t: 2 } } } });

describe('fichiers .xlsx inhabituels', () => {
  it('balises préfixées (<x:worksheet>) : relues', async () => {
    expect(unprefixRoot('<?xml version="1.0"?><x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData/></x:worksheet>')).toBe(
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
    );
    const zip = await zipOf(SIMPLE);
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    zip.file('xl/worksheets/sheet1.xml', sheet.replace(/<(\/?)(?!\?)([a-zA-Z])/g, '<$1x:$2').replace('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"', 'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'));
    const wb = await importXlsx(await bytes(zip));
    expect(first(wb).cellData![1]![1]).toMatchObject({ v: 15 });
  });

  it('lignes et cellules sans numéro : numérotées dans l’ordre', () => {
    expect(numberRowsAndCells('<sheetData><row><c t="s"><v>0</v></c><c/></row><row r="5"><c/></row><row><c r="C7"/><c/></row></sheetData>')).toBe(
      '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"/></row><row r="5"><c r="A5"/></row><row r="6"><c r="C7"/><c r="D6"/></row></sheetData>',
    );
  });

  it('feuille nommée « sheet.xml » et propriétés du document inhabituelles : relues', async () => {
    const zip = await zipOf(SIMPLE);
    zip.file('xl/worksheets/sheet.xml', await zip.file('xl/worksheets/sheet1.xml')!.async('string'));
    zip.remove('xl/worksheets/sheet1.xml');
    const rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string');
    zip.file('xl/_rels/workbook.xml.rels', rels.replace('worksheets/sheet1.xml', 'worksheets/sheet.xml'));
    zip.file('docProps/core.xml', '<?xml version="1.0"?><coreProperties xmlns="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"><lastModifiedBy>x</lastModifiedBy></coreProperties>');
    const wb = await importXlsx(await bytes(zip));
    expect(first(wb).cellData![1]![0]).toMatchObject({ v: 'Léa' });
  });

  it('validation de données sur une plage devenue #REF! : ignorée sans bloquer l’ouverture', async () => {
    const zip = await zipOf(SIMPLE);
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    zip.file('xl/worksheets/sheet1.xml', sheet.replace('</sheetData>', '</sheetData><dataValidations count="1"><dataValidation type="whole" sqref="#REF!"><formula1>1</formula1></dataValidation></dataValidations>'));
    await expect(importXlsx(await bytes(zip))).resolves.toBeTruthy();
  });

  it('nombre avec un lien : reste un nombre (pas l’adresse du lien)', async () => {
    const wb = book({ cellData: { 0: { 0: { v: 19551, t: 2, custom: { link: 'https://exemple.fr/dossier?id=19551' } } } } });
    const back = await importXlsx(await exportXlsx(wb));
    expect(first(back).cellData![0]![0]).toMatchObject({ v: 19551, t: 2, custom: { link: 'https://exemple.fr/dossier?id=19551' } });
  });

  it('en-têtes de tableau : texte exact gardé, saut de ligne écrit comme Excel', async () => {
    const tables = { s: { tables: [{ id: 't', name: 'Tableau1', range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }, options: {}, filters: { tableColumnFilterList: [] }, columns: [] }], tableFilteredOutRows: [] } };
    const wb = book(
      { cellData: { 0: { 0: { v: 'Genre ', t: 1 }, 1: { v: 'Note\nfinale', t: 1 } }, 1: { 0: { v: 'F', t: 1 }, 1: { v: 12, t: 2 } } } },
      [{ name: 'SHEET_TABLE_PLUGIN', data: JSON.stringify(tables) }],
    );
    const zip = await zipOf(wb);
    expect(await zip.file('xl/tables/table1.xml')!.async('string')).toContain('<tableColumn id="1" name="Genre "/><tableColumn id="2" name="Note_x000a_finale"/>');
    const back = await importXlsx(await bytes(zip));
    expect(first(back).cellData![0]![0]).toMatchObject({ v: 'Genre ' });
    const t = JSON.parse(back.resources!.find((r) => r.name === 'SHEET_TABLE_PLUGIN')!.data);
    expect(Object.values(t as Record<string, { tables: { columns: { displayName: string }[] }[] }>)[0]!.tables[0]!.columns.map((c) => c.displayName)).toEqual(['Genre ', 'Note\nfinale']);
  });

  it('milliers de cellules fusionnées : ouverture et enregistrement rapides', async () => {
    const mergeData = Array.from({ length: 20_000 }, (_, i) => ({ startRow: i, endRow: i, startColumn: 0, endColumn: 1 }));
    const cellData = Object.fromEntries(mergeData.map((m) => [m.startRow, { 0: { v: m.startRow, t: 2 } }]));
    const t0 = Date.now();
    const back = await importXlsx(await exportXlsx(book({ cellData, mergeData })));
    expect(Date.now() - t0).toBeLessThan(10_000);
    expect(first(back).mergeData).toHaveLength(20_000);
    expect(first(back).mergeData![123]).toEqual({ startRow: 123, endRow: 123, startColumn: 0, endColumn: 1 });
  });

  it('tableau croisé dont la source est en colonnes entières (A1:D1048576) : enregistré sans tout parcourir', async () => {
    const wb = pivotWorkbook({ ...PIVOT_DEF, sourceRange: 'A1:D1048576' });
    const zip = await zipOf(wb);
    expect(await zip.file('xl/pivotCache/pivotCacheRecords1.xml')!.async('string')).toContain('count="6"');
  });
});

describe('mises en forme conditionnelles d’Excel', () => {
  it('jeu d’icônes par défaut (sans nom) : les trois feux tricolores, comme Excel', async () => {
    const zip = await zipOf(SIMPLE);
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    const cf = '<conditionalFormatting sqref="B2"><cfRule type="iconSet" priority="1"><iconSet><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule></conditionalFormatting>';
    zip.file('xl/worksheets/sheet1.xml', sheet.replace(/(<pageMargins)/, `${cf}$1`));
    const wb = await importXlsx(await bytes(zip));
    const rules = JSON.parse(wb.resources!.find((r) => r.name === 'SHEET_CONDITIONAL_FORMATTING_PLUGIN')!.data);
    expect(JSON.stringify(rules)).toContain('"iconType":"3TrafficLights1"');
  });
});

describe('éléments d’Excel autrefois perdus à l’ouverture', () => {
  const COMMENTS =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Prof</author></authors><commentList>' +
    '<comment ref="D5" authorId="0"><text><r><rPr><b/></rPr><t>Prof:</t></r><r><t xml:space="preserve">\nÀ revoir &amp; compléter</t></r></text></comment></commentList></comments>';

  it('note sur une cellule vide : gardée', async () => {
    const zip = await zipOf(SIMPLE);
    zip.file('xl/comments1.xml', COMMENTS);
    const rels = (await zip.file('xl/worksheets/_rels/sheet1.xml.rels')?.async('string')) ?? '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    zip.file('xl/worksheets/_rels/sheet1.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/></Relationships>'));
    const wb = await importXlsx(await bytes(zip));
    const notes = JSON.parse(wb.resources!.find((r) => r.name === 'SHEET_NOTE_PLUGIN')!.data);
    expect(Object.values(notes as Record<string, Record<number, Record<number, { note: string }>>>)[0]![4]![3]!.note).toBe('Prof:\nÀ revoir & compléter');
  });

  it('feuille graphique : devient une feuille qui porte le graphique, à sa place', async () => {
    const zip = await zipOf(SIMPLE);
    const chart = '<?xml version="1.0" encoding="UTF-8"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>Feuil1!$B$2:$B$2</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>';
    zip.file('xl/charts/chart1.xml', chart);
    zip.file('xl/drawings/drawing1.xml', '<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="8568151" cy="5828082"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>');
    zip.file('xl/drawings/_rels/drawing1.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>');
    zip.file('xl/chartsheets/sheet1.xml', '<?xml version="1.0"?><chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetViews><sheetView workbookViewId="0"/></sheetViews><drawing r:id="rId1"/></chartsheet>');
    zip.file('xl/chartsheets/_rels/sheet1.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>');
    const wbXml = await zip.file('xl/workbook.xml')!.async('string');
    zip.file('xl/workbook.xml', wbXml.replace(/(<sheet [^>]*\/>)/, '<sheet name="Graphique" sheetId="9" r:id="rId99"/>$1'));
    const rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string');
    zip.file('xl/_rels/workbook.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/></Relationships>'));
    const wb = await importXlsx(await bytes(zip));
    expect(wb.sheetOrder.map((id) => wb.sheets[id]!.name)).toEqual(['Graphique', 'Feuil1']);
    const drawings = JSON.parse(wb.resources!.find((r) => r.name === 'SHEET_DRAWING_PLUGIN')!.data);
    expect(Object.values(drawings[wb.sheetOrder[0]!].data as Record<string, { componentKey: string }>)[0]!.componentKey).toBe('GrillonChart');
  });

  it('tableau croisé dont la source est un nom (tableau ou nom défini) : relu', async () => {
    const zip = await zipOf(pivotWorkbook());
    const cache = await zip.file('xl/pivotCache/pivotCacheDefinition1.xml')!.async('string');
    zip.file('xl/pivotCache/pivotCacheDefinition1.xml', cache.replace(/<worksheetSource [^>]*\/>/, '<worksheetSource name="Ventes2026"/>'));
    const wbXml = await zip.file('xl/workbook.xml')!.async('string');
    zip.file('xl/workbook.xml', wbXml.replace(/<\/sheets>/, '</sheets><definedNames><definedName name="Ventes2026">Ventes!$A$1:$D$7</definedName></definedNames>'));
    const wb = await importXlsx(await bytes(zip));
    const pivot = Object.values(wb.sheets).map((s) => (s.custom as { grillon?: { pivot?: { sourceSheet: string; sourceRange: string } } } | undefined)?.grillon?.pivot).find(Boolean);
    expect(pivot).toMatchObject({ sourceSheet: 'Ventes', sourceRange: 'A1:D7' });
  });
});

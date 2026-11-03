import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { detectLosses, type LossKind } from '../../src/io/xlsx/losses';

/** Crée une archive .xlsx en mémoire à partir d'un dictionnaire de fichiers. */
async function makeZip(files: Record<string, string | Uint8Array>): Promise<Uint8Array> {
  const zip = new JSZip();
  for (const [path, content] of Object.entries(files)) {
    zip.file(path, content);
  }
  return zip.generateAsync({ type: 'uint8array' });
}

describe('detectLosses', () => {
  it('renvoie un tableau vide pour des données non zip', async () => {
    expect(await detectLosses(new Uint8Array([0, 1, 2, 3]))).toEqual([]);
    expect(await detectLosses(new Uint8Array(0))).toEqual([]);
    expect(await detectLosses(new ArrayBuffer(16))).toEqual([]);
  });

  it('renvoie un tableau vide pour un classeur propre sans pertes', async () => {
    const data = await makeZip({
      'xl/workbook.xml': '<workbook><sheets><sheet name="Feuil1" sheetId="1"/></sheets></workbook>',
      'xl/worksheets/sheet1.xml': '<worksheet><sheetData/></worksheet>',
    });
    expect(await detectLosses(data)).toEqual([]);
  });

  it('détecte les macros VBA (xl/vbaProject.bin)', async () => {
    const data = await makeZip({
      'xl/vbaProject.bin': new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]),
    });
    expect(await detectLosses(data)).toEqual(['macros']);

    // Insensible à la casse
    const dataUpper = await makeZip({
      'XL/VBAPROJECT.BIN': new Uint8Array([0]),
    });
    expect(await detectLosses(dataUpper)).toEqual(['macros']);
  });

  it('détecte les commentaires hiérarchisés (xl/threadedComments/*.xml)', async () => {
    const data = await makeZip({
      'xl/threadedComments/threadedComment1.xml': '<ThreadedCommentList/>',
    });
    expect(await detectLosses(data)).toEqual(['threadedComments']);
  });

  it('détecte le SmartArt (xl/diagrams/*)', async () => {
    const data = await makeZip({
      'xl/diagrams/data1.xml': '<dataModel/>',
    });
    expect(await detectLosses(data)).toEqual(['smartArt']);
  });

  describe('graphiques non pris en charge (unsupportedCharts)', () => {
    it('détecte les graphiques 3D, surfaces et boursiers dans chart*.xml', async () => {
      const kinds = [
        'bar3DChart',
        'line3DChart',
        'pie3DChart',
        'area3DChart',
        'surface3DChart',
        'surfaceChart',
        'stockChart',
      ];
      for (const kind of kinds) {
        const data = await makeZip({
          'xl/charts/chart1.xml': `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:${kind}/></c:plotArea></c:chart></c:chartSpace>`,
        });
        expect(await detectLosses(data)).toEqual(['unsupportedCharts']);
      }
    });

    it('ignore les graphiques chartEx pris en charge seuls (cascade, entonnoir, colonnes, boîte à moustaches)', async () => {
      const supported = ['waterfall', 'funnel', 'clusteredColumn', 'boxWhisker'];
      for (const layoutId of supported) {
        const data = await makeZip({
          'xl/charts/chartEx1.xml': `<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart><cx:plotArea><cx:series layoutId="${layoutId}"/></cx:plotArea></cx:chart></cx:chartSpace>`,
        });
        expect(await detectLosses(data)).toEqual([]);
      }
    });

    it('détecte les graphiques chartEx Pareto', async () => {
      const data = await makeZip({
        'xl/charts/chartEx1.xml':
          '<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">' +
          '<cx:chart><cx:plotArea>' +
          '<cx:series layoutId="clusteredColumn"/>' +
          '<cx:series layoutId="paretoLine"/>' +
          '</cx:plotArea></cx:chart></cx:chartSpace>',
      });
      expect(await detectLosses(data)).toEqual(['unsupportedCharts']);
    });

    it('détecte les autres graphiques chartEx non pris en charge (treemap, sunburst, regionMap)', async () => {
      for (const layoutId of ['treemap', 'sunburst', 'regionMap']) {
        const data = await makeZip({
          'xl/charts/chartEx1.xml': `<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart><cx:plotArea><cx:series layoutId="${layoutId}"/></cx:plotArea></cx:chart></cx:chartSpace>`,
        });
        expect(await detectLosses(data)).toEqual(['unsupportedCharts']);
      }
    });
  });

  describe('noms définis avec formules (formulaNames)', () => {
    it('ignore les plages simples', async () => {
      const validCases = [
        '<definedName name="ZoneA">Feuil1!$A$1</definedName>',
        '<definedName name="ZoneB">Feuil1!$A$1:$B$2</definedName>',
        '<definedName name="ZoneC">A1</definedName>',
        '<definedName name="ZoneD">A1:B2</definedName>',
        '<definedName name="ZoneE">A:B</definedName>',
        '<definedName name="ZoneF">$A:$B</definedName>',
        '<definedName name="ZoneG">1:2</definedName>',
        '<definedName name="ZoneH">$1:$2</definedName>',
        '<definedName name="ZoneI">Feuil1!A:B</definedName>',
        '<definedName name="ZoneJ">Feuil1!1:2</definedName>',
      ];
      for (const xmlSnippet of validCases) {
        const data = await makeZip({
          'xl/workbook.xml': `<workbook><definedNames>${xmlSnippet}</definedNames></workbook>`,
        });
        expect(await detectLosses(data)).toEqual([]);
      }
    });

    it('ignore les plages simples sur plusieurs zones', async () => {
      const data = await makeZip({
        'xl/workbook.xml':
          '<workbook><definedNames>' +
          '<definedName name="Multi">Feuil1!$A$1,Feuil2!$B$2</definedName>' +
          '<definedName name="MultiEsp">Feuil1!A1:B2, Feuil1!D1:E2</definedName>' +
          '</definedNames></workbook>',
      });
      expect(await detectLosses(data)).toEqual([]);
    });

    it('ignore les noms de feuilles entre apostrophes (y compris avec apostrophe doublée)', async () => {
      const data = await makeZip({
        'xl/workbook.xml':
          '<workbook><definedNames>' +
          "<definedName name=\"NomEspaces\">'Ma feuille'!$A$1</definedName>" +
          "<definedName name=\"NomQuote\">'L''été'!$A$1:$B$2</definedName>" +
          '</definedNames></workbook>',
      });
      expect(await detectLosses(data)).toEqual([]);
    });

    it('ignore les noms réservés (_xlnm.)', async () => {
      const data = await makeZip({
        'xl/workbook.xml':
          '<workbook><definedNames>' +
          '<definedName name="_xlnm.Print_Area">OFFSET(Feuil1!$A$1, 0, 0, 10, 1)</definedName>' +
          '<definedName name="_xlnm._FilterDatabase">Feuil1!$A$1:$Z$100</definedName>' +
          '</definedNames></workbook>',
      });
      expect(await detectLosses(data)).toEqual([]);
    });

    it('ignore les noms masqués (hidden="1")', async () => {
      const data = await makeZip({
        'xl/workbook.xml':
          '<workbook><definedNames>' +
          '<definedName name="Masque" hidden="1">0.2</definedName>' +
          '</definedNames></workbook>',
      });
      expect(await detectLosses(data)).toEqual([]);
    });

    it('ignore les références en erreur (#REF!)', async () => {
      const data = await makeZip({
        'xl/workbook.xml':
          '<workbook><definedNames>' +
          '<definedName name="Erreur1">#REF!</definedName>' +
          '<definedName name="Erreur2">Feuil1!#REF!</definedName>' +
          '</definedNames></workbook>',
      });
      expect(await detectLosses(data)).toEqual([]);
    });

    it('détecte les formules, constantes et textes', async () => {
      const formulaData = await makeZip({
        'xl/workbook.xml':
          '<workbook><definedNames>' +
          '<definedName name="Calcul">OFFSET(Feuil1!$A$1, 0, 0, 10, 1)</definedName>' +
          '</definedNames></workbook>',
      });
      expect(await detectLosses(formulaData)).toEqual(['formulaNames']);

      const numberData = await makeZip({
        'xl/workbook.xml':
          '<workbook><definedNames>' +
          '<definedName name="Taux">0.2</definedName>' +
          '</definedNames></workbook>',
      });
      expect(await detectLosses(numberData)).toEqual(['formulaNames']);

      const textData = await makeZip({
        'xl/workbook.xml':
          '<workbook><definedNames>' +
          '<definedName name="Titre">"texte"</definedName>' +
          '</definedNames></workbook>',
      });
      expect(await detectLosses(textData)).toEqual(['formulaNames']);
    });
  });

  describe('filtres avancés (advancedFilters)', () => {
    it('détecte les éléments colorFilter, top10, dynamicFilter, iconFilter', async () => {
      const snippets = [
        '<worksheet><autoFilter ref="A1:C10"><filterColumn colId="0"><colorFilter cellColor="1"/></filterColumn></autoFilter></worksheet>',
        '<worksheet><autoFilter ref="A1:C10"><filterColumn colId="0"><top10 val="10"/></filterColumn></autoFilter></worksheet>',
        '<worksheet><autoFilter ref="A1:C10"><filterColumn colId="0"><dynamicFilter type="today"/></filterColumn></autoFilter></worksheet>',
        '<worksheet><autoFilter ref="A1:C10"><filterColumn colId="0"><iconFilter iconSet="3TrafficLights1" iconId="0"/></filterColumn></autoFilter></worksheet>',
      ];
      for (const xml of snippets) {
        const data = await makeZip({
          'xl/worksheets/sheet1.xml': xml,
        });
        expect(await detectLosses(data)).toEqual(['advancedFilters']);
      }
    });

    it('ignore les filtres automatiques ordinaires', async () => {
      const xml =
        '<worksheet><autoFilter ref="A1:C10"><filterColumn colId="0"><filters><filter val="Paris"/></filters></filterColumn></autoFilter></worksheet>';
      const data = await makeZip({
        'xl/worksheets/sheet1.xml': xml,
      });
      expect(await detectLosses(data)).toEqual([]);
    });
  });

  describe('tableaux croisés multiples (extraPivots)', () => {
    it('ne détecte pas un seul tableau croisé par feuille', async () => {
      const data = await makeZip({
        'xl/worksheets/_rels/sheet1.xml.rels':
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>' +
          '</Relationships>',
        'xl/worksheets/_rels/sheet2.xml.rels':
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable2.xml"/>' +
          '</Relationships>',
      });
      expect(await detectLosses(data)).toEqual([]);
    });

    it('détecte deux tableaux croisés sur une même feuille', async () => {
      const data = await makeZip({
        'xl/worksheets/_rels/sheet1.xml.rels':
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>' +
          '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable2.xml"/>' +
          '</Relationships>',
      });
      expect(await detectLosses(data)).toEqual(['extraPivots']);
    });
  });

  describe('chronologies Excel (excelTimelines)', () => {
    it('détecte les fichiers timelines et timelineCaches', async () => {
      const data1 = await makeZip({
        'xl/timelines/timeline1.xml': '<timeline/>',
      });
      expect(await detectLosses(data1)).toEqual(['excelTimelines']);

      const data2 = await makeZip({
        'xl/timelineCaches/timelineCacheDefinition1.xml': '<timelineCacheDefinition/>',
      });
      expect(await detectLosses(data2)).toEqual(['excelTimelines']);
    });
  });

  describe('ordre et unicité', () => {
    it('respecte l ordre officiel et l unicité avec plusieurs pertes', async () => {
      const expectedOrder: LossKind[] = [
        'macros',
        'threadedComments',
        'smartArt',
        'unsupportedCharts',
        'formulaNames',
        'advancedFilters',
        'extraPivots',
        'excelTimelines',
      ];

      const dataAll = await makeZip({
        'xl/timelines/timeline1.xml': '<timeline/>',
        'xl/timelineCaches/cache1.xml': '<cache/>',
        'xl/worksheets/_rels/sheet1.xml.rels':
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>' +
          '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable2.xml"/>' +
          '</Relationships>',
        'xl/worksheets/sheet1.xml': '<worksheet><colorFilter/></worksheet>',
        'xl/worksheets/sheet2.xml': '<worksheet><top10 val="5"/></worksheet>',
        'xl/workbook.xml': '<workbook><definedNames><definedName name="TVA">0.2</definedName></definedNames></workbook>',
        'xl/charts/chart1.xml': '<c:bar3DChart/>',
        'xl/charts/chart2.xml': '<c:pie3DChart/>',
        'xl/diagrams/data1.xml': '<data/>',
        'xl/threadedComments/threadedComment1.xml': '<comments/>',
        'xl/vbaProject.bin': new Uint8Array([1, 2, 3]),
      });

      const losses = await detectLosses(dataAll);
      expect(losses).toEqual(expectedOrder);
      // Vérification que chaque type apparaît au plus une fois
      expect(new Set(losses).size).toBe(losses.length);
    });

    it('ordonne correctement un sous-ensemble de pertes', async () => {
      const dataSubset = await makeZip({
        'xl/timelines/timeline1.xml': '<timeline/>',
        'xl/diagrams/data1.xml': '<data/>',
        'xl/vbaProject.bin': new Uint8Array([1, 2]),
      });
      expect(await detectLosses(dataSubset)).toEqual(['macros', 'smartArt', 'excelTimelines']);
    });
  });
});

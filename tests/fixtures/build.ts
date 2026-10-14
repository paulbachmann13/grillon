/**
 * Génère les classeurs .xlsx de test avec ExcelJS.
 * Utilisé par les tests Vitest (en mémoire) et par scripts/make-fixtures.ts (sur disque, pour Playwright).
 */
import ExcelJS from 'exceljs';
import JSZip from 'jszip';

export type FixtureName = 'styles' | 'formats' | 'formulas' | 'layout' | 'conditional' | 'data' | 'large';

const thin = { style: 'thin' as const, color: { argb: 'FF000000' } };

export async function buildStyles(): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Styles');
  ws.getCell('A1').value = 'Gras';
  ws.getCell('A1').font = { name: 'Calibri', size: 14, bold: true };
  ws.getCell('A2').value = 'Italique souligné';
  ws.getCell('A2').font = { name: 'Arial', size: 10, italic: true, underline: true };
  ws.getCell('A3').value = 'Barré rouge';
  ws.getCell('A3').font = { strike: true, color: { argb: 'FFFF0000' } };
  ws.getCell('A4').value = 'Fond jaune';
  ws.getCell('A4').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
  ws.getCell('A5').value = 'Bordures';
  ws.getCell('A5').border = { top: thin, left: thin, bottom: { style: 'double', color: { argb: 'FF0000FF' } }, right: { style: 'medium' } };
  ws.getCell('B1').value = 'Centré';
  ws.getCell('B1').alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getCell('B2').value = 'Un texte assez long qui doit revenir à la ligne automatiquement';
  ws.getCell('B2').alignment = { wrapText: true, vertical: 'top' };
  ws.getCell('B3').value = 'Droite';
  ws.getCell('B3').alignment = { horizontal: 'right' };
  ws.getCell('B4').value = 'Thème accent 1';
  ws.getCell('B4').font = { color: { theme: 4 } };
  ws.getCell('B5').value = 'Retrait';
  ws.getCell('B5').alignment = { indent: 2 };
  ws.getCell('C1').value = 'Rotation';
  ws.getCell('C1').alignment = { textRotation: 45 };
  // Cellule vide mais stylée (fond) : doit survivre.
  ws.getCell('C3').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00B050' } };
  ws.getColumn(2).width = 30;
  // Texte enrichi : plusieurs styles dans une cellule, avec un saut de ligne.
  ws.getCell('A7').value = {
    richText: [
      { text: 'Gras ', font: { bold: true, size: 12, name: 'Calibri' } },
      { text: 'rouge', font: { color: { argb: 'FFFF0000' }, size: 12, name: 'Calibri' } },
      { text: ' normal\nligne 2' },
    ],
  };
  ws.getCell('A7').alignment = { wrapText: true };
  return wb;
}

export async function buildFormats(): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Formats');
  const rows: [string, ExcelJS.CellValue, string?][] = [
    ['Entier', 42],
    ['Décimal', 3.14159, '0.00'],
    ['Pourcentage', 0.256, '0.0%'],
    ['Monétaire', 1234.5, '#,##0.00 "€"'],
    ['Date', new Date(Date.UTC(2026, 9, 14)), 'dd/mm/yyyy'],
    ['Date et heure', new Date(Date.UTC(2026, 0, 31, 13, 45, 0)), 'dd/mm/yyyy hh:mm'],
    ['Heure', 0.5, 'hh:mm:ss'],
    ['Scientifique', 123456789, '0.00E+00'],
    ['Booléen vrai', true],
    ['Booléen faux', false],
    ['Texte', 'Élève à l’école, garçon, œuvre, ÇA'],
    ['Texte numérique', '00123'],
    ['Négatif', -17.5, '#,##0.00;[Red]-#,##0.00'],
    ['Fraction', 0.75, '# ?/?'],
  ];
  rows.forEach(([label, value, fmt], i) => {
    ws.getCell(i + 1, 1).value = label;
    const c = ws.getCell(i + 1, 2);
    c.value = value;
    if (fmt) c.numFmt = fmt;
  });
  ws.getColumn(1).width = 18;
  ws.getColumn(2).width = 22;
  return wb;
}

export async function buildFormulas(): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  const data = wb.addWorksheet('Données');
  data.addRow(['Produit', 'Prix', 'Quantité', 'Total']);
  const items: [string, number, number][] = [
    ['Pomme', 0.5, 10],
    ['Poire', 0.8, 5],
    ['Kiwi', 0.3, 20],
    ['Mangue', 1.5, 2],
  ];
  items.forEach(([p, prix, q], i) => {
    const r = i + 2;
    data.addRow([p, prix, q]);
    data.getCell(`D${r}`).value = { formula: `B${r}*C${r}`, result: prix * q };
  });
  data.getCell('A7').value = 'Somme';
  data.getCell('D7').value = { formula: 'SUM(D2:D5)', result: 18 };
  data.getCell('A8').value = 'Moyenne';
  data.getCell('D8').value = { formula: 'AVERAGE(D2:D5)', result: 4.5 };
  data.getCell('A9').value = 'Test';
  data.getCell('D9').value = { formula: 'IF(D7>10,"Gros","Petit")', result: 'Gros' };
  data.getCell('A10').value = 'Recherche';
  data.getCell('D10').value = { formula: 'VLOOKUP("Kiwi",A2:D5,4,FALSE)', result: 6 };
  data.getCell('A11').value = 'Erreur';
  data.getCell('D11').value = { formula: '1/0', result: { error: '#DIV/0!' } };
  wb.definedNames.add("'Données'!$D$2:$D$5", 'Totaux');

  const synth = wb.addWorksheet('Synthèse');
  synth.getCell('A1').value = 'Total général';
  synth.getCell('B1').value = { formula: "SUM('Données'!D2:D5)", result: 18 };
  synth.getCell('A2').value = 'Via nom défini';
  synth.getCell('B2').value = { formula: 'SUM(Totaux)', result: 18 };
  synth.getCell('A3').value = 'SUMIFS';
  synth.getCell('B3').value = { formula: 'SUMIFS(\'Données\'!D2:D5,\'Données\'!C2:C5,">4")', result: 15 };
  return wb;
}

export async function buildLayout(): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  const a = wb.addWorksheet('Mise en page', {
    views: [{ state: 'frozen', xSplit: 1, ySplit: 2 }],
    properties: { tabColor: { argb: 'FF4472C4' } } as Partial<ExcelJS.WorksheetProperties>,
  });
  a.getCell('A1').value = 'Titre fusionné';
  a.mergeCells('A1:D1');
  a.getCell('A1').alignment = { horizontal: 'center' };
  a.getCell('B3').value = 'Bloc';
  a.mergeCells('B3:C5');
  a.getColumn(1).width = 25;
  a.getColumn(3).width = 5;
  a.getColumn(5).hidden = true;
  a.getRow(2).height = 40;
  a.getCell('A6').value = 'Ligne masquée';
  a.getRow(6).hidden = true;
  for (let r = 7; r <= 30; r++) a.getCell(r, 1).value = `Ligne ${r}`;

  const b = wb.addWorksheet('Deuxième feuille');
  b.getCell('A1').value = 'Feuille 2';
  const c = wb.addWorksheet('Masquée', { state: 'hidden' });
  c.getCell('A1').value = 'Cachée';
  const d = wb.addWorksheet('Sans quadrillage', { views: [{ showGridLines: false }] });
  d.getCell('A1').value = 'Pas de grille';
  return wb;
}

export async function buildConditional(): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('MFC');
  ws.addRow(['Note', 'Nom', 'Date', 'Barre', 'Échelle', 'Icônes']);
  const noms = ['Alice', 'Bob', 'alice', 'Chloé', 'Bob', 'David', 'Émile', 'Fanny', 'Gaël', 'Hugo'];
  for (let i = 0; i < 10; i++) {
    ws.addRow([i * 2, noms[i], new Date(Date.UTC(2026, 9, 10 + i)), i * 10, i, i * 11]);
  }
  ws.getColumn(3).numFmt = 'dd/mm/yyyy';
  const red = { font: { color: { argb: 'FF9C0006' } }, fill: { type: 'pattern' as const, pattern: 'solid' as const, bgColor: { argb: 'FFFFC7CE' } } };
  const green = { font: { color: { argb: 'FF006100' } }, fill: { type: 'pattern' as const, pattern: 'solid' as const, bgColor: { argb: 'FFC6EFCE' } } };
  ws.addConditionalFormatting({ ref: 'A2:A11', rules: [
    { type: 'cellIs', operator: 'greaterThan', formulae: ['10'], priority: 1, style: red },
    { type: 'cellIs', operator: 'between', formulae: ['2', '6'], priority: 2, style: green },
  ] });
  ws.addConditionalFormatting({ ref: 'A2:A11', rules: [{ type: 'top10', rank: 3, percent: false, bottom: false, priority: 3, style: { font: { bold: true } } }] });
  ws.addConditionalFormatting({ ref: 'A2:A11', rules: [{ type: 'aboveAverage', aboveAverage: false, priority: 4, style: { font: { italic: true } } }] });
  ws.addConditionalFormatting({ ref: 'B2:B11', rules: [
    { type: 'containsText', operator: 'containsText', text: 'li', priority: 5, style: red } as ExcelJS.ConditionalFormattingRule,
    { type: 'expression', formulae: ['COUNTIF($B$2:$B$11,B2)>1'], priority: 6, style: { font: { underline: true } } },
    { type: 'expression', formulae: ['LEFT(B2,LEN("G"))="G"'], priority: 7, style: green },
  ] });
  ws.addConditionalFormatting({ ref: 'C2:C11', rules: [{ type: 'timePeriod', timePeriod: 'today', priority: 8, style: green } as ExcelJS.ConditionalFormattingRule] });
  ws.addConditionalFormatting({ ref: 'D2:D11', rules: [{ type: 'dataBar', priority: 9, cfvo: [{ type: 'min' }, { type: 'max' }], color: { argb: 'FF638EC6' }, gradient: true } as unknown as ExcelJS.ConditionalFormattingRule] });
  ws.addConditionalFormatting({ ref: 'E2:E11', rules: [{ type: 'colorScale', priority: 10, cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }], color: [{ argb: 'FFF8696B' }, { argb: 'FFFFEB84' }, { argb: 'FF63BE7B' }] } as unknown as ExcelJS.ConditionalFormattingRule] });
  ws.addConditionalFormatting({ ref: 'F2:F11', rules: [{ type: 'iconSet', priority: 11, iconSet: '3Arrows', cfvo: [{ type: 'percent', value: 0 }, { type: 'percent', value: 33 }, { type: 'percent', value: 67 }] } as unknown as ExcelJS.ConditionalFormattingRule] });
  ws.addConditionalFormatting({ ref: 'A13:A20 C13:C20', rules: [{ type: 'expression', formulae: ['MOD(ROW(),2)=0'], priority: 12, style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFDDEBF7' } } } }] });
  return wb;
}

export async function buildLarge(rows = 50_000): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Gros volume');
  ws.addRow(['ID', 'Nom', 'Ville', 'Montant', 'Date', 'Actif']);
  ws.getRow(1).font = { bold: true };
  const villes = ['Lyon', 'Paris', 'Montréal', 'Genève', 'Québec', 'Nîmes'];
  const base = Date.UTC(2020, 0, 1);
  for (let i = 1; i <= rows; i++) {
    ws.addRow([i, `Élément ${i}`, villes[i % villes.length], Math.round(i * 13.37) / 100, new Date(base + (i % 2000) * 86_400_000), i % 3 === 0]);
  }
  ws.getColumn(5).numFmt = 'dd/mm/yyyy';
  ws.getCell(`G1`).value = 'Total';
  ws.getCell(`G2`).value = { formula: `SUM(D2:D${rows + 1})`, result: 0 };
  return wb;
}

/**
 * Validations, filtre automatique avec critères et tableau structuré.
 * ExcelJS n'écrit pas les critères de filtre : on les ajoute dans le XML, comme le ferait Excel.
 */
export async function buildData(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Données');
  ws.addRow(['Nom', 'Ville', 'Montant', 'Date', 'Code', 'Catégorie']);
  const villes = ['Lyon', 'Paris', 'Nîmes'];
  for (let i = 0; i < 10; i++) {
    ws.addRow([`Client ${i + 1}`, villes[i % 3], (i + 1) * 10, new Date(Date.UTC(2026, 0, 1 + i)), `C${i}`, i % 2 ? 'A' : 'B']);
  }
  ws.getColumn(4).numFmt = 'dd/mm/yyyy';
  ws.getCell('H1').value = 'A';
  ws.getCell('H2').value = 'B';
  ws.getCell('H3').value = 'C';
  const dv = (ws as unknown as { dataValidations: { add(a: string, v: ExcelJS.DataValidation): void } }).dataValidations;
  dv.add('B2:B11', { type: 'list', allowBlank: true, formulae: ['"Lyon,Paris,Nîmes"'], showErrorMessage: true, errorTitle: 'Ville', error: 'Choisir une ville de la liste' });
  dv.add('C2:C11', {
    type: 'whole',
    operator: 'between',
    formulae: [0, 1000],
    showInputMessage: true,
    promptTitle: 'Montant',
    prompt: 'Entier entre 0 et 1000',
    showErrorMessage: true,
    errorStyle: 'warning',
    error: 'Hors limites',
  });
  dv.add('D2:D11', { type: 'date', operator: 'greaterThanOrEqual', formulae: [new Date(Date.UTC(2026, 0, 1))], showErrorMessage: true });
  dv.add('E2:E11', { type: 'textLength', operator: 'lessThanOrEqual', formulae: [5] });
  dv.add('F2:F11', { type: 'list', formulae: ['$H$1:$H$3'] });
  dv.add('A2:A11', { type: 'custom', formulae: ['LEN(A2)>3'] });
  // Validation sur une colonne entière (ExcelJS la déplierait en un million de cellules).
  dv.add('J1:J1048576', { type: 'decimal', operator: 'greaterThan', formulae: [0] });
  ws.autoFilter = 'A1:F11';

  const t = wb.addWorksheet('Tableau');
  t.addTable({
    name: 'Ventes',
    ref: 'A1',
    headerRow: true,
    totalsRow: true,
    style: { theme: 'TableStyleMedium9', showRowStripes: true },
    columns: [
      { name: 'Produit', totalsRowLabel: 'Total', filterButton: true },
      { name: 'Qté', totalsRowFunction: 'sum', filterButton: true },
      { name: 'Prix', totalsRowFunction: 'average', filterButton: true },
    ],
    rows: [
      ['Café', 2, 3.5],
      ['Thé', 5, 2],
      ['Jus', 1, 4],
    ],
  } as ExcelJS.TableProperties);
  t.getCell('E1').value = 'Somme des prix';
  t.getCell('F1').value = { formula: 'SUM(Ventes[Prix])', result: 9.5 };

  const bytes = new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
  // Critères de filtre : Ville ∈ {Lyon, Paris} et Montant > 20 ; les lignes exclues sont masquées.
  const zip = await JSZip.loadAsync(bytes);
  const path = 'xl/worksheets/sheet1.xml';
  let xml = await zip.file(path)!.async('string');
  xml = xml.replace(
    /<autoFilter ref="A1:F11"\s*\/>/,
    '<autoFilter ref="A1:F11"><filterColumn colId="1"><filters><filter val="Lyon"/><filter val="Paris"/></filters></filterColumn>' +
      '<filterColumn colId="2"><customFilters><customFilter operator="greaterThan" val="20"/></customFilters></filterColumn></autoFilter>',
  );
  for (const row of [2, 3, 4, 7, 10]) xml = xml.replace(new RegExp(`<row r="${row}"`), `<row r="${row}" hidden="1"`);
  zip.file(path, xml);
  return zip.generateAsync({ type: 'uint8array' });
}

export async function buildFixture(name: FixtureName): Promise<Uint8Array> {
  if (name === 'data') return buildData();
  const builders: Record<Exclude<FixtureName, 'data'>, () => Promise<ExcelJS.Workbook>> = {
    styles: buildStyles,
    formats: buildFormats,
    formulas: buildFormulas,
    layout: buildLayout,
    conditional: buildConditional,
    large: () => buildLarge(),
  };
  const wb = await builders[name]();
  return new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
}

export const FIXTURES: FixtureName[] = ['styles', 'formats', 'formulas', 'layout', 'conditional', 'data', 'large'];

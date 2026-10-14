/**
 * Génère les classeurs .xlsx de test avec ExcelJS.
 * Utilisé par les tests Vitest (en mémoire) et par scripts/make-fixtures.ts (sur disque, pour Playwright).
 */
import ExcelJS from 'exceljs';

export type FixtureName = 'styles' | 'formats' | 'formulas' | 'layout' | 'large';

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

export async function buildFixture(name: FixtureName): Promise<Uint8Array> {
  const builders: Record<FixtureName, () => Promise<ExcelJS.Workbook>> = {
    styles: buildStyles,
    formats: buildFormats,
    formulas: buildFormulas,
    layout: buildLayout,
    large: () => buildLarge(),
  };
  const wb = await builders[name]();
  return new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
}

export const FIXTURES: FixtureName[] = ['styles', 'formats', 'formulas', 'layout', 'large'];

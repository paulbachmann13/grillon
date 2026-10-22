import { expect, test } from '@playwright/test';
import { gotoApp } from './helpers';

// Pour chaque scénario : chaque formule et chaque valeur vues dans l'app doivent être identiques après enregistrement + réouverture.
const SCENARIOS: Record<string, string> = {
  'remplissage simple': ``,
  'suppression de la ligne de la 1re formule partagée': `ws.deleteRows(2, 1);`,
  'insertion d’une ligne au milieu': `ws.insertRowsBefore(3, 1);`,
  'tri décroissant de A1:B8': `await api.executeCommand('sheet.command.sort-range', { unitId: wb.getId(), subUnitId: ws.getSheetId(), range: { startRow: 0, endRow: 7, startColumn: 0, endColumn: 1 }, orderRules: [{ colIndex: 0, type: 'desc' }], hasTitle: false });`,
  'effacement de la 1re cellule partagée': `ws.getRange('B2').clearContent();`,
  'déplacement de colonnes': `ws.moveColumns(ws.getRange('B1:B1'), 4);`,
  'formule matricielle débordante': `ws.getRange('F1').setFormula('=SEQUENCE(3)');`,
  'copie de la feuille': `wb.duplicateSheet(ws);`,
};

for (const [name, action] of Object.entries(SCENARIOS)) {
  test(`formules après « ${name} », enregistrement et réouverture`, async ({ page }) => {
    await gotoApp(page);
    const run = (b: string) => page.evaluate(`(async () => { const api = window.grillon.api; const wb = api.getActiveWorkbook(); const ws = wb.getActiveSheet(); ${b} })()`);
    await run(`
      ws.getRange('A1:A8').setValues([[5],[3],[8],[1],[9],[2],[7],[4]]);
      ws.getRange('B1').setFormula('=A1*10');
      ws.getRange('B1').autoFill(ws.getRange('B1:B8'), 'COPY');
      ws.getRange('C1').setFormula('=SUM($A$1:A1)');
      ws.getRange('C1').autoFill(ws.getRange('C1:C8'), 'COPY');
    `);
    await page.waitForTimeout(600);
    await run(action);
    await page.waitForTimeout(800);
    const read = `const out = []; for (const s of wb.getSheets()) { const lr = Math.max(s.getLastRow(), 0), lc = Math.max(s.getLastColumn(), 0); for (let r = 0; r <= lr; r++) for (let c = 0; c <= lc; c++) { const g = s.getRange(r, c); const f = g.getFormula(); const v = g.getDisplayValue(); if (f || v) out.push(s.getSheetName() + '!' + g.getA1Notation() + ' ' + (f || '') + ' = ' + v); } } return out;`;
    const before = (await run(read)) as string[];
    await page.evaluate(`window.grillon.doc.exportXlsxBytes().then((b) => window.__b = b)`);
    await page.evaluate(`window.grillon.doc.openBytes('audit.xlsx', window.__b)`);
    await page.waitForTimeout(1500);
    const after = (await run(read)) as string[];
    expect(after).toEqual(before);
    // Pas d'erreur : les formules recopiées et triées restent justes.
    expect(before.filter((x) => /#REF!|#NOM\?|#VALEUR!/.test(x))).toEqual([]);
  });
}

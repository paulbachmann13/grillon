import { expect, test, type Page } from '@playwright/test';
import { gotoApp } from './helpers';

const beta = (page: Page) => page.addInitScript(() => localStorage.setItem('grillon-beta', 'on'));
const run = <T = unknown>(page: Page, body: string) => page.evaluate(`(async () => { const api = window.grillon.api; const wb = api.getActiveWorkbook(); const ws = wb.getActiveSheet(); ${body} })()`) as Promise<T>;

async function fill(page: Page) {
  await run(page, `ws.getRange('A1:D4').setValues([['Nom','Ville','Note 1','Note 2'],['Léa','Lyon',15,12],['Tom','Paris',8,10],['Zoé','Lyon',19,17]]); wb.setActiveRange(ws.getRange('B3'));`);
}
const ribbonButton = async (page: Page, id: string) => {
  await page.getByRole('tab', { name: 'Formules', exact: true }).click();
  await page.locator(`[data-u-command="${id}"]`).first().click();
};

test('sans l’interrupteur bêta : aucun bouton « Demander à Claude »', async ({ page }) => {
  await gotoApp(page);
  await page.getByRole('tab', { name: 'Formules', exact: true }).click();
  await expect(page.locator('[data-u-command="grillon.beta.ask-claude"]')).toHaveCount(0);
  await page.getByRole('button', { name: 'Fichier' }).click();
  await expect(page.getByText('Demander à Claude')).toHaveCount(0);
});

test('Paramètres : l’interrupteur des fonctions bêta est mémorisé', async ({ page }) => {
  await gotoApp(page);
  await page.evaluate(`window.grillon.doc.showHome('settings')`);
  const box = page.getByTestId('setting-beta');
  await expect(box).not.toBeChecked();
  await box.check();
  expect(await page.evaluate(() => localStorage.getItem('grillon-beta'))).toBe('on');
  await box.uncheck();
  expect(await page.evaluate(() => localStorage.getItem('grillon-beta'))).toBeNull();
});

test('Demander à Claude : message prêt à coller, textes masqués, copie dans le presse-papiers', async ({ page, context }) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await beta(page);
  await gotoApp(page);
  await fill(page);
  await ribbonButton(page, 'grillon.beta.ask-claude');
  const dialog = page.getByTestId('dialog-ask-claude');
  await expect(dialog).toBeVisible();
  const preview = dialog.getByTestId('ask-preview');
  // Une seule cellule sélectionnée : tout le tableau autour.
  await expect(dialog.getByTestId('ask-summary')).toContainText('4 lignes et 4 colonnes');
  await expect(dialog.getByTestId('ask-header')).toBeChecked();
  await expect(dialog.getByTestId('ask-hide')).toBeChecked();
  await expect(preview).toHaveValue(/cellules A1:D4/);
  await dialog.getByTestId('ask-question').fill('Ajoute la moyenne des deux notes');
  await expect(preview).toHaveValue(/Ma demande : Ajoute la moyenne des deux notes/);
  // Textes masqués : ni les noms ni les villes ne sortent.
  const masked = await preview.inputValue();
  expect(masked).toContain('1\tNom\tVille\tNote 1\tNote 2');
  expect(masked).toContain('2\tTexte 1\tTexte 2\t15\t12');
  expect(masked).not.toContain('Léa');
  expect(masked).not.toContain('Lyon');
  await dialog.getByTestId('ask-hide').uncheck();
  await expect(preview).toHaveValue(/2\tLéa\tLyon\t15\t12/);
  // Sans question : on ne copie pas.
  await dialog.getByTestId('ask-question').fill('');
  await dialog.getByRole('button', { name: 'Copier le message' }).click();
  await expect(dialog.locator('.g-dialog-error')).toContainText('Écrivez votre question');
  await dialog.getByTestId('ask-question').fill('Ajoute la moyenne');
  const text = await preview.inputValue();
  await dialog.getByRole('button', { name: 'Copier le message' }).click();
  await expect(dialog).toBeHidden();
  expect(await page.evaluate(() => navigator.clipboard.readText())).toBe(text);
  expect(text).toContain('```grillon');
});

test('Coller la réponse : aperçu avant / après, une seule annulation, la feuille ne bouge pas avant « Appliquer »', async ({ page }) => {
  await beta(page);
  await gotoApp(page);
  await fill(page);
  await ribbonButton(page, 'grillon.beta.paste-reply');
  const dialog = page.getByTestId('dialog-paste-reply');
  await expect(dialog).toBeVisible();
  const apply = dialog.getByRole('button', { name: 'Appliquer' });
  await dialog.getByTestId('paste-reply').fill("Voici la moyenne :\n```grillon\nE1: Moyenne\nE2: =MOYENNE(C2;D2)\nE3: =MOYENNE(C3;D3)\nE4: =MOYENNE(C4;D4)\nF1: 12,5\nB2: Nîmes\nZZZZ0: x\n```\nBonne journée");
  const preview = dialog.getByTestId('paste-preview');
  await expect(preview.locator('tbody tr')).toHaveCount(6);
  await expect(preview).toContainText('B2');
  await expect(preview).toContainText('Lyon'); // avant
  await expect(preview).toContainText('Nîmes'); // après
  await expect(dialog).toContainText('6 cellules seront modifiées');
  await expect(dialog).toContainText('1 lignes ignorées');
  expect(await run(page, `return ws.getRange('E2').getValue();`)).toBeNull();
  await apply.click();
  await expect(dialog).toBeHidden();
  await expect.poll(() => run(page, `return ws.getRange('E2').getDisplayValue();`)).toBe('13,5');
  expect(await run(page, `return ws.getRange('E1').getValue();`)).toBe('Moyenne');
  expect(await run(page, `return ws.getRange('E4').getDisplayValue();`)).toBe('18');
  expect(await run(page, `return ws.getRange('F1').getDisplayValue();`)).toBe('12,5');
  // F1 est un vrai nombre (pas du texte).
  expect(await run(page, `const c = ws.getRange('F1').getCellDataGrid()[0][0]; return [typeof c.v, c.t];`)).toEqual(['number', 2]);
  expect(await run(page, `return ws.getRange('B2').getValue();`)).toBe('Nîmes');
  // La formule est une vraie formule (anglais stocké), pas du texte.
  expect(await run(page, `return ws.getRange('E2').getFormula();`)).toBe('=AVERAGE(C2,D2)');
  // Un seul Ctrl+Z défait tout.
  await run(page, `await api.executeCommand('univer.command.undo');`);
  await expect.poll(() => run(page, `return ws.getRange('E2').getValue();`)).toBeNull();
  expect(await run(page, `return ws.getRange('E1').getValue();`)).toBeNull();
  expect(await run(page, `return ws.getRange('B2').getValue();`)).toBe('Lyon');
});

test('Coller la réponse : texte sans bloc exploitable, rien à appliquer', async ({ page }) => {
  await beta(page);
  await gotoApp(page);
  await fill(page);
  await page.getByRole('button', { name: 'Fichier' }).click();
  await page.getByText('Coller la réponse').click();
  const dialog = page.getByTestId('dialog-paste-reply');
  await dialog.getByTestId('paste-reply').fill('Je n’ai pas bien compris, peux-tu préciser quelles colonnes ?');
  await expect(dialog).toContainText('Aucune cellule à remplir');
  await dialog.getByRole('button', { name: 'Appliquer' }).click();
  await expect(dialog.locator('.g-dialog-error')).toContainText('Aucune cellule');
  await expect(dialog).toBeVisible();
});

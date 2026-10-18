import { expect, test, type Page } from '@playwright/test';
import { join } from 'node:path';
import { FIXTURE_DIR, useFallbackFilePickers, waitReady } from './helpers';

// Ici on part « comme un vrai démarrage » : pas de réglage qui masque l'écran d'accueil.
async function start(page: Page) {
  await page.goto('./');
  await waitReady(page);
}
const home = (page: Page) => page.getByTestId('start-screen');

test('au démarrage sans rien à reprendre : écran d’accueil avec le numéro de version', async ({ page }) => {
  await start(page);
  await expect(home(page)).toBeVisible();
  await expect(home(page).getByRole('heading', { name: 'Bonjour' })).toBeVisible();
  await expect(home(page).getByTestId('start-new')).toBeVisible();
  await expect(home(page).getByTestId('start-open')).toBeVisible();
  await expect(home(page).getByTestId('start-csv')).toBeVisible();
  await expect(home(page).getByTestId('start-no-recent')).toBeVisible();
  const pkg = JSON.parse(await (await import('node:fs/promises')).readFile(join(import.meta.dirname, '..', '..', 'package.json'), 'utf8')) as { version: string };
  await expect(home(page).getByTestId('start-version')).toHaveText(`Version ${pkg.version}`);
});

test('« Classeur vide » ferme l’accueil et on peut travailler', async ({ page }) => {
  await start(page);
  await home(page).getByTestId('start-new').click();
  await expect(home(page)).toBeHidden();
  await page.evaluate(`(() => { const wb = window.grillon.api.getActiveWorkbook(); wb.getActiveSheet().getRange('A1').setValue('ok'); })()`);
  expect(await page.evaluate(`window.grillon.api.getActiveWorkbook().getActiveSheet().getRange('A1').getValue()`)).toBe('ok');
});

test('Ouvrir : le classeur s’ouvre et l’accueil disparaît', async ({ page }) => {
  await useFallbackFilePickers(page);
  await start(page);
  const chooser = page.waitForEvent('filechooser');
  await home(page).getByTestId('start-open').click();
  await (await chooser).setFiles(join(FIXTURE_DIR, 'data.xlsx'));
  await expect(page.getByTestId('doc-title')).toHaveText(/^data/, { timeout: 60_000 });
  await expect(home(page)).toBeHidden();
});

test('Fichier → Accueil : revient à l’accueil, « Retour au classeur » et Échap reviennent au classeur', async ({ page }) => {
  await start(page);
  await home(page).getByTestId('start-new').click();
  await page.getByTestId('file-menu').click();
  await page.getByTestId('file-home').click();
  await expect(home(page)).toBeVisible();
  await home(page).getByTestId('start-back').click();
  await expect(home(page)).toBeHidden();
  await page.getByTestId('file-menu').click();
  await page.getByTestId('file-home').click();
  await expect(home(page)).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(home(page)).toBeHidden();
});

test('Paramètres : apparence et numéro de version', async ({ page }) => {
  await start(page);
  await home(page).getByTestId('start-nav-settings').click();
  await expect(home(page).getByRole('heading', { name: 'À propos de Grillon' })).toBeVisible();
  await home(page).getByTestId('setting-theme-dark').check();
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');
  await home(page).getByTestId('setting-theme-system').check();
  await expect(page.locator('html')).not.toHaveAttribute('data-theme', 'dark');
});

test('Fichier → Paramètres ouvre les paramètres au-dessus du classeur', async ({ page }) => {
  await start(page);
  await home(page).getByTestId('start-new').click();
  await page.getByTestId('file-menu').click();
  await page.getByTestId('file-settings').click();
  await expect(home(page).getByRole('heading', { name: 'Paramètres' })).toBeVisible();
  await expect(home(page).getByTestId('start-back')).toBeVisible();
});

// Session précédente : on la fabrique avec le réglage des tests (qui masque l'accueil), puis on relance « comme un vrai démarrage ».
async function previousSession(page: Page, dirty: boolean) {
  await page.addInitScript(() => localStorage.setItem('grillon-no-home', '1'));
  await page.goto('./');
  await waitReady(page);
  if (dirty) await page.evaluate(`window.grillon.api.getActiveWorkbook().getActiveSheet().getRange('A1').setValue('à reprendre')`);
  await page.evaluate(`window.grillon.doc.autosaveNow()`);
  await page.addInitScript(() => localStorage.removeItem('grillon-no-home'));
  await page.reload();
  await waitReady(page);
}
const a1 = (page: Page) => page.evaluate(`window.grillon.api.getActiveWorkbook().getActiveSheet().getRange('A1').getValue()`);

test('travail non enregistré : on arrive sur l’accueil, qui propose de le reprendre', async ({ page }) => {
  await previousSession(page, true);
  await expect(home(page)).toBeVisible();
  // Pas rouvert d'office : le classeur derrière est vide.
  expect(await a1(page)).toBeNull();
  const draft = home(page).getByTestId('start-draft');
  await expect(draft).toContainText('Reprendre : Classeur1.xlsx');
  await expect(draft).toContainText('Travail non enregistré');
  await draft.getByRole('button').first().click();
  await expect(home(page)).toBeHidden();
  expect(await a1(page)).toBe('à reprendre');
  expect(await page.evaluate(`window.grillon.doc.dirty`)).toBe(true);
  // Une fois repris, le brouillon n'est plus proposé.
  await page.getByTestId('file-menu').click();
  await page.getByTestId('file-home').click();
  await expect(home(page).getByTestId('start-draft')).toHaveCount(0);
});

test('dernière session déjà enregistrée : accueil sans « Reprendre »', async ({ page }) => {
  await previousSession(page, false);
  await expect(home(page)).toBeVisible();
  await expect(home(page).getByTestId('start-draft')).toHaveCount(0);
});

test('supprimer le travail non enregistré : confirmation, puis il n’est plus proposé', async ({ page }) => {
  await previousSession(page, true);
  const draft = home(page).getByTestId('start-draft');
  await expect(draft).toBeVisible();
  page.once('dialog', (d) => void d.dismiss());
  await draft.getByTestId('start-discard-draft').click();
  await expect(draft).toBeVisible();
  page.once('dialog', (d) => void d.accept());
  await draft.getByTestId('start-discard-draft').click();
  await expect(draft).toHaveCount(0);
  await page.reload();
  await waitReady(page);
  await expect(home(page)).toBeVisible();
  await expect(home(page).getByTestId('start-draft')).toHaveCount(0);
});

test('fichiers récents : liste, retrait d’un fichier, fichier introuvable', async ({ page }) => {
  const now = Date.now();
  await page.addInitScript((at) => {
    localStorage.setItem('grillon-recent', JSON.stringify([
      { id: '/Users/maman/Documents/Budget/Budget 2026.xlsx', name: 'Budget 2026.xlsx', path: '/Users/maman/Documents/Budget/Budget 2026.xlsx', at },
      { id: '/Users/maman/Documents/Notes.xlsx', name: 'Notes.xlsx', path: '/Users/maman/Documents/Notes.xlsx', at: at - 3 * 86_400_000 },
    ]));
  }, now);
  await start(page);
  const rows = home(page).getByTestId('start-recent');
  await expect(rows).toHaveCount(2);
  await expect(rows.first()).toContainText('Budget 2026.xlsx');
  await expect(rows.first()).toContainText('Budget');
  await expect(rows.first()).toContainText('Aujourd’hui');
  await expect(rows.nth(1)).toContainText('Il y a 3 jours');

  // Fichier introuvable (dans le navigateur, le chemin n'existe pas) : message clair et retrait de la liste.
  await rows.first().getByRole('button').first().click();
  await expect(page.getByTestId('status')).toContainText('Impossible de rouvrir « Budget 2026.xlsx »');
  await expect(home(page).getByTestId('start-recent')).toHaveCount(1);

  // Retrait manuel.
  await home(page).getByTestId('start-forget').click();
  await expect(home(page).getByTestId('start-no-recent')).toBeVisible();
});

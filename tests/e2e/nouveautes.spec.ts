import { expect, test } from '@playwright/test';
import { gotoApp } from './helpers';


test('Fichier → Nouveautés : l’historique des versions, de la plus récente à la plus ancienne', async ({ page }) => {
  await gotoApp(page);
  await page.getByRole('button', { name: 'Fichier' }).click();
  await page.getByText('Nouveautés et historique des versions…').click();
  const dialog = page.getByTestId('dialog-history');
  await expect(dialog).toBeVisible();
  const versions = await dialog.locator('.g-release h3').allTextContents();
  expect(versions.length).toBeGreaterThan(10);
  expect(versions[0]).toMatch(/^Grillon \d+\.\d+\.\d+/);
  expect(versions.at(-1)).toContain('0.1.0');
  // Les notes sont du texte, jamais du HTML.
  expect(await dialog.locator('.g-notes script, .g-notes img').count()).toBe(0);
  await expect(dialog.locator('.g-release').first().locator('ul li').first()).toBeVisible();
  await dialog.getByRole('button', { name: 'Fermer' }).first().click();
  await expect(dialog).toBeHidden();
});

test('après une mise à jour : ce qui a changé est raconté une seule fois', async ({ page }) => {
  await gotoApp(page);
  await page.evaluate(() => localStorage.setItem('grillon-last-version', '1.13.0'));
  await page.evaluate(`void window.grillon.showWhatsNewIfUpdated(); 0`);
  const dialog = page.getByTestId('dialog-whats-new');
  await expect(dialog).toBeVisible();
  await expect(dialog.locator('h2')).toContainText('Nouveautés de Grillon depuis la version 1.13.0');
  const shown = await dialog.locator('.g-release h3').allTextContents();
  expect(shown.some((v) => v.includes('1.14.4'))).toBe(true);
  expect(shown.some((v) => v.includes('1.13.0'))).toBe(false);
  await dialog.getByRole('button', { name: 'Compris' }).click();
  await expect(dialog).toBeHidden();
  // Déjà racontée : plus de fenêtre.
  await page.evaluate(`void window.grillon.showWhatsNewIfUpdated(); 0`);
  await expect(page.getByTestId('dialog-whats-new')).toHaveCount(0);
});

test('« Voir tout l’historique » ouvre l’historique', async ({ page }) => {
  await gotoApp(page);
  await page.evaluate(() => localStorage.setItem('grillon-last-version', '1.13.0'));
  await page.evaluate(`void window.grillon.showWhatsNewIfUpdated(); 0`);
  await page.getByTestId('dialog-whats-new').getByRole('button', { name: 'Voir tout l’historique' }).click();
  await expect(page.getByTestId('dialog-history')).toBeVisible();
});

test('l’accueil : le numéro de version ouvre l’historique', async ({ page }) => {
  await gotoApp(page);
  await page.evaluate(`window.grillon.doc.showHome('home')`);
  await page.getByTestId('start-version').click();
  await expect(page.getByTestId('dialog-history')).toBeVisible();
});

test('installation neuve : aucune fenêtre de nouveautés ; copie déjà utilisée sans version mémorisée : la version en cours est racontée', async ({ page }) => {
  await gotoApp(page);
  await page.evaluate(() => localStorage.removeItem('grillon-last-version'));
  await page.evaluate(`void window.grillon.showWhatsNewIfUpdated(); 0`);
  await page.waitForTimeout(500);
  await expect(page.getByTestId('dialog-whats-new')).toHaveCount(0);
  await page.evaluate(() => {
    localStorage.removeItem('grillon-last-version');
    localStorage.setItem('grillon-recent', JSON.stringify([{ id: 'r1', name: 'a.xlsx', at: Date.now() }]));
  });
  await page.evaluate(`void window.grillon.showWhatsNewIfUpdated(); 0`);
  await expect(page.getByTestId('dialog-whats-new')).toBeVisible();
  await expect(page.getByTestId('dialog-whats-new').locator('.g-release')).toHaveCount(1);
});

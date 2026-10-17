import { expect, test } from '@playwright/test';

// Simule l'application de bureau : seules les commandes natives qui comptent sont observées.
test.beforeEach(async ({ page }) => {
  await page.addInitScript(() => {
    const calls: [string, unknown][] = [];
    (window as any).__calls = calls;
    (window as any).__TAURI_INTERNALS__ = {
      metadata: { currentWindow: { label: 'main' }, currentWebview: { windowLabel: 'main', label: 'main' } },
      transformCallback: () => 0,
      unregisterCallback: () => {},
      convertFileSrc: (p: string) => p,
      invoke: async (cmd: string, args: unknown) => {
        calls.push([cmd, args]);
        if (cmd === 'take_pending_opens') return [];
        return null;
      },
    };
  });
  await page.goto('/');
  await expect(page.locator('#univer canvas').first()).toBeVisible();
});

test('application de bureau : Ouvrir et Nouveau passent dans une nouvelle fenêtre quand le classeur est utilisé', async ({ page }) => {
  const calls = () => page.evaluate(() => (window as any).__calls.filter((c: any[]) => c[0] === 'new_window'));
  // Classeur vide et intact : un fichier le remplace (pas de nouvelle fenêtre).
  expect(await page.evaluate(() => (window as any).grillon.doc.isPristine())).toBe(true);
  expect(await page.evaluate(() => (window as any).grillon.doc.newInNewWindow())).toBe(false);
  // Classeur modifié : Nouveau ouvre une fenêtre vide, un fichier s'ouvre dans la sienne.
  await page.evaluate(() => {
    const ws = (window as any).grillon.api.getActiveWorkbook().getActiveSheet();
    ws.getRange('A1').setValue('x');
  });
  await expect.poll(() => page.evaluate(() => (window as any).grillon.doc.dirty)).toBe(true);
  expect(await page.evaluate(() => (window as any).grillon.doc.newInNewWindow())).toBe(true);
  await page.evaluate(async () => {
    // Un « handle » de l'appli de bureau : seul son chemin compte ici.
    await (window as any).grillon.doc.openHandle({ kind: 'file', name: 'Ventes.xlsx', path: '/Users/moi/Ventes.xlsx' });
  });
  expect(await calls()).toEqual([
    ['new_window', { path: null }],
    ['new_window', { path: '/Users/moi/Ventes.xlsx' }],
  ]);
  // Le classeur d'origine est resté tel quel.
  expect(await page.evaluate(() => (window as any).grillon.api.getActiveWorkbook().getActiveSheet().getRange('A1').getValue())).toBe('x');
});

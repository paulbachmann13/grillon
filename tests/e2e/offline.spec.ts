import { expect, test } from '@playwright/test';
import { gotoApp, waitReady } from './helpers';

test('PWA : manifeste valide et fonctionnement hors ligne', async ({ page, context }) => {
  await gotoApp(page);
  const manifest = await page.evaluate(async () => {
    const href = document.querySelector('link[rel="manifest"]')?.getAttribute('href');
    return href ? (await fetch(href)).json() : null;
  });
  expect(manifest).toMatchObject({ name: 'Grillon — tableur', lang: 'fr', display: 'standalone' });

  // Attendre que le service worker ait tout mis en cache et contrôle la page.
  await page.evaluate(async () => {
    await navigator.serviceWorker.ready;
  });
  await page.reload();
  await page.waitForFunction(() => !!navigator.serviceWorker.controller);

  await context.setOffline(true);
  await page.reload();
  await waitReady(page);
  await expect(page.getByText('Démarrer', { exact: true }).first()).toBeVisible();
  // Le convertisseur xlsx (chargé à la demande) doit aussi être disponible hors ligne.
  const ok = await page.evaluate(async () => (await (window as any).grillon.doc.exportXlsxBytes()).byteLength > 0);
  expect(ok).toBe(true);
  await context.setOffline(false);
});

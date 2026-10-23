import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';
import { join } from 'node:path';

export const FIXTURE_DIR = join(import.meta.dirname, '..', 'fixtures', 'out');

/**
 * Désactive la File System Access API pour passer par <input type=file> et les
 * téléchargements (chemin Safari/iOS), que Playwright sait piloter.
 */
export async function useFallbackFilePickers(page: Page): Promise<void> {
  await page.addInitScript(() => {
    delete (window as unknown as Record<string, unknown>).showOpenFilePicker;
    delete (window as unknown as Record<string, unknown>).showSaveFilePicker;
  });
}

export async function gotoApp(page: Page, fresh = true): Promise<void> {
  // Les tests travaillent directement sur le classeur vide : pas d'écran d'accueil (voir tests/e2e/home.spec.ts).
  await page.addInitScript(() => localStorage.setItem('grillon-no-home', '1'));
  if (fresh) {
    await page.goto('./');
    await page.evaluate(
      () =>
        new Promise<void>((resolve) => {
          const req = indexedDB.deleteDatabase('grillon');
          req.onsuccess = req.onerror = req.onblocked = () => resolve();
        }),
    );
  }
  await page.goto('./');
  await waitReady(page);
}

export async function waitReady(page: Page): Promise<void> {
  await page.waitForFunction(() => {
    const g = (window as unknown as { grillon?: { api: { getActiveWorkbook(): unknown } } }).grillon;
    return !!g?.api.getActiveWorkbook();
  });
  await expect(page.getByTestId('file-menu')).toBeVisible();
}

export async function openViaMenu(page: Page, file: string): Promise<void> {
  const chooser = page.waitForEvent('filechooser');
  await page.getByTestId('file-menu').click();
  await page.getByTestId('file-open').click();
  await (await chooser).setFiles(join(FIXTURE_DIR, file));
  const base = file.replace(/\.[^.]+$/, '');
  await expect(page.getByTestId('doc-title')).toHaveText(new RegExp(`^${base}`), { timeout: 60_000 });
}

/** Valeur affichée d'une cellule (feuille par nom, sinon feuille active). */
export async function display(page: Page, a1: string, sheet?: string): Promise<string> {
  return page.evaluate(
    ({ a1, sheet }) => {
      const g = (window as unknown as { grillon: { api: any } }).grillon;
      const wb = g.api.getActiveWorkbook();
      const ws = sheet ? wb.getSheetByName(sheet) : wb.getActiveSheet();
      return String(ws.getRange(a1).getDisplayValue());
    },
    { a1, sheet },
  );
}

/** Le tri et le filtre d'Univer enregistrent leurs commandes une fois la grille dessinée : on les attend. */
export async function waitSortFilterReady(page: Page): Promise<void> {
  await page.waitForFunction(() => {
    const g = (window as unknown as { grillon?: { hasCommand?(id: string): boolean } }).grillon;
    return !!g?.hasCommand?.('sheet.command.sort-range-asc') && !!g.hasCommand('sheet.command.smart-toggle-filter');
  });
}

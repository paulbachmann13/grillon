import JSZip from 'jszip';
import { expect, test } from '@playwright/test';
import { gotoApp } from './helpers';

/**
 * Un fichier Excel qui contient ce que Grillon ne sait pas garder (ici des macros et un commentaire en fil) :
 * Grillon prévient avant d'écraser l'original, et ne réécrit rien si l'on annule.
 */
test('avant d’écraser un fichier : liste de ce qui serait perdu ; Annuler n’écrit rien', async ({ page }) => {
  await gotoApp(page);
  const base = (await page.evaluate(`(async () => { const g = window.grillon; g.api.getActiveWorkbook().getActiveSheet().getRange('A1').setValue(42); return Array.from(await g.doc.exportXlsxBytes()); })()`)) as number[];
  const zip = await JSZip.loadAsync(new Uint8Array(base));
  zip.file('xl/vbaProject.bin', new Uint8Array([1, 2, 3]));
  zip.file('xl/threadedComments/threadedComment1.xml', '<ThreadedComments/>');
  const bytes = Array.from(await zip.generateAsync({ type: 'uint8array' }));

  await page.evaluate(async (b) => {
    const w = window as unknown as { __writes: number; grillon: { doc: { openBytes(n: string, b: Uint8Array, h?: unknown): Promise<void> } } };
    w.__writes = 0;
    // Faux fichier du disque : compte les écritures.
    const handle = { kind: 'file', name: 'macros.xlsx', createWritable: async () => ({ write: async () => void w.__writes++, close: async () => undefined }) };
    await w.grillon.doc.openBytes('macros.xlsx', new Uint8Array(b), handle);
  }, bytes);
  await expect(page.getByText('Grillon vous préviendra avant de l’écraser')).toBeVisible();

  await page.evaluate(`void window.grillon.doc.save(); 0`);
  const dialog = page.getByTestId('dialog-losses');
  await expect(dialog).toBeVisible();
  await expect(dialog).toContainText('macros (VBA)');
  await expect(dialog).toContainText('commentaires en fil');
  await dialog.getByRole('button', { name: 'Annuler' }).click();
  await expect(dialog).toBeHidden();
  expect(await page.evaluate(`window.__writes`)).toBe(0);

  await page.evaluate(`void window.grillon.doc.save(); 0`);
  await page.getByTestId('dialog-losses').getByRole('button', { name: 'Enregistrer quand même' }).click();
  await expect.poll(() => page.evaluate(`window.__writes`)).toBe(1);
  // Le fichier ne contient plus ces éléments : plus de question à l'enregistrement suivant.
  await page.evaluate(`window.grillon.doc.save()`);
  await expect(page.getByTestId('dialog-losses')).toHaveCount(0);
  expect(await page.evaluate(`window.__writes`)).toBe(2);
});

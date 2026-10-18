import { expect, test } from '@playwright/test';
import { gotoApp } from './helpers';

// Sous macOS, un clic droit sur la grille coloriait tout le fond en bleu : le navigateur sélectionnait le texte de la page.
test('clic droit sur une cellule : la cellule est sélectionnée, sans sélection de texte de la page', async ({ page }) => {
  await gotoApp(page);
  const box = (await page.locator('canvas[id^="univer-sheet-main-canvas"]').boundingBox())!;
  await page.mouse.click(box.x + 150, box.y + 80, { button: 'right' });
  await expect.poll(() => page.evaluate(`window.grillon.api.getActiveWorkbook().getActiveRange().getA1Notation()`)).toBe('B3');
  const collapsed = await page.evaluate(() => getSelection()?.isCollapsed ?? true);
  expect(collapsed).toBe(true);
});

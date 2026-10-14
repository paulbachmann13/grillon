import { expect, test, type Page } from '@playwright/test';
import { gotoApp } from './helpers';

async function run<T = unknown>(page: Page, body: string): Promise<T> {
  return page.evaluate(`(async () => { const api = window.grillon.api; const wb = api.getActiveWorkbook(); const ws = wb.getActiveSheet(); ${body} })()`) as Promise<T>;
}
const active = (page: Page) => run<string>(page, `return wb.getActiveRange().getA1Notation();`);
const value = (page: Page, a1: string) => run<string>(page, `return ws.getRange('${a1}').getDisplayValue();`);

test.beforeEach(async ({ page }) => {
  await gotoApp(page);
  await run(page, `ws.getRange('A1:C4').setValues([[1,2,3],[4,5,6],[7,8,9],[10,11,12]]);`);
  await page.locator('#univer').click({ position: { x: 600, y: 500 } });
  await run(page, `wb.setActiveRange(ws.getRange('A1'));`);
});

test('navigation : Ctrl+flèches et Ctrl+Maj+flèches', async ({ page }) => {
  await page.keyboard.press('Control+ArrowDown');
  await expect.poll(() => active(page)).toBe('A4');
  await page.keyboard.press('Control+ArrowRight');
  await expect.poll(() => active(page)).toBe('C4');
  await run(page, `wb.setActiveRange(ws.getRange('A1'));`);
  await page.keyboard.press('Control+Shift+ArrowDown');
  await expect.poll(() => active(page)).toBe('A1:A4');
});

test('mise en forme et annulation : Ctrl+B, Ctrl+I, Ctrl+U, Ctrl+Z, Ctrl+Y', async ({ page }) => {
  const style = () => run<{ bl?: number; it?: number; ul?: { s: number } }>(page, `return ws.getRange('A1').getCellStyleData() || {};`);
  await page.keyboard.press('Control+B');
  await expect.poll(async () => (await style()).bl).toBe(1);
  await page.keyboard.press('Control+I');
  await expect.poll(async () => (await style()).it).toBe(1);
  await page.keyboard.press('Control+U');
  await expect.poll(async () => (await style()).ul?.s).toBe(1);
  await page.keyboard.press('Control+Z');
  await expect.poll(async () => (await style()).ul?.s ?? 0).toBe(0);
  await page.keyboard.press('Control+Y');
  await expect.poll(async () => (await style()).ul?.s).toBe(1);
});

test('saisie : F2, Suppr, Ctrl+; (date du jour), Ctrl+: (heure)', async ({ page }) => {
  await page.keyboard.press('Delete');
  await expect.poll(() => value(page, 'A1')).toBe('');
  await run(page, `wb.setActiveRange(ws.getRange('B1'));`);
  await page.keyboard.press('F2');
  await page.keyboard.type('0');
  await page.keyboard.press('Enter');
  await expect.poll(() => value(page, 'B1')).toBe('20');
  await run(page, `wb.setActiveRange(ws.getRange('E1'));`);
  await page.keyboard.press('Control+;');
  const today = new Date();
  const fr = `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${today.getFullYear()}`;
  await expect.poll(() => value(page, 'E1')).toBe(fr);
  await run(page, `wb.setActiveRange(ws.getRange('E2'));`);
  await page.keyboard.press('Control+:');
  await expect.poll(() => value(page, 'E2')).toMatch(/^\d{2}:\d{2}$/);
});

test('copier/coller : Ctrl+C, Ctrl+V', async ({ page, context }) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.keyboard.press('Control+C');
  await run(page, `wb.setActiveRange(ws.getRange('E5'));`);
  await page.keyboard.press('Control+V');
  await expect.poll(() => value(page, 'E5')).toBe('1');
});

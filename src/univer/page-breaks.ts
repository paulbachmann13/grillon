/**
 * Sauts de page : Insérer / Supprimer / Rétablir tous les sauts de page, et Aperçu des sauts de page
 * (lignes bleues pleines pour les sauts manuels, en pointillés pour les automatiques, « Page n »
 * en filigrane). Les sauts sont rangés dans la mise en page de la feuille (custom.grillon.page).
 */
import type { FUniver } from '@univerjs/presets';
import type { FWorkbook, FWorksheet } from '@univerjs/sheets/facade';
import { hasBreakAt, insertBreaks, removeBreaks, withBreaks, type BreakSelection } from '../data/page-breaks';
import { computeLayout } from '../data/page-layout';
import { parseRowRange } from '../data/page-setup';
import { t } from '../i18n';
import { rangeFromA1 } from '../io/xlsx/units';
import { getPageSetup, savePageSetup } from '../app/page-setup-store';

interface Skeleton {
  rowHeightAccumulation: number[];
  columnWidthAccumulation: number[];
  rowHeaderWidth: number;
  columnHeaderHeight: number;
}
type UiSheet = FWorksheet & {
  getSkeleton(): Skeleton | null | undefined;
  getScrollState(): { offsetX: number; offsetY: number; sheetViewStartRow: number; sheetViewStartColumn: number };
  getZoom(): number;
};

function workbook(api: FUniver): FWorkbook | null {
  return api.getActiveWorkbook() as unknown as FWorkbook | null;
}

function selection(api: FUniver): BreakSelection | null {
  const wb = workbook(api);
  const ws = wb?.getActiveSheet();
  const range = wb?.getActiveRange();
  if (!ws || !range) return null;
  return {
    startRow: range.getRow(),
    startCol: range.getColumn(),
    wholeRow: range.getLastColumn() - range.getColumn() + 1 >= ws.getMaxColumns(),
    wholeCol: range.getLastRow() - range.getRow() + 1 >= ws.getMaxRows(),
  };
}

export function insertPageBreak(api: FUniver): void {
  const sel = selection(api);
  if (!sel) return;
  const page = getPageSetup(api);
  savePageSetup(api, withBreaks(page, insertBreaks(page, sel)));
  notify();
}

export function removePageBreak(api: FUniver): void {
  const sel = selection(api);
  if (!sel) return;
  const page = getPageSetup(api);
  if (!hasBreakAt(page, sel)) return;
  savePageSetup(api, withBreaks(page, removeBreaks(page, sel)));
  notify();
}

export function resetPageBreaks(api: FUniver): void {
  const page = getPageSetup(api);
  savePageSetup(api, withBreaks(page, {}));
  notify();
}

/** Les sauts changent les métadonnées de la feuille, pas une commande Univer : on signale le classeur modifié. */
let onChange: (() => void) | undefined;
function notify(): void {
  onChange?.();
  redraw?.();
}
let redraw: (() => void) | undefined;

let previewOn = false;
export const pageBreakPreviewEnabled = () => previewOn;

export function setPageBreakPreview(on: boolean): void {
  previewOn = on;
  redraw?.();
}

export function installPageBreaks(api: FUniver, markDirty: () => void): void {
  onChange = markDirty;
  const layer = document.createElement('div');
  layer.className = 'g-pagebreaks';
  layer.dataset.testid = 'page-break-preview';
  let timer = 0;
  const schedule = () => {
    if (!timer) timer = window.setTimeout(() => ((timer = 0), draw()), 16);
  };
  redraw = schedule;

  const canvas = () => document.querySelector<HTMLCanvasElement>('canvas[id^="univer-sheet-main-canvas"]');
  const add = (cls: string, style: Partial<CSSStyleDeclaration>, text?: string, testId?: string) => {
    const d = document.createElement('div');
    d.className = cls;
    Object.assign(d.style, style);
    if (text) d.textContent = text;
    if (testId) d.dataset.testid = testId;
    layer.append(d);
    return d;
  };

  function draw(): void {
    const ws = workbook(api)?.getActiveSheet() as unknown as UiSheet | undefined;
    const cv = canvas();
    if (!previewOn || !ws || !cv?.parentElement) return void layer.remove();
    const sk = ws.getSkeleton();
    if (!sk) return;
    const page = getPageSetup(api);
    // Zone : zone d'impression, sinon cellules utilisées.
    let [r1, c1, r2, c2] = [0, 0, Math.max(0, ws.getLastRow()), Math.max(0, ws.getLastColumn())];
    if (page.printArea) {
      try {
        const a = rangeFromA1(page.printArea.replace(/\$/g, '').replace(/^.*!/, ''));
        [r1, c1, r2, c2] = [a.startRow, a.startColumn, a.endRow, a.endColumn];
      } catch {
        /* zone invalide : cellules utilisées */
      }
    }
    const before = (acc: number[], i: number) => (i <= 0 ? 0 : acc[Math.min(i, acc.length) - 1] ?? 0);
    const rowSize = (r: number) => before(sk.rowHeightAccumulation, r + 1) - before(sk.rowHeightAccumulation, r);
    const colSize = (c: number) => before(sk.columnWidthAccumulation, c + 1) - before(sk.columnWidthAccumulation, c);
    const titles = parseRowRange(page.titleRows);
    let repeatedHeight = 0;
    if (titles) for (let r = titles.first; r <= titles.last; r++) repeatedHeight += rowSize(r);
    const layout = computeLayout({
      page,
      rows: { from: r1, to: r2 },
      cols: { from: c1, to: c2 },
      rowSize: (r) => (titles && r >= titles.first && r <= titles.last ? 0 : rowSize(r)),
      colSize,
      repeatedHeight: repeatedHeight + (page.headings ? 20 : 0),
      repeatedWidth: page.headings ? 34 : 0,
    });

    layer.replaceChildren();
    const parent = cv.parentElement;
    if (getComputedStyle(parent).position === 'static') parent.style.position = 'relative';
    if (layer.parentElement !== parent) parent.append(layer);
    const box = cv.getBoundingClientRect();
    const pbox = parent.getBoundingClientRect();
    Object.assign(layer.style, { left: `${box.left - pbox.left}px`, top: `${box.top - pbox.top}px`, width: `${box.width}px`, height: `${box.height}px` });

    const zoom = ws.getZoom() || 1;
    const scroll = ws.getScrollState();
    const headW = sk.rowHeaderWidth * zoom;
    const headH = sk.columnHeaderHeight * zoom;
    // Position (pixels du canevas) du bord de la ligne / colonne i ; hors de la vue, la valeur sort du cadre.
    const x = (c: number) => headW + (before(sk.columnWidthAccumulation, c) - before(sk.columnWidthAccumulation, scroll.sheetViewStartColumn) - scroll.offsetX) * zoom;
    const y = (r: number) => headH + (before(sk.rowHeightAccumulation, r) - before(sk.rowHeightAccumulation, scroll.sheetViewStartRow) - scroll.offsetY) * zoom;

    const left = x(c1);
    const right = x(c2 + 1);
    const top = y(r1);
    const bottom = y(r2 + 1);
    // Cadre de la zone imprimée.
    add('g-pb-area', { left: `${left}px`, top: `${top}px`, width: `${right - left}px`, height: `${bottom - top}px` });
    for (const p of layout.rowPages.slice(1)) add(`g-pb-line g-pb-h ${p.manual ? 'manual' : 'auto'}`, { left: `${left}px`, width: `${right - left}px`, top: `${y(p.start)}px` }, undefined, `pb-row-${p.manual ? 'manual' : 'auto'}-${p.start}`);
    for (const p of layout.colPages.slice(1)) add(`g-pb-line g-pb-v ${p.manual ? 'manual' : 'auto'}`, { top: `${top}px`, height: `${bottom - top}px`, left: `${x(p.start)}px` }, undefined, `pb-col-${p.manual ? 'manual' : 'auto'}-${p.start}`);
    for (const pg of layout.pages) {
      const [l, r, tp, b] = [x(pg.cols.start), x(pg.cols.end + 1), y(pg.rows.start), y(pg.rows.end + 1)];
      // Étiquette au centre de la partie visible de la page.
      const [cl, cr, ct, cb] = [Math.max(l, headW), Math.min(r, box.width), Math.max(tp, headH), Math.min(b, box.height)];
      if (cr <= cl || cb <= ct) continue;
      add('g-pb-label', { left: `${cl}px`, top: `${ct}px`, width: `${cr - cl}px`, height: `${cb - ct}px` }, t('pageBreaks.pageN', { n: pg.n }), `pb-page-${pg.n}`);
    }
    const bar = add('g-pb-bar', {}, `${t('pageBreaks.preview')} — ${t('pageBreaks.hint')} `);
    const normal = document.createElement('button');
    normal.type = 'button';
    normal.textContent = t('pageBreaks.normal');
    normal.dataset.testid = 'pb-normal';
    normal.addEventListener('click', () => setPageBreakPreview(false));
    bar.append(normal);
  }

  api.addEvent(api.Event.Scroll, schedule);
  api.addEvent(api.Event.ActiveSheetChanged, schedule);
  api.addEvent(api.Event.CommandExecuted, schedule);
  window.addEventListener('resize', schedule);
}

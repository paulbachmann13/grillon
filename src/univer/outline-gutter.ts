/**
 * Marge du plan, comme Excel : à gauche des numéros de ligne (et au-dessus des lettres de
 * colonne), les boutons de niveau 1, 2, 3… et, pour chaque groupe, un crochet avec un bouton
 * − (réduire) ou + (développer) sur la ligne de synthèse, juste après le groupe.
 * Les niveaux viennent des métadonnées (custom.grillon.outline) posées par Données → Plan.
 */
import type { FUniver } from '@univerjs/presets';
import type { FWorksheet } from '@univerjs/sheets/facade';
import { hiddenForLevel, outlineBlocks, rowsToExpand } from '../data/outline';
import { t } from '../i18n';
import { GRILLON_CUSTOM_KEY } from '../model/univer-types';

const STEP = 14; // largeur d'un niveau dans la marge, en pixels
const HEADER = 46; // largeur par défaut des numéros de ligne
const COL_HEADER = 20; // hauteur par défaut des lettres de colonne

type Axis = 'row' | 'col';
type Custom = { [GRILLON_CUSTOM_KEY]?: { outline?: number } } | undefined;
interface Core {
  getRowManager(): { getRowData(): Record<number, { custom?: Custom }> };
  getColumnManager(): { getColumnData(): Record<number, { custom?: Custom }> };
  getRowVisible(r: number): boolean;
  getColVisible(c: number): boolean;
}
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
  setRowHeaderWidth(w: number): unknown;
  setColumnHeaderHeight(h: number): unknown;
  getFreeze(): { startRow: number; startColumn: number; xSplit: number; ySplit: number };
  hideColumns(c: number, n?: number): unknown;
  showColumns(c: number, n?: number): unknown;
};

/** Niveaux de plan de la feuille, indexés par ligne ou colonne (tableau creux complété par 0). */
export function readLevels(ws: FWorksheet, axis: Axis): number[] {
  const core = ws.getSheet() as unknown as Core;
  const data = axis === 'row' ? core.getRowManager().getRowData() : core.getColumnManager().getColumnData();
  const levels: number[] = [];
  for (const key of Object.keys(data)) {
    const level = data[Number(key)]?.custom?.[GRILLON_CUSTOM_KEY]?.outline ?? 0;
    if (level > 0) levels[Number(key)] = level;
  }
  return Array.from(levels, (l) => l ?? 0);
}

const visible = (ws: FWorksheet, axis: Axis, i: number) => {
  const core = ws.getSheet() as unknown as Core;
  return axis === 'row' ? core.getRowVisible(i) : core.getColVisible(i);
};

/** Affiche ou masque des lignes/colonnes par séries contiguës (une commande annulable chacune). */
function setVisible(ws: UiSheet, axis: Axis, indexes: number[], show: boolean): void {
  const sorted = [...new Set(indexes)].sort((a, b) => a - b).filter((i) => visible(ws, axis, i) !== show);
  let i = 0;
  while (i < sorted.length) {
    let j = i;
    while (j + 1 < sorted.length && sorted[j + 1] === sorted[j] + 1) j++;
    const [start, count] = [sorted[i], j - i + 1];
    if (axis === 'row') (show ? ws.showRows(start, count) : ws.hideRows(start, count));
    else (show ? ws.showColumns(start, count) : ws.hideColumns(start, count));
    i = j + 1;
  }
}

/** Boutons 1, 2, 3… : n'affiche que les niveaux inférieurs à `show`. */
export function showLevel(ws: FWorksheet, axis: Axis, show: number): void {
  const levels = readLevels(ws, axis);
  const hidden = hiddenForLevel(levels, show);
  const toHide: number[] = [];
  const toShow: number[] = [];
  hidden.forEach((h, i) => (levels[i] > 0 ? (h ? toHide : toShow).push(i) : undefined));
  setVisible(ws as UiSheet, axis, toShow, true);
  setVisible(ws as UiSheet, axis, toHide, false);
}

/** Bouton − / + d'un groupe. */
export function toggleBlock(ws: FWorksheet, axis: Axis, block: { from: number; to: number; level: number }): void {
  const levels = readLevels(ws, axis);
  const collapsed = isCollapsed(ws, axis, block);
  const all = Array.from({ length: block.to - block.from + 1 }, (_, k) => block.from + k);
  if (collapsed) setVisible(ws as UiSheet, axis, rowsToExpand(levels, block), true);
  else setVisible(ws as UiSheet, axis, all, false);
}

function isCollapsed(ws: FWorksheet, axis: Axis, block: { from: number; to: number }): boolean {
  for (let i = block.from; i <= block.to; i++) if (visible(ws, axis, i)) return false;
  return true;
}

export function installOutlineGutter(api: FUniver): void {
  const layer = document.createElement('div');
  layer.className = 'g-outline';
  layer.dataset.testid = 'outline-gutter';
  let frame = 0;
  const schedule = () => {
    if (!frame) frame = window.setTimeout(() => ((frame = 0), draw()), 16);
  };

  function canvas(): HTMLCanvasElement | null {
    return document.querySelector<HTMLCanvasElement>('canvas[id^="univer-sheet-main-canvas"]');
  }

  function button(text: string, x: number, y: number, title: string, onClick: () => void, testId?: string): void {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'g-outline-btn';
    b.textContent = text;
    b.title = title;
    b.style.left = `${x}px`;
    b.style.top = `${y}px`;
    if (testId) b.dataset.testid = testId;
    b.addEventListener('mousedown', (e) => e.stopPropagation());
    b.addEventListener('click', (e) => {
      e.stopPropagation();
      onClick();
    });
    layer.append(b);
  }

  function line(x: number, y: number, w: number, h: number): void {
    const l = document.createElement('div');
    l.className = 'g-outline-line';
    Object.assign(l.style, { left: `${x}px`, top: `${y}px`, width: `${w}px`, height: `${h}px` });
    layer.append(l);
  }

  function draw(): void {
    const ws = api.getActiveWorkbook()?.getActiveSheet() as unknown as UiSheet | undefined;
    const cv = canvas();
    if (!ws || !cv || !cv.parentElement) return void layer.remove();
    const rows = readLevels(ws, 'row');
    const cols = readLevels(ws, 'col');
    const maxRow = rows.reduce((m, l) => Math.max(m, l), 0);
    const maxCol = cols.reduce((m, l) => Math.max(m, l), 0);
    const gutterW = maxRow ? (maxRow + 1) * STEP + 4 : 0;
    const gutterH = maxCol ? (maxCol + 1) * STEP + 4 : 0;
    // Agrandit les en-têtes pour faire de la place à la marge (sans boucle : seulement si ça change).
    const sk = ws.getSkeleton();
    if (!sk) return;
    if (sk.rowHeaderWidth !== HEADER + gutterW) ws.setRowHeaderWidth(HEADER + gutterW);
    if (sk.columnHeaderHeight !== COL_HEADER + gutterH) ws.setColumnHeaderHeight(COL_HEADER + gutterH);
    layer.replaceChildren();
    if (!maxRow && !maxCol) return void layer.remove();

    const parent = cv.parentElement;
    if (getComputedStyle(parent).position === 'static') parent.style.position = 'relative';
    if (layer.parentElement !== parent) parent.append(layer);
    const box = cv.getBoundingClientRect();
    const pbox = parent.getBoundingClientRect();
    Object.assign(layer.style, { left: `${box.left - pbox.left}px`, top: `${box.top - pbox.top}px`, width: `${box.width}px`, height: `${box.height}px` });

    const zoom = ws.getZoom() || 1;
    const scroll = ws.getScrollState();
    const freeze = ws.getFreeze();
    const headerW = sk.rowHeaderWidth * zoom;
    const headerH = sk.columnHeaderHeight * zoom;
    const before = (acc: number[], i: number) => (i <= 0 ? 0 : acc[Math.min(i, acc.length) - 1] ?? 0);
    // Position (en pixels du canevas) du début de la ligne/colonne i, ou null si hors de la vue.
    const pos = (axis: Axis, i: number): number | null => {
      const acc = axis === 'row' ? sk.rowHeightAccumulation : sk.columnWidthAccumulation;
      const frozen = axis === 'row' ? Math.max(0, freeze.startRow) : Math.max(0, freeze.startColumn);
      const head = axis === 'row' ? headerH : headerW;
      const limit = axis === 'row' ? box.height : box.width;
      if (i < frozen) return head + before(acc, i) * zoom;
      const start = axis === 'row' ? scroll.sheetViewStartRow : scroll.sheetViewStartColumn;
      const offset = axis === 'row' ? scroll.offsetY : scroll.offsetX;
      const p = head + (before(acc, i) - before(acc, start) - offset) * zoom;
      const frozenEnd = head + before(acc, frozen) * zoom;
      return p < frozenEnd - 1 || p > limit ? null : p;
    };
    const size = (axis: Axis, i: number) => {
      const acc = axis === 'row' ? sk.rowHeightAccumulation : sk.columnWidthAccumulation;
      return (before(acc, i + 1) - before(acc, i)) * zoom;
    };

    const drawAxis = (axis: Axis, levels: number[], max: number) => {
      if (!max) return;
      // Boutons de niveau dans le coin (comme les 1 2 3 d'Excel).
      for (let n = 1; n <= max + 1; n++) {
        const x = axis === 'row' ? 2 + (n - 1) * STEP : headerW - STEP - 2;
        const y = axis === 'row' ? headerH - STEP - 2 : 2 + (n - 1) * STEP;
        button(String(n), x, y, t('data.outlineLevel', { n: String(n) }), () => showLevel(ws, axis, n), `outline-${axis}-level-${n}`);
      }
      for (const block of outlineBlocks(levels)) {
        const offset = 2 + (block.level - 1) * STEP;
        const first = pos(axis, block.from);
        const summary = block.to + 1;
        const last = pos(axis, summary) ?? (pos(axis, block.to) !== null ? (pos(axis, block.to) as number) + size(axis, block.to) : null);
        const collapsed = isCollapsed(ws, axis, block);
        // Crochet : un trait le long du groupe, fermé au niveau de la ligne de synthèse.
        if (!collapsed && (first !== null || last !== null)) {
          const a = first ?? (axis === 'row' ? headerH : headerW);
          const b = last ?? (axis === 'row' ? box.height : box.width);
          if (axis === 'row') line(offset + STEP / 2, a + 2, 1, Math.max(0, b - a - 2));
          else line(a + 2, offset + STEP / 2, Math.max(0, b - a - 2), 1);
        }
        const at = pos(axis, summary);
        if (at === null) continue;
        const centre = at + size(axis, summary) / 2 - STEP / 2;
        const label = collapsed ? '+' : '−';
        const title = collapsed ? t('data.showDetail') : t('data.hideDetail');
        const testId = `outline-${axis}-toggle-${summary}`;
        if (axis === 'row') button(label, offset, centre, title, () => toggleBlock(ws, axis, block), testId);
        else button(label, centre, offset, title, () => toggleBlock(ws, axis, block), testId);
      }
    };
    drawAxis('row', rows, maxRow);
    drawAxis('col', cols, maxCol);
  }

  api.addEvent(api.Event.Scroll, schedule);
  api.addEvent(api.Event.ActiveSheetChanged, schedule);
  api.addEvent(api.Event.CommandExecuted, schedule);
  window.addEventListener('resize', schedule);
  setTimeout(schedule, 0);
}

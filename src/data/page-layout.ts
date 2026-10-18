/**
 * Découpage d'une feuille en pages imprimées : sert à l'impression et à l'aperçu des sauts de page,
 * pour que les deux montrent les mêmes pages.
 */
import { normalizeBreaks, pageGrid, paginate, type PageRect, type PageSpan } from './page-breaks';
import { printableSizePx, printZoom, type PageSetup } from './page-setup';

export interface LayoutInput {
  page: PageSetup;
  /** Zone imprimée (indices à partir de 0, bornes comprises). */
  rows: { from: number; to: number };
  cols: { from: number; to: number };
  /** Hauteur / largeur en pixels (0 pour une ligne ou une colonne masquée). */
  rowSize: (r: number) => number;
  colSize: (c: number) => number;
  /** Hauteur des lignes répétées en haut de chaque page (titres, en-têtes de lignes et de colonnes). */
  repeatedHeight?: number;
  /** Largeur de la colonne des numéros de ligne, répétée sur chaque page quand les en-têtes sont imprimés. */
  repeatedWidth?: number;
}

export interface PageLayout {
  zoom: number;
  rowPages: PageSpan[];
  colPages: PageSpan[];
  pages: PageRect[];
  /** Les sauts manuels sont ignorés quand on imprime « ajusté à N pages de large », comme dans Excel. */
  manualIgnored: boolean;
}

export function computeLayout(input: LayoutInput): PageLayout {
  const { page, rows, cols } = input;
  let contentWidth = input.repeatedWidth ?? 0;
  for (let c = cols.from; c <= cols.to; c++) contentWidth += Math.max(0, input.colSize(c));
  const zoom = printZoom(page, contentWidth);
  const size = printableSizePx(page);
  const manualIgnored = !!page.fitWidth;
  const rowCapacity = size.height / zoom - (input.repeatedHeight ?? 0);
  const colCapacity = size.width / zoom - (input.repeatedWidth ?? 0);
  const rowPages = paginate(input.rowSize, rows.from, rows.to, rowCapacity, manualIgnored ? [] : normalizeBreaks(page.rowBreaks));
  const colPages = paginate(input.colSize, cols.from, cols.to, colCapacity, manualIgnored ? [] : normalizeBreaks(page.colBreaks));
  return { zoom, rowPages, colPages, pages: pageGrid(rowPages, colPages), manualIgnored };
}

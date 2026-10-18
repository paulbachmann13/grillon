/**
 * Impression et export PDF : on construit une page HTML de la feuille (valeurs affichées,
 * mise en forme principale, fusions, graphiques et images) et on ouvre la boîte d'impression
 * du navigateur, qui propose « Enregistrer au format PDF ».
 */
import type { FUniver } from '@univerjs/presets';
import type { FWorkbook, FWorksheet } from '@univerjs/sheets/facade';
import { t } from '../i18n';
import type { IBorderStyleData, IStyleData, IWorkbookData } from '../model/univer-types';
import { hasBreaks } from '../data/page-breaks';
import { computeLayout } from '../data/page-layout';
import { hfToCssContent, parseRowRange, PAPER_SIZES, DEFAULT_PAGE, type HeaderFooterSpec, type PageSetup } from '../data/page-setup';
import { columnToLetters, rangeFromA1 } from '../io/xlsx/units';
import { actionButton, box, checkbox, field, openDialog, select } from '../ui/dialog';
import { getPageSetup } from './page-setup-store';

export interface PrintOptions {
  selectionOnly: boolean;
  /** Mise en page effective (réglages de la feuille, éventuellement modifiés dans la boîte d'impression). */
  page: PageSetup;
  fileName: string;
  /** Une seule page (à partir de 0) : sert à l'aperçu, qui montre les vraies pages. */
  pageIndex?: number;
}

/** Hauteur de la ligne des lettres de colonne quand les en-têtes sont imprimés. */
const HEADING_ROW_PX = 20;

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const color = (c?: { rgb?: string | null | void } | null | void) => (c && c.rgb ? c.rgb : undefined);

function borderCss(b?: IBorderStyleData | null | void): string | undefined {
  if (!b || !b.s) return undefined;
  const width = [8, 9, 10, 11].includes(b.s) ? 2 : b.s === 13 ? 3 : 1;
  const style = b.s === 7 ? 'double' : [3, 2].includes(b.s) ? 'dotted' : [4, 5, 6, 9, 10, 11, 12].includes(b.s) ? 'dashed' : 'solid';
  return `${b.s === 7 ? 3 : width}px ${style} ${color(b.cl) ?? '#000'}`;
}

function styleCss(s: IStyleData | null | undefined): string {
  if (!s) return '';
  const css: string[] = [];
  if (s.bl) css.push('font-weight:700');
  if (s.it) css.push('font-style:italic');
  const deco = [s.ul?.s ? 'underline' : '', s.st?.s ? 'line-through' : ''].filter(Boolean).join(' ');
  if (deco) css.push(`text-decoration:${deco}`);
  if (s.ff) css.push(`font-family:'${s.ff.replace(/'/g, '')}',sans-serif`);
  if (s.fs) css.push(`font-size:${s.fs}pt`);
  if (color(s.cl)) css.push(`color:${color(s.cl)}`);
  if (color(s.bg)) css.push(`background:${color(s.bg)}`);
  if (s.ht === 2) css.push('text-align:center');
  else if (s.ht === 3) css.push('text-align:right');
  else if (s.ht === 1) css.push('text-align:left');
  if (s.vt === 1) css.push('vertical-align:top');
  else if (s.vt === 2) css.push('vertical-align:middle');
  if (s.tb === 3) css.push('white-space:pre-wrap');
  for (const [side, key] of [['top', 't'], ['right', 'r'], ['bottom', 'b'], ['left', 'l']] as const) {
    const b = borderCss(s.bd?.[key]);
    if (b) css.push(`border-${side}:${b}`);
  }
  return css.join(';');
}

/** Construit le document HTML imprimable d'une feuille. */
export function buildPrintHtml(api: FUniver, opts: PrintOptions): string | null {
  return buildPrint(api, opts)?.html ?? null;
}

/** Document HTML imprimable et nombre de pages (sauts manuels et automatiques compris). */
export function buildPrint(api: FUniver, opts: PrintOptions): { html: string; pageCount: number } | null {
  const wb = api.getActiveWorkbook() as unknown as FWorkbook | null;
  const ws = wb?.getActiveSheet() as unknown as FWorksheet | undefined;
  if (!wb || !ws) return null;
  const data = wb.save() as unknown as IWorkbookData;
  const sheet = data.sheets[ws.getSheetId()];
  let r1 = 0;
  let c1 = 0;
  let r2 = ws.getLastRow();
  let c2 = ws.getLastColumn();
  const sel = wb.getActiveRange();
  if (opts.selectionOnly && sel) {
    r1 = sel.getRow();
    c1 = sel.getColumn();
    r2 = sel.getLastRow();
    c2 = sel.getLastColumn();
  } else if (opts.page.printArea) {
    try {
      const area = rangeFromA1(opts.page.printArea.replace(/\$/g, '').replace(/^.*!/, ''));
      r1 = area.startRow;
      c1 = area.startColumn;
      r2 = area.endRow;
      c2 = area.endColumn;
    } catch {
      /* zone d'impression invalide : on imprime tout */
    }
  }
  if (r2 < 0 || c2 < 0) {
    r2 = 0;
    c2 = 0;
  }
  const rows = r2 - r1 + 1;
  const cols = c2 - c1 + 1;
  const display = ws.getRange(r1, c1, rows, cols).getDisplayValues();
  const styles = data.styles ?? {};
  const resolve = (s: unknown): IStyleData | null => (typeof s === 'string' ? (styles[s] ?? null) : ((s as IStyleData) ?? null));
  // Fusions : cellule maître (colspan/rowspan) et cellules couvertes.
  const covered = new Map<string, { row: number; col: number }>();
  const spans = new Map<string, { rs: number; cs: number }>();
  for (const m of sheet?.mergeData ?? []) {
    spans.set(`${m.startRow},${m.startColumn}`, { rs: m.endRow - m.startRow + 1, cs: m.endColumn - m.startColumn + 1 });
    for (let r = m.startRow; r <= m.endRow; r++) for (let c = m.startColumn; c <= m.endColumn; c++) if (r !== m.startRow || c !== m.startColumn) covered.set(`${r},${c}`, { row: m.startRow, col: m.startColumn });
  }
  const widths = Array.from({ length: cols }, (_, i) => (sheet?.columnData?.[c1 + i]?.hd ? 0 : ws.getColumnWidth(c1 + i)));
  const page = opts.page;
  const headings = !!page.headings;
  const headingW = headings ? 34 : 0;
  const rowHidden = (r: number) => !!sheet?.rowData?.[r]?.hd;
  const rowPx = (r: number) => (rowHidden(r) ? 0 : ws.getRowHeight(r));
  const titles = parseRowRange(page.titleRows);
  const isTitle = (r: number) => !!titles && r >= titles.first && r <= titles.last;
  // Lignes de titre répétées en haut de chaque page (celles hors de la zone imprimée comprises).
  const titleRows: number[] = [];
  if (titles) for (let r = titles.first; r <= titles.last; r++) if (!rowHidden(r)) titleRows.push(r);
  const repeatedHeight = titleRows.reduce((h, r) => h + ws.getRowHeight(r), 0) + (headings ? HEADING_ROW_PX : 0);
  // Pages : les mêmes que dans l'aperçu des sauts de page (sauts manuels et automatiques).
  const layout = computeLayout({
    page,
    rows: { from: r1, to: r2 },
    cols: { from: c1, to: c2 },
    rowSize: (r) => (isTitle(r) ? 0 : rowPx(r)),
    colSize: (c) => widths[c - c1],
    repeatedHeight,
    repeatedWidth: headingW,
  });
  const zoom = layout.zoom;
  const only = opts.pageIndex === undefined ? undefined : layout.pages[Math.min(Math.max(0, opts.pageIndex), layout.pages.length - 1)];
  const rowStarts = new Set(only ? [] : layout.rowPages.slice(1).map((p) => p.start));
  const rowHtml = (r: number, band: { start: number; end: number }): string => {
    const i = r - r1;
    if (rowHidden(r)) return '';
    const cells: string[] = headings ? [`<th class="rh">${r + 1}</th>`] : [];
    for (let c = band.start; c <= band.end; c++) {
      const j = c - c1;
      const master = covered.get(`${r},${c}`);
      // Une fusion dont la cellule principale est sur une autre page se lit comme une cellule vide.
      if ((master && master.col >= band.start) || !widths[j]) continue;
      const cell = sheet?.cellData?.[r]?.[c];
      const span = spans.get(`${r},${c}`);
      const colspan = span ? Math.min(span.cs, band.end - c + 1) : 0;
      const value = display[i]?.[j] ?? '';
      const numeric = typeof cell?.v === 'number' || cell?.t === 2;
      const css = `${numeric ? 'text-align:right;' : ''}${styleCss(resolve(cell?.s))}`;
      cells.push(`<td${span ? ` rowspan="${span.rs}" colspan="${colspan}"` : ''}${css ? ` style="${css}"` : ''}>${esc(value)}</td>`);
    }
    const breakBefore = rowStarts.has(r) ? 'break-before:page;' : '';
    return `<tr style="${breakBefore}height:${ws.getRowHeight(r)}px">${cells.join('')}</tr>`;
  };
  // Une table par bande de colonnes (Excel imprime de haut en bas, puis de gauche à droite).
  const tables = (only ? [only.cols] : layout.colPages).map((band, k) => {
    const bandWidths = widths.slice(band.start - c1, band.end - c1 + 1);
    const tableWidth = bandWidths.reduce((a, b) => a + b, 0) + headingW;
    const headRows: string[] = [];
    if (headings) {
      const letters = bandWidths.map((w, j) => (w ? `<th class="ch">${columnToLetters(band.start + j)}</th>` : '')).join('');
      headRows.push(`<tr><th class="rh"></th>${letters}</tr>`);
    }
    for (const r of titleRows) headRows.push(rowHtml(r, band));
    const body: string[] = [];
    for (let r = only ? only.rows.start : r1; r <= (only ? only.rows.end : r2); r++) if (!isTitle(r)) body.push(rowHtml(r, band));
    const colsHtml = (headings ? ['<col style="width:34px">'] : []).concat(bandWidths.filter(Boolean).map((w) => `<col style="width:${w}px">`)).join('');
    return `<table${k ? ' class="band"' : ''} style="width:${tableWidth}px"><colgroup>${colsHtml}</colgroup>${headRows.length ? `<thead>${headRows.join('')}</thead>` : ''}<tbody>${body.join('')}</tbody></table>`;
  });
  // Graphiques et images : capturés depuis l'écran.
  const extras: string[] = [];
  if (!opts.selectionOnly && (!only || only.n === layout.pages.length)) {
    document.querySelectorAll<HTMLCanvasElement>('.g-chart-canvas canvas').forEach((cv) => {
      try {
        extras.push(`<img class="chart" src="${cv.toDataURL('image/png')}" alt="">`);
      } catch {
        /* canvas indisponible */
      }
    });
    const drawings = data.resources?.find((r) => r.name === 'SHEET_DRAWING_PLUGIN');
    if (drawings) {
      const items = (JSON.parse(drawings.data)[ws.getSheetId()]?.data ?? {}) as Record<string, { drawingType: number; source?: string; transform?: { width: number; height: number } }>;
      for (const d of Object.values(items)) {
        if (d.drawingType === 0 && d.source) extras.push(`<img src="${d.source}" style="width:${d.transform?.width ?? 200}px" alt="">`);
      }
    }
  }
  const paper = PAPER_SIZES[page.paper ?? DEFAULT_PAGE.paper];
  const m = page.margins ?? DEFAULT_PAGE.margins;
  const now = new Date();
  const ctx = { file: opts.fileName, sheet: ws.getSheetName(), date: now.toLocaleDateString('fr-FR'), time: now.toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' }) };
  const boxes = (prefix: 'top' | 'bottom', hf: HeaderFooterSpec | undefined): string =>
    (['left', 'center', 'right'] as const)
      .filter((k) => hf?.[k])
      .map((k) => `@${prefix}-${k} { content: ${hfToCssContent(hf![k]!, ctx)}; font: 9pt Calibri, Carlito, Arial, sans-serif; color: #000; }`)
      .join(' ');
  const html = `<!doctype html><html lang="fr"><head><meta charset="utf-8"><title>${esc(ws.getSheetName())}</title><style>
@page { size: ${paper.widthMm}mm ${paper.heightMm}mm${page.orientation === 'landscape' ? ' landscape' : ''}; margin: ${m.top}in ${m.right}in ${m.bottom}in ${m.left}in; ${boxes('top', page.header)} ${boxes('bottom', page.footer)} }
body { margin: 0; font: 11pt Calibri, Carlito, Arial, sans-serif; color: #000; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
.sheet { zoom: ${zoom.toFixed(4)}; }
table { border-collapse: collapse; table-layout: fixed; ${page.centerH ? 'margin: 0 auto;' : ''} }
table.band { break-before: page; }
thead { display: table-header-group; }
td { padding: 1px 3px; overflow: hidden; white-space: nowrap; vertical-align: bottom; ${page.gridlines ? 'border: 1px solid #c8c8c8;' : ''} }
th.rh, th.ch { background: #f0f0f0; border: 1px solid #c8c8c8; font-weight: 400; font-size: 9pt; text-align: center; }
tr { page-break-inside: avoid; }
img { display: block; margin-top: 12px; max-width: 100%; page-break-inside: avoid; }
</style></head><body><div class="sheet">${tables.join('')}${extras.join('')}</div></body></html>`;
  return { html, pageCount: layout.pages.length };
}

function printHtml(html: string): void {
  const frame = document.createElement('iframe');
  frame.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0';
  frame.setAttribute('aria-hidden', 'true');
  document.body.append(frame);
  const doc = frame.contentDocument!;
  doc.open();
  doc.write(html);
  doc.close();
  const go = () => {
    frame.contentWindow?.focus();
    frame.contentWindow?.print();
    setTimeout(() => frame.remove(), 60_000);
  };
  // Laisse aux images (graphiques) le temps de se charger.
  if (doc.readyState === 'complete') setTimeout(go, 150);
  else frame.addEventListener('load', () => setTimeout(go, 150));
}

const MM_TO_PX = 96 / 25.4;

/**
 * Aperçu avant impression (comme le panneau de droite d'Excel) : une page aux proportions du
 * papier, marges comprises, et la navigation entre les pages. Chaque page est celle du découpage
 * réel (sauts manuels et automatiques) : on affiche le contenu de la page choisie.
 */
function printPreview(): { el: HTMLElement; show(build: (pageIndex: number) => { html: string; pageCount: number } | null, page: PageSetup): void } {
  const el = document.createElement('div');
  el.className = 'g-print-preview';
  el.dataset.testid = 'print-preview';
  const sheet = document.createElement('div');
  sheet.className = 'g-print-page';
  const frame = document.createElement('iframe');
  frame.title = t('print.preview');
  frame.setAttribute('scrolling', 'no');
  sheet.append(frame);
  // Le cadre réserve la place de la page réduite (transform ne change pas la mise en page).
  const holder = document.createElement('div');
  holder.className = 'g-print-holder';
  holder.append(sheet);
  const nav = document.createElement('div');
  nav.className = 'g-print-nav';
  const prev = actionButton('◀', () => go(index - 1));
  const next = actionButton('▶', () => go(index + 1));
  prev.setAttribute('aria-label', t('print.previousPage'));
  next.setAttribute('aria-label', t('print.nextPage'));
  const label = document.createElement('span');
  label.dataset.testid = 'print-page-count';
  nav.append(prev, label, next);
  el.append(holder, nav);
  let index = 0;
  let pages = 1;
  let builder: ((pageIndex: number) => { html: string; pageCount: number } | null) | null = null;

  function go(i: number): void {
    const built = builder?.(Math.max(0, i));
    if (!built) return;
    pages = Math.max(1, built.pageCount);
    index = Math.max(0, Math.min(pages - 1, i));
    // Si le nombre de pages a baissé, on redessine la dernière page.
    const html = index === Math.max(0, i) ? built.html : (builder?.(index)?.html ?? built.html);
    frame.srcdoc = html;
    label.textContent = t('print.pageOf', { n: index + 1, total: pages });
    prev.disabled = index === 0;
    next.disabled = index >= pages - 1;
  }

  return {
    el,
    show(build, page: PageSetup) {
      const paper = PAPER_SIZES[page.paper ?? DEFAULT_PAGE.paper];
      const landscape = page.orientation === 'landscape';
      const w = (landscape ? paper.heightMm : paper.widthMm) * MM_TO_PX;
      const h = (landscape ? paper.widthMm : paper.heightMm) * MM_TO_PX;
      const m = page.margins ?? DEFAULT_PAGE.margins;
      const [top, right, bottom, left] = [m.top, m.right, m.bottom, m.left].map((x) => x * 96);
      // La page est réduite pour tenir dans l'aperçu ; le contenu garde ses dimensions d'impression.
      const scale = Math.min(1, 380 / w, 520 / h);
      Object.assign(sheet.style, { width: `${w}px`, height: `${h}px`, padding: `${top}px ${right}px ${bottom}px ${left}px`, transform: `scale(${scale})` });
      Object.assign(holder.style, { width: `${w * scale}px`, height: `${h * scale}px` });
      Object.assign(frame.style, { width: `${w - left - right}px`, height: `${h - top - bottom}px` });
      builder = build;
      go(index);
    },
  };
}

export async function printDialog(api: FUniver, fileName = ''): Promise<void> {
  const base = getPageSetup(api);
  const orientation = select(
    [
      ['portrait', t('print.portrait')],
      ['landscape', t('print.landscape')],
    ],
    base.orientation ?? 'portrait',
  );
  const what = select(
    [
      ['sheet', t('print.activeSheet')],
      ['selection', t('print.selection')],
    ],
    'sheet',
  );
  const gridlines = checkbox(t('print.gridlines'), !!base.gridlines);
  const fit = checkbox(t('print.fitWidth'), base.fitWidth !== undefined ? !!base.fitWidth : !base.scale);
  const preview = printPreview();

  // Les réglages de la feuille (papier, marges, zone, titres, en-têtes…) restent ; ceux de la boîte priment.
  const currentPage = (): PageSetup => {
    const current = getPageSetup(api);
    const page: PageSetup = { ...current, orientation: orientation.value === 'landscape' ? 'landscape' : 'portrait', gridlines: gridlines.input.checked };
    if (fit.input.checked) page.fitWidth = current.fitWidth || 1;
    else {
      delete page.fitWidth;
      delete page.fitHeight;
    }
    return page;
  };
  const build = () => buildPrintHtml(api, { selectionOnly: what.value === 'selection', page: currentPage(), fileName });
  const refresh = () => preview.show((pageIndex) => buildPrint(api, { selectionOnly: what.value === 'selection', page: currentPage(), fileName, pageIndex }), currentPage());
  for (const c of [orientation, what, gridlines.input, fit.input]) c.addEventListener('change', refresh);
  const setup = actionButton(t('print.pageSetup'), () => void import('./page-setup').then((m) => m.pageSetupDialog(api)).then(refresh));
  const fitNote = Object.assign(document.createElement('p'), { className: 'g-eval-status', textContent: t('pageBreaks.ignoredWithFit') });
  fitNote.dataset.testid = 'print-fit-note';
  const updateNote = () => (fitNote.hidden = !(fit.input.checked && hasBreaks(getPageSetup(api))));
  fit.input.addEventListener('change', updateNote);
  updateNote();
  const settings = box(t('print.help'), field(t('print.what'), what), field(t('print.orientation'), orientation), gridlines.el, fit.el, fitNote, setup);
  settings.classList.add('g-print-settings');
  const layout = document.createElement('div');
  layout.className = 'g-print-layout';
  layout.append(settings, preview.el);
  refresh();
  await openDialog({
    title: t('print.title'),
    testId: 'dialog-print',
    okLabel: t('print.ok'),
    xl: true,
    body: layout,
    onOk: () => {
      const html = build();
      if (html) printHtml(html);
    },
  });
}

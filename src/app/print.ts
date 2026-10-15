/**
 * Impression et export PDF : on construit une page HTML de la feuille (valeurs affichées,
 * mise en forme principale, fusions, graphiques et images) et on ouvre la boîte d'impression
 * du navigateur, qui propose « Enregistrer au format PDF ».
 */
import type { FUniver } from '@univerjs/presets';
import type { FWorkbook, FWorksheet } from '@univerjs/sheets/facade';
import { t } from '../i18n';
import type { IBorderStyleData, IStyleData, IWorkbookData } from '../model/univer-types';
import { hfToCssContent, parseRowRange, PAPER_SIZES, printZoom, DEFAULT_PAGE, type HeaderFooterSpec, type PageSetup } from '../data/page-setup';
import { columnToLetters, rangeFromA1 } from '../io/xlsx/units';
import { actionButton, box, checkbox, field, openDialog, select } from '../ui/dialog';
import { getPageSetup } from './page-setup-store';

export interface PrintOptions {
  selectionOnly: boolean;
  /** Mise en page effective (réglages de la feuille, éventuellement modifiés dans la boîte d'impression). */
  page: PageSetup;
  fileName: string;
}

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
  const covered = new Set<string>();
  const spans = new Map<string, { rs: number; cs: number }>();
  for (const m of sheet?.mergeData ?? []) {
    spans.set(`${m.startRow},${m.startColumn}`, { rs: m.endRow - m.startRow + 1, cs: m.endColumn - m.startColumn + 1 });
    for (let r = m.startRow; r <= m.endRow; r++) for (let c = m.startColumn; c <= m.endColumn; c++) if (r !== m.startRow || c !== m.startColumn) covered.add(`${r},${c}`);
  }
  const widths = Array.from({ length: cols }, (_, i) => (sheet?.columnData?.[c1 + i]?.hd ? 0 : ws.getColumnWidth(c1 + i)));
  const tableWidth = widths.reduce((a, b) => a + b, 0) + (opts.page.headings ? 34 : 0);
  const page = opts.page;
  const headings = !!page.headings;
  const rowHtml = (r: number): string => {
    const i = r - r1;
    if (sheet?.rowData?.[r]?.hd) return '';
    const cells: string[] = headings ? [`<th class="rh">${r + 1}</th>`] : [];
    for (let j = 0; j < cols; j++) {
      const c = c1 + j;
      if (covered.has(`${r},${c}`) || !widths[j]) continue;
      const cell = sheet?.cellData?.[r]?.[c];
      const span = spans.get(`${r},${c}`);
      const value = display[i]?.[j] ?? '';
      const numeric = typeof cell?.v === 'number' || cell?.t === 2;
      const css = `${numeric ? 'text-align:right;' : ''}${styleCss(resolve(cell?.s))}`;
      cells.push(`<td${span ? ` rowspan="${span.rs}" colspan="${span.cs}"` : ''}${css ? ` style="${css}"` : ''}>${esc(value)}</td>`);
    }
    return `<tr style="height:${ws.getRowHeight(r)}px">${cells.join('')}</tr>`;
  };
  const titles = parseRowRange(page.titleRows);
  const headRows: string[] = [];
  const body: string[] = [];
  for (let r = r1; r <= r2; r++) {
    if (titles && r >= titles.first && r <= titles.last) headRows.push(rowHtml(r));
    else body.push(rowHtml(r));
  }
  // Lignes de titre hors de la zone imprimée : répétées en haut de chaque page.
  if (titles) for (let r = titles.first; r <= titles.last && r < r1; r++) headRows.unshift(rowHtml(r));
  if (headings) {
    const letters = Array.from({ length: cols }, (_, j) => (widths[j] ? `<th class="ch">${columnToLetters(c1 + j)}</th>` : '')).join('');
    headRows.unshift(`<tr><th class="rh"></th>${letters}</tr>`);
  }
  // Graphiques et images : capturés depuis l'écran.
  const extras: string[] = [];
  if (!opts.selectionOnly) {
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
  const zoom = printZoom(page, tableWidth);
  const colsHtml = (headings ? ['<col style="width:34px">'] : []).concat(widths.filter(Boolean).map((w) => `<col style="width:${w}px">`)).join('');
  const paper = PAPER_SIZES[page.paper ?? DEFAULT_PAGE.paper];
  const m = page.margins ?? DEFAULT_PAGE.margins;
  const now = new Date();
  const ctx = { file: opts.fileName, sheet: ws.getSheetName(), date: now.toLocaleDateString('fr-FR'), time: now.toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' }) };
  const boxes = (prefix: 'top' | 'bottom', hf: HeaderFooterSpec | undefined): string =>
    (['left', 'center', 'right'] as const)
      .filter((k) => hf?.[k])
      .map((k) => `@${prefix}-${k} { content: ${hfToCssContent(hf![k]!, ctx)}; font: 9pt Calibri, Carlito, Arial, sans-serif; color: #000; }`)
      .join(' ');
  return `<!doctype html><html lang="fr"><head><meta charset="utf-8"><title>${esc(ws.getSheetName())}</title><style>
@page { size: ${paper.widthMm}mm ${paper.heightMm}mm${page.orientation === 'landscape' ? ' landscape' : ''}; margin: ${m.top}in ${m.right}in ${m.bottom}in ${m.left}in; ${boxes('top', page.header)} ${boxes('bottom', page.footer)} }
body { margin: 0; font: 11pt Calibri, Carlito, Arial, sans-serif; color: #000; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
.sheet { zoom: ${zoom.toFixed(4)}; }
table { border-collapse: collapse; table-layout: fixed; width: ${tableWidth}px; ${page.centerH ? 'margin: 0 auto;' : ''} }
thead { display: table-header-group; }
td { padding: 1px 3px; overflow: hidden; white-space: nowrap; vertical-align: bottom; ${page.gridlines ? 'border: 1px solid #c8c8c8;' : ''} }
th.rh, th.ch { background: #f0f0f0; border: 1px solid #c8c8c8; font-weight: 400; font-size: 9pt; text-align: center; }
tr { page-break-inside: avoid; }
img { display: block; margin-top: 12px; max-width: 100%; page-break-inside: avoid; }
</style></head><body><div class="sheet"><table><colgroup>${colsHtml}</colgroup>${headRows.length ? `<thead>${headRows.join('')}</thead>` : ''}<tbody>${body.join('')}</tbody></table>${extras.join('')}</div></body></html>`;
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
  const setup = actionButton(t('print.pageSetup'), () => void import('./page-setup').then((m) => m.pageSetupDialog(api)));
  await openDialog({
    title: t('print.title'),
    testId: 'dialog-print',
    okLabel: t('print.ok'),
    body: box(t('print.help'), field(t('print.what'), what), field(t('print.orientation'), orientation), gridlines.el, fit.el, setup),
    onOk: () => {
      // Les réglages de la feuille (papier, marges, zone, titres, en-têtes…) restent ; ceux de la boîte priment.
      const current = getPageSetup(api);
      const page: PageSetup = {
        ...current,
        orientation: orientation.value === 'landscape' ? 'landscape' : 'portrait',
        gridlines: gridlines.input.checked,
      };
      if (fit.input.checked) page.fitWidth = current.fitWidth || 1;
      else {
        delete page.fitWidth;
        delete page.fitHeight;
      }
      const html = buildPrintHtml(api, { selectionOnly: what.value === 'selection', page, fileName });
      if (html) printHtml(html);
    },
  });
}

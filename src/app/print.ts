/**
 * Impression et export PDF : on construit une page HTML de la feuille (valeurs affichées,
 * mise en forme principale, fusions, graphiques et images) et on ouvre la boîte d'impression
 * du navigateur, qui propose « Enregistrer au format PDF ».
 */
import type { FUniver } from '@univerjs/presets';
import type { FWorkbook, FWorksheet } from '@univerjs/sheets/facade';
import { t } from '../i18n';
import type { IBorderStyleData, IStyleData, IWorkbookData } from '../model/univer-types';
import { box, checkbox, field, openDialog, select } from '../ui/dialog';

export interface PrintOptions {
  landscape: boolean;
  gridlines: boolean;
  fitWidth: boolean;
  selectionOnly: boolean;
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
  const tableWidth = widths.reduce((a, b) => a + b, 0);
  const body: string[] = [];
  for (let i = 0; i < rows; i++) {
    const r = r1 + i;
    if (sheet?.rowData?.[r]?.hd) continue;
    const cells: string[] = [];
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
    body.push(`<tr style="height:${ws.getRowHeight(r)}px">${cells.join('')}</tr>`);
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
  const pageWidthPx = (opts.landscape ? 297 - 24 : 210 - 24) * 3.78;
  const zoom = opts.fitWidth && tableWidth > pageWidthPx ? pageWidthPx / tableWidth : 1;
  const cols_ = widths.filter(Boolean).map((w) => `<col style="width:${w}px">`).join('');
  return `<!doctype html><html lang="fr"><head><meta charset="utf-8"><title>${esc(ws.getSheetName())}</title><style>
@page { size: A4 ${opts.landscape ? 'landscape' : 'portrait'}; margin: 12mm; }
body { margin: 0; font: 11pt Calibri, Carlito, Arial, sans-serif; color: #000; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
.sheet { zoom: ${zoom.toFixed(4)}; }
table { border-collapse: collapse; table-layout: fixed; width: ${tableWidth}px; }
td { padding: 1px 3px; overflow: hidden; white-space: nowrap; vertical-align: bottom; ${opts.gridlines ? 'border: 1px solid #c8c8c8;' : ''} }
tr { page-break-inside: avoid; }
img { display: block; margin-top: 12px; max-width: 100%; page-break-inside: avoid; }
</style></head><body><div class="sheet"><table><colgroup>${cols_}</colgroup>${body.join('')}</table>${extras.join('')}</div></body></html>`;
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

export async function printDialog(api: FUniver): Promise<void> {
  const orientation = select(
    [
      ['portrait', t('print.portrait')],
      ['landscape', t('print.landscape')],
    ],
    'portrait',
  );
  const what = select(
    [
      ['sheet', t('print.activeSheet')],
      ['selection', t('print.selection')],
    ],
    'sheet',
  );
  const gridlines = checkbox(t('print.gridlines'), false);
  const fit = checkbox(t('print.fitWidth'), true);
  await openDialog({
    title: t('print.title'),
    testId: 'dialog-print',
    okLabel: t('print.ok'),
    body: box(t('print.help'), field(t('print.what'), what), field(t('print.orientation'), orientation), gridlines.el, fit.el),
    onOk: () => {
      const html = buildPrintHtml(api, {
        landscape: orientation.value === 'landscape',
        gridlines: gridlines.input.checked,
        fitWidth: fit.input.checked,
        selectionOnly: what.value === 'selection',
      });
      if (html) printHtml(html);
    },
  });
}

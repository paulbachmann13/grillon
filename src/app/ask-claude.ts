/**
 * « Demander à Claude » : deux fenêtres. La première prépare le message à coller dans Claude (la question et les cellules
 * sélectionnées, en texte), la seconde lit la réponse de Claude et écrit les cellules proposées, après les avoir montrées.
 * Grillon n'envoie rien : on copie, on colle. Logique dans src/data/ask-claude.ts.
 */
import type { FUniver } from '@univerjs/presets';
import type { FWorkbook, FWorksheet } from '@univerjs/sheets/facade';
import { buildPrompt, DEFAULT_MAX_COLS, DEFAULT_MAX_ROWS, MAX_EDITS, parseReply, toCellContent, type AskCell, type ReplyEdit } from '../data/ask-claude';
import { englishToFrench } from '../data/formula-locale';
import { t } from '../i18n';
import { toA1 } from '../io/xlsx/units';
import { box, checkbox, openDialog } from '../ui/dialog';
import { currentRegionRect, type Rect } from '../univer/region';

type Notify = (message: string, kind?: 'info' | 'success' | 'error') => void;

/** Lignes lues au plus (le message n'en garde que 100 ; le reste n'a pas besoin d'être lu). */
const READ_ROWS = DEFAULT_MAX_ROWS + 1;

function target(api: FUniver): { wb: FWorkbook; ws: FWorksheet; rect: Rect } | null {
  const wb = api.getActiveWorkbook() as unknown as FWorkbook | null;
  const ws = wb?.getActiveSheet() as unknown as FWorksheet | undefined;
  const range = wb?.getActiveRange();
  if (!wb || !ws || !range) return null;
  const sel: Rect = { r1: range.getRow(), c1: range.getColumn(), r2: range.getLastRow(), c2: range.getLastColumn() };
  // Une seule cellule : tout le tableau autour (comme Ctrl+A) ; une sélection plus grande : elle-même, sans les lignes vides du bas.
  if (sel.r1 === sel.r2 && sel.c1 === sel.c2) return { wb, ws, rect: currentRegionRect(ws, sel) };
  return { wb, ws, rect: { ...sel, r2: Math.max(sel.r1, Math.min(sel.r2, ws.getLastRow())), c2: Math.max(sel.c1, Math.min(sel.c2, ws.getLastColumn())) } };
}

/** Cellules de la zone : ce qu'on y voit, la formule (en français) et si c'est un nombre. */
function readCells(ws: FWorksheet, rect: Rect): AskCell[][] {
  const rows = Math.min(rect.r2 - rect.r1 + 1, READ_ROWS);
  const cols = rect.c2 - rect.c1 + 1;
  const range = ws.getRange(rect.r1, rect.c1, rows, cols);
  const shown = range.getDisplayValues();
  const formulas = range.getFormulas();
  const data = range.getCellDataGrid() as ({ v?: unknown; t?: number } | null | undefined)[][];
  return Array.from({ length: rows }, (_, i) =>
    Array.from({ length: cols }, (_c, j) => {
      const f = formulas[i]?.[j];
      const cell = data[i]?.[j];
      const numeric = typeof cell?.v === 'number' || cell?.t === 2 || cell?.t === 3;
      return { shown: shown[i]?.[j] ?? '', numeric, ...(f ? { formula: englishToFrench(f) } : {}) };
    }),
  );
}

/** Copie dans le presse-papiers ; faux si le navigateur refuse (le texte reste alors visible pour le copier à la main). */
async function copyText(text: string, fallback: HTMLTextAreaElement): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    fallback.focus();
    fallback.select();
    try {
      return document.execCommand('copy');
    } catch {
      return false;
    }
  }
}

function textarea(rows: number, placeholder = '', readOnly = false): HTMLTextAreaElement {
  const el = document.createElement('textarea');
  el.rows = rows;
  el.placeholder = placeholder;
  el.readOnly = readOnly;
  el.className = 'g-ask-text';
  return el;
}

function note(text: string, className = 'g-start-sub'): HTMLParagraphElement {
  const p = document.createElement('p');
  p.className = className;
  p.textContent = text;
  return p;
}

/** Fenêtre 1 : la question, ce qui sera copié (visible en entier), et le bouton « Copier le message ». */
export async function askClaudeDialog(api: FUniver, notify: Notify): Promise<void> {
  const sel = target(api);
  if (!sel) return void notify(t('askClaude.noData'), 'error');
  const { ws, rect } = sel;
  const cells = readCells(ws, rect);
  if (!cells.some((row) => row.some((c) => c.shown || c.formula))) return void notify(t('askClaude.noData'), 'error');
  const totalRows = rect.r2 - rect.r1 + 1;
  const totalCols = rect.c2 - rect.c1 + 1;
  // Titres : première ligne faite uniquement de textes non vides (pas de nombre, pas de formule).
  const looksLikeHeader = cells.length > 1 && cells[0]!.every((c) => c.shown && !c.numeric && !c.formula);

  const question = textarea(3, t('askClaude.questionPlaceholder'));
  question.dataset.testid = 'ask-question';
  const hide = checkbox(t('askClaude.hideTexts'), true);
  hide.input.dataset.testid = 'ask-hide';
  const header = checkbox(t('askClaude.header'), looksLikeHeader);
  header.input.dataset.testid = 'ask-header';
  const summary = note('', 'g-ask-summary');
  summary.dataset.testid = 'ask-summary';
  const preview = textarea(10, '', true);
  preview.dataset.testid = 'ask-preview';

  const refresh = () => {
    const p = buildPrompt({
      question: question.value,
      sheetName: ws.getSheetName(),
      row: rect.r1,
      col: rect.c1,
      cells,
      totalRows,
      totalCols,
      hideTexts: hide.input.checked,
      header: header.input.checked,
    });
    preview.value = p.text;
    const base = p.truncated ? t('askClaude.summary', { rows: p.rows, totalRows: p.totalRows, cols: p.cols, totalCols: p.totalCols }) : t('askClaude.summaryAll', { rows: p.rows, cols: p.cols });
    summary.textContent = hide.input.checked && p.hiddenTexts ? `${base} ${t('askClaude.hidden', { count: p.hiddenTexts })}` : base;
  };
  for (const el of [question, hide.input, header.input]) el.addEventListener(el === question ? 'input' : 'change', refresh);
  refresh();

  const heading = document.createElement('h3');
  heading.className = 'g-ask-heading';
  heading.textContent = t('askClaude.preview');
  const label = document.createElement('label');
  label.className = 'g-field';
  const span = document.createElement('span');
  span.textContent = t('askClaude.question');
  label.append(span, question);

  await openDialog({
    title: t('askClaude.askTitle'),
    testId: 'dialog-ask-claude',
    wide: true,
    okLabel: t('askClaude.copy'),
    cancelLabel: t('askClaude.close'),
    body: box(note(t('askClaude.askIntro')), label, hide.el, header.el, heading, summary, preview),
    onOk: async () => {
      if (!question.value.trim()) return t('askClaude.needQuestion');
      refresh();
      if (!(await copyText(preview.value, preview))) return t('askClaude.copyFailed');
      notify(t('askClaude.copied'), 'success');
    },
  });
}

interface Row {
  edit: ReplyEdit;
  content: ReturnType<typeof toCellContent>;
  before: string;
  after: string;
}

/** Ce que la cellule contient (formule en français, sinon le texte affiché). */
function currentText(ws: FWorksheet, row: number, col: number): string {
  const range = ws.getRange(row, col);
  const f = range.getFormula();
  return f ? englishToFrench(f) : range.getDisplayValue();
}

/** Fenêtre 2 : la réponse de Claude collée dans un champ ; aperçu avant/après ; « Appliquer » écrit tout en une seule annulation. */
export async function pasteReplyDialog(api: FUniver, notify: Notify): Promise<void> {
  const wb = api.getActiveWorkbook() as unknown as FWorkbook | null;
  const ws = wb?.getActiveSheet() as unknown as FWorksheet | undefined;
  if (!wb || !ws) return;
  const maxRows = ws.getMaxRows();
  const maxCols = ws.getMaxColumns();

  const reply = textarea(7, t('askClaude.pastePlaceholder'));
  reply.dataset.testid = 'paste-reply';
  const table = document.createElement('div');
  table.className = 'g-ask-edits';
  table.dataset.testid = 'paste-preview';
  const info = note('', 'g-ask-summary');
  let rows: Row[] = [];

  const refresh = () => {
    const parsed = parseReply(reply.value);
    const inside = parsed.edits.filter((e) => e.row < maxRows && e.col < maxCols);
    const outside = parsed.edits.length - inside.length;
    rows = inside.map((edit) => {
      const content = toCellContent(edit.content);
      return { edit, content, before: currentText(ws, edit.row, edit.col) || t('askClaude.empty'), after: content === null ? t('askClaude.cleared') : edit.content };
    });
    table.textContent = '';
    if (rows.length) {
      const grid = document.createElement('table');
      const head = grid.createTHead().insertRow();
      for (const h of [t('askClaude.cell'), t('askClaude.before'), t('askClaude.after')]) head.insertCell().textContent = h;
      const body = grid.createTBody();
      for (const r of rows) {
        const tr = body.insertRow();
        tr.insertCell().textContent = toA1(r.edit.row, r.edit.col);
        tr.insertCell().textContent = r.before;
        tr.insertCell().textContent = r.after;
      }
      table.append(grid);
    }
    const lines: string[] = [];
    if (rows.length) lines.push(t('askClaude.count', { count: rows.length }));
    else if (reply.value.trim()) lines.push(t('askClaude.nothing'));
    if (parsed.ignored.length) lines.push(t('askClaude.ignored', { count: parsed.ignored.length }));
    if (outside) lines.push(t('askClaude.outside', { count: outside }));
    info.textContent = lines.join(' ');
    if (parsed.edits.length >= MAX_EDITS) info.textContent += ` (${MAX_EDITS} cellules au plus)`;
  };
  reply.addEventListener('input', refresh);
  refresh();

  await openDialog({
    title: t('askClaude.pasteTitle'),
    testId: 'dialog-paste-reply',
    wide: true,
    okLabel: t('askClaude.apply'),
    cancelLabel: t('askClaude.close'),
    body: box(note(t('askClaude.pasteIntro')), reply, info, table),
    onOk: async () => {
      if (!rows.length) return t('askClaude.nothing');
      // Une seule commande (donc une seule annulation) ; les cellules non citées ne bougent pas.
      const value: Record<number, Record<number, object>> = {};
      let r1 = Infinity;
      let r2 = -1;
      let c1 = Infinity;
      let c2 = -1;
      for (const { edit, content } of rows) {
        r1 = Math.min(r1, edit.row);
        r2 = Math.max(r2, edit.row);
        c1 = Math.min(c1, edit.col);
        c2 = Math.max(c2, edit.col);
        const cell =
          content === null ? { v: null, f: null, si: null, p: null }
          : 'f' in content ? { f: content.f, si: null, v: null, p: null }
          : { v: typeof content.v === 'boolean' ? (content.v ? 1 : 0) : content.v, t: typeof content.v === 'number' ? 2 : typeof content.v === 'boolean' ? 3 : 1, f: null, si: null, p: null };
        (value[edit.row] ??= {})[edit.col] = cell;
      }
      await api.executeCommand('sheet.command.set-range-values', { value, range: { startRow: r1, endRow: r2, startColumn: c1, endColumn: c2 }, unitId: wb.getId(), subUnitId: ws.getSheetId() });
      notify(t('askClaude.applied', { count: rows.length }), 'success');
    },
  });
}

/**
 * Formules → Audit de formules : « Vérification des erreurs » et « Évaluer la formule », comme Excel.
 * Les sous-expressions sont calculées par le moteur d'Univer (executeFormulas), dans le contexte de
 * la cellule évaluée.
 */
import type { FUniver } from '@univerjs/presets';
import type { FWorkbook, FWorksheet } from '@univerjs/sheets/facade';
import { nextStep } from '../data/evaluate-formula';
import { FORMULA_ERRORS, formulaLiteral, frenchError, isErrorValue, numberFromText, TEXT_NUMBER, type ErrorInfo } from '../data/formula-errors';
import { t } from '../i18n';
import { toA1 } from '../io/xlsx/units';
import { GRILLON_CUSTOM_KEY } from '../model/univer-types';
import { alertDialog, box, openDialog } from '../ui/dialog';

type Formula = {
  executeFormulas(map: Record<string, Record<string, Record<number, Record<number, string[]>>>>, timeout?: number): Promise<Record<string, Record<string, Record<number, Record<number, { value: unknown }[]>>>>>;
};

/** Calcule des sous-expressions dans le contexte de la cellule (ligne, colonne). */
async function compute(api: FUniver, ws: FWorksheet, row: number, col: number, exprs: string[]): Promise<unknown[]> {
  const wb = api.getActiveWorkbook() as unknown as FWorkbook;
  const unit = wb.getId();
  const sheet = ws.getSheetId();
  const formula = (api as unknown as { getFormula(): Formula }).getFormula();
  const res = await formula.executeFormulas({ [unit]: { [sheet]: { [row]: { [col]: [`=${exprs[0]}`, ...exprs] } } } }, 5000);
  return (res[unit]?.[sheet]?.[row]?.[col] ?? []).slice(1).map((r) => r.value);
}

function button(label: string, onClick: () => void, testId?: string): HTMLButtonElement {
  const b = document.createElement('button');
  b.type = 'button';
  b.textContent = label;
  if (testId) b.dataset.testid = testId;
  b.addEventListener('click', onClick);
  return b;
}

const cellName = (ws: FWorksheet, row: number, col: number) => `${ws.getSheetName()}!${toA1(row, col)}`;

// --- Évaluer la formule ---

export async function evaluateFormulaDialog(api: FUniver, at?: { ws: FWorksheet; row: number; col: number }): Promise<void> {
  const wb = api.getActiveWorkbook() as unknown as FWorkbook | null;
  const range = wb?.getActiveRange();
  const ws = at?.ws ?? (wb?.getActiveSheet() as unknown as FWorksheet | undefined);
  if (!ws || (!at && !range)) return;
  const row = at?.row ?? range!.getRow();
  const col = at?.col ?? range!.getColumn();
  const formula = ws.getRange(row, col).getFormula();
  if (!formula) return alertDialog(t('audit.evaluate'), t('formulas.noFormula'));
  const original = formula.replace(/^=/, '');
  let expr = original;

  const ref = document.createElement('p');
  ref.textContent = `${t('audit.reference')} ${cellName(ws, row, col)}`;
  const view = document.createElement('div');
  view.className = 'g-eval-expr';
  view.dataset.testid = 'evaluate-expression';
  const status = document.createElement('p');
  status.className = 'g-eval-status';
  const evalBtn = button(t('audit.evaluateStep'), () => void step(), 'evaluate-step');
  const restart = button(t('audit.restart'), () => {
    expr = original;
    render();
  });
  const actions = document.createElement('div');
  actions.className = 'g-eval-actions';
  actions.append(evalBtn, restart);

  function render(): void {
    const span = nextStep(expr);
    view.replaceChildren();
    view.append('=');
    if (!span) {
      view.append(expr);
      status.textContent = t('audit.done');
      evalBtn.disabled = true;
      return;
    }
    const mark = document.createElement('u');
    mark.textContent = expr.slice(span[0], span[1]);
    view.append(expr.slice(0, span[0]), mark, expr.slice(span[1]));
    status.textContent = t('audit.underlined');
    evalBtn.disabled = false;
  }

  async function step(): Promise<void> {
    const span = nextStep(expr);
    if (!span) return;
    evalBtn.disabled = true;
    try {
      const part = expr.slice(span[0], span[1]);
      const [value] = await compute(api, ws!, row, col, [part]);
      // Une plage d'une seule cellule renvoie sa valeur, pas un tableau.
      const v = Array.isArray(value) && value.length === 1 && Array.isArray(value[0]) && value[0].length === 1 ? value[0][0] : value;
      expr = expr.slice(0, span[0]) + formulaLiteral(v) + expr.slice(span[1]);
    } catch (e) {
      console.warn('Évaluation impossible', e);
      expr = expr.slice(0, span[0]) + '#VALUE!' + expr.slice(span[1]);
    }
    render();
  }

  render();
  await openDialog({ title: t('audit.evaluate'), testId: 'dialog-evaluate', wide: true, body: box(ref, view, status, actions), okLabel: t('dialog.close'), onOk: () => true });
}

// --- Vérification des erreurs ---

interface Issue {
  row: number;
  col: number;
  info: ErrorInfo;
  textNumber?: number;
}

type Custom = { [GRILLON_CUSTOM_KEY]?: { ignoredErrors?: string[] } & Record<string, unknown> };

function ignoredOf(ws: FWorksheet): Set<string> {
  return new Set(((ws.getCustomMetadata() ?? {}) as Custom)[GRILLON_CUSTOM_KEY]?.ignoredErrors ?? []);
}

function ignore(ws: FWorksheet, row: number, col: number): void {
  const custom = (ws.getCustomMetadata() ?? {}) as Custom;
  const grillon = { ...(custom[GRILLON_CUSTOM_KEY] ?? {}) };
  grillon.ignoredErrors = [...new Set([...(grillon.ignoredErrors ?? []), `${row},${col}`])];
  ws.setCustomMetadata({ ...custom, [GRILLON_CUSTOM_KEY]: grillon });
}

/** Cellules en erreur et nombres stockés sous forme de texte, en partant de la cellule active. */
export function findIssues(ws: FWorksheet, from: { row: number; col: number }): Issue[] {
  const ignored = ignoredOf(ws);
  const out: Issue[] = [];
  const matrix = (ws.getSheet() as unknown as { getCellMatrix(): { forValue(cb: (r: number, c: number, cell: { v?: unknown; f?: string; si?: string; t?: number }) => void): void } }).getCellMatrix();
  matrix.forValue((r, c, cell) => {
    if (!cell || ignored.has(`${r},${c}`)) return;
    if (isErrorValue(cell.v)) out.push({ row: r, col: c, info: FORMULA_ERRORS[cell.v] });
    else if (!cell.f && !cell.si && cell.t === 1 && typeof cell.v === 'string') {
      const n = numberFromText(cell.v);
      if (n !== null) out.push({ row: r, col: c, info: TEXT_NUMBER, textNumber: n });
    }
  });
  out.sort((a, b) => a.row - b.row || a.col - b.col);
  const start = out.findIndex((i) => i.row > from.row || (i.row === from.row && i.col >= from.col));
  return start > 0 ? [...out.slice(start), ...out.slice(0, start)] : out;
}

export async function errorCheckingDialog(api: FUniver): Promise<void> {
  const wb = api.getActiveWorkbook() as unknown as FWorkbook | null;
  const ws = wb?.getActiveSheet() as unknown as FWorksheet | undefined;
  const range = wb?.getActiveRange();
  if (!wb || !ws) return;
  const issues = findIssues(ws, { row: range?.getRow() ?? 0, col: range?.getColumn() ?? 0 });
  if (!issues.length) return alertDialog(t('audit.errorChecking'), t('audit.noErrors'));
  let index = 0;

  const where = document.createElement('p');
  where.className = 'g-eval-where';
  where.dataset.testid = 'error-cell';
  const content = document.createElement('div');
  content.className = 'g-eval-expr';
  const title = document.createElement('p');
  title.className = 'g-eval-title';
  title.dataset.testid = 'error-title';
  const help = document.createElement('p');
  const side = document.createElement('div');
  side.className = 'g-eval-actions';
  const nav = document.createElement('div');
  nav.className = 'g-eval-actions';
  const close = () => (where.closest('.g-dialog')?.querySelector('.g-dialog-actions button.primary') as HTMLButtonElement | null)?.click();

  const evalBtn = button(t('audit.showSteps'), () => void evaluateFormulaDialog(api, { ws, row: issues[index].row, col: issues[index].col }));
  const convert = button(t('audit.convert'), () => {
    const issue = issues[index];
    ws.getRange(issue.row, issue.col).setValue(issue.textNumber!);
    drop();
  }, 'error-convert');
  const ignoreBtn = button(t('audit.ignore'), () => {
    ignore(ws, issues[index].row, issues[index].col);
    drop();
  }, 'error-ignore');
  const edit = button(t('audit.edit'), close);
  const prev = button(t('audit.previous'), () => go(index - 1), 'error-previous');
  const next = button(t('audit.next'), () => go(index + 1), 'error-next');
  side.append(evalBtn, convert, ignoreBtn, edit);
  nav.append(prev, next);

  function drop(): void {
    issues.splice(index, 1);
    if (!issues.length) {
      close();
      void alertDialog(t('audit.errorChecking'), t('audit.noErrors'));
      return;
    }
    go(Math.min(index, issues.length - 1));
  }

  function go(i: number): void {
    index = (i + issues.length) % issues.length;
    const issue = issues[index];
    wb!.setActiveRange(ws!.getRange(issue.row, issue.col));
    const cell = ws!.getRange(issue.row, issue.col);
    const f = cell.getFormula();
    where.textContent = t('audit.errorIn', { cell: toA1(issue.row, issue.col) });
    content.textContent = f || String(cell.getValue() ?? '');
    title.textContent = issue.info.title + (issue.info.fr ? ` (${frenchError(Object.keys(FORMULA_ERRORS).find((k) => FORMULA_ERRORS[k] === issue.info) ?? '')})` : '');
    help.textContent = issue.info.help;
    evalBtn.hidden = !f;
    convert.hidden = issue.textNumber === undefined;
  }

  go(0);
  await openDialog({ title: t('audit.errorChecking'), testId: 'dialog-error-checking', wide: true, body: box(where, content, title, help, side, nav), okLabel: t('dialog.close'), onOk: () => true });
}

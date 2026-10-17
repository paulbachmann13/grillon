/**
 * Formules → Fenêtre Espion : surveille la valeur de cellules choisies, même d'une autre feuille.
 * La liste est gardée en mémoire pour la session (comme dans Excel, elle suit le classeur ouvert).
 */
import type { FUniver } from '@univerjs/presets';
import type { FWorkbook, FWorksheet } from '@univerjs/sheets/facade';
import { addWatch, removeWatch, sortByColumn, type WatchItem } from '../data/watch';
import { t } from '../i18n';
import { toA1 } from '../io/xlsx/units';
import { box, openDialog } from '../ui/dialog';

let watches: WatchItem[] = [];
let watchedUnit = '';

interface Row {
  item: WatchItem;
  cell: string;
  value: string;
  formula: string;
}

function sheetOf(wb: FWorkbook, item: WatchItem): FWorksheet | undefined {
  return (wb.getSheets() as unknown as FWorksheet[]).find((s) => s.getSheetId() === item.sheetId);
}

function rowsOf(wb: FWorkbook): Row[] {
  const out: Row[] = [];
  for (const item of watches) {
    const ws = sheetOf(wb, item);
    if (!ws) continue;
    const range = ws.getRange(item.row, item.col);
    out.push({ item: { ...item, sheetName: ws.getSheetName() }, cell: toA1(item.row, item.col), value: String(range.getDisplayValue?.() ?? range.getValue() ?? ''), formula: range.getFormula() ?? '' });
  }
  return out;
}

export async function watchWindow(api: FUniver): Promise<void> {
  const wb = api.getActiveWorkbook() as unknown as FWorkbook | null;
  if (!wb) return;
  if (watchedUnit !== wb.getId()) {
    watches = [];
    watchedUnit = wb.getId();
  }
  let selected: WatchItem | null = null;
  let sortKey: keyof Pick<Row, 'cell' | 'value' | 'formula'> | 'sheet' | null = null;
  let descending = false;

  const table = document.createElement('table');
  table.className = 'g-watch';
  table.dataset.testid = 'watch-table';
  const scroll = document.createElement('div');
  scroll.className = 'g-watch-scroll';
  scroll.append(table);
  const hint = document.createElement('p');
  hint.className = 'g-eval-status';
  hint.textContent = t('watch.hint');
  const actions = document.createElement('div');
  actions.className = 'g-eval-actions';
  const addBtn = document.createElement('button');
  addBtn.type = 'button';
  addBtn.dataset.testid = 'watch-add';
  const removeBtn = document.createElement('button');
  removeBtn.type = 'button';
  removeBtn.dataset.testid = 'watch-remove';
  removeBtn.textContent = t('watch.remove');
  actions.append(addBtn, removeBtn);

  const active = () => {
    const range = wb!.getActiveRange();
    const ws = wb!.getActiveSheet() as unknown as FWorksheet | undefined;
    return range && ws ? { ws, row: range.getRow(), col: range.getColumn() } : null;
  };

  function render(): void {
    let rows = rowsOf(wb!);
    if (sortKey) {
      const key = sortKey;
      rows = sortByColumn(rows, (r) => (key === 'sheet' ? r.item.sheetName : r[key]), descending);
    }
    const head = document.createElement('thead');
    const hr = document.createElement('tr');
    for (const [key, label] of [['sheet', t('watch.sheet')], ['cell', t('watch.cell')], ['value', t('watch.value')], ['formula', t('watch.formula')]] as const) {
      const th = document.createElement('th');
      th.textContent = label;
      th.addEventListener('click', () => {
        descending = sortKey === key ? !descending : false;
        sortKey = key;
        render();
      });
      hr.append(th);
    }
    head.append(hr);
    const body = document.createElement('tbody');
    for (const r of rows) {
      const tr = document.createElement('tr');
      tr.dataset.testid = 'watch-row';
      if (selected && r.item.sheetId === selected.sheetId && r.item.row === selected.row && r.item.col === selected.col) tr.className = 'selected';
      for (const text of [r.item.sheetName, r.cell, r.value, r.formula]) {
        const td = document.createElement('td');
        td.textContent = text;
        tr.append(td);
      }
      tr.addEventListener('click', () => {
        selected = r.item;
        render();
      });
      tr.addEventListener('dblclick', () => {
        const ws = sheetOf(wb!, r.item);
        if (!ws) return;
        wb!.setActiveSheet(ws);
        wb!.setActiveRange(ws.getRange(r.item.row, r.item.col));
        (table.closest('.g-dialog')?.querySelector('.g-dialog-actions button.primary') as HTMLButtonElement | null)?.click();
      });
      body.append(tr);
    }
    table.replaceChildren(head, body);
    const a = active();
    addBtn.textContent = a ? t('watch.addActive', { cell: `${a.ws.getSheetName()}!${toA1(a.row, a.col)}` }) : t('watch.add');
    addBtn.disabled = !a;
    removeBtn.disabled = !selected;
    if (!rows.length) {
      const empty = document.createElement('caption');
      empty.textContent = t('watch.empty');
      table.prepend(empty);
    }
  }

  addBtn.addEventListener('click', () => {
    const a = active();
    if (!a) return;
    watches = addWatch(watches, { sheetId: a.ws.getSheetId(), sheetName: a.ws.getSheetName(), row: a.row, col: a.col });
    render();
  });
  removeBtn.addEventListener('click', () => {
    if (!selected) return;
    watches = removeWatch(watches, selected);
    selected = null;
    render();
  });

  render();
  await openDialog({ title: t('watch.title'), testId: 'dialog-watch', wide: true, body: box(scroll, actions, hint), okLabel: t('dialog.close'), onOk: () => true });
}

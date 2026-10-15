/**
 * Analyse rapide (comme Excel) : un petit bouton apparaît sous le coin de la sélection ; il ouvre
 * un panneau pour mettre en forme, créer un graphique, ajouter des totaux ou des mini-graphiques.
 */
import { ICommandService, type FUniver } from '@univerjs/presets';
import type { FWorksheet } from '@univerjs/sheets/facade';
import { createElement } from 'react';
import { insertChart } from '../charts/chart-store';
import { quickSparklines } from '../charts/sparkline-view';
import type { AutoSumFunction } from '../data/autosum';
import { t } from '../i18n';
import { autoSum } from './basics';

const MAX_CELLS = 200_000;

interface QaItem {
  label: string;
  run: () => unknown;
}
interface QaTab {
  id: string;
  label: string;
  items: QaItem[];
}

function tabs(api: FUniver): QaTab[] {
  const ws = () => api.getActiveWorkbook()?.getActiveSheet() as unknown as (FWorksheet & {
    newConditionalFormattingRule(): { setRanges(r: object[]): unknown; setDataBar(c: object): { setRanges(r: object[]): { build(): object } }; setColorScale(c: object): { setRanges(r: object[]): { build(): object } } };
    addConditionalFormattingRule(rule: object): unknown;
  }) | undefined;
  const range = () => api.getActiveWorkbook()?.getActiveRange();
  const cf = (kind: 'bar' | 'scale') => () => {
    const sheet = ws();
    const r = range();
    if (!sheet || !r) return;
    const rule = sheet.newConditionalFormattingRule();
    const ranges = [r.getRange()];
    const built =
      kind === 'bar'
        ? rule.setDataBar({ min: { type: 'min' }, max: { type: 'max' }, positiveColor: '#638EC6', nativeColor: '#FF0000', isGradient: true, isShowValue: true } as never).setRanges(ranges).build()
        : rule.setColorScale({ config: [{ index: 0, color: '#F8696B', value: { type: 'min' } }, { index: 1, color: '#FFEB84', value: { type: 'percentile', value: 50 } }, { index: 2, color: '#63BE7B', value: { type: 'max' } }] } as never).setRanges(ranges).build();
    sheet.addConditionalFormattingRule(built);
  };
  const total = (fn: AutoSumFunction, direction: 'down' | 'right') => () => autoSum(api, fn, direction);
  const totals = (direction: 'down' | 'right', suffix: string): QaItem[] =>
    (['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'] as AutoSumFunction[]).map((fn) => ({ label: `${t(`basics.${{ SUM: 'sum', AVERAGE: 'average', COUNT: 'count', MAX: 'max', MIN: 'min' }[fn]}`)} ${suffix}`, run: total(fn, direction) }));
  return [
    {
      id: 'format',
      label: t('qa.format'),
      items: [
        { label: t('qa.dataBars'), run: cf('bar') },
        { label: t('qa.colorScale'), run: cf('scale') },
      ],
    },
    {
      id: 'charts',
      label: t('qa.charts'),
      items: [
        { label: t('qa.column'), run: () => insertChart(api, 'column') },
        { label: t('qa.line'), run: () => insertChart(api, 'line-markers') },
        { label: t('qa.pie'), run: () => insertChart(api, 'pie') },
        { label: t('qa.bar'), run: () => insertChart(api, 'bar') },
        { label: t('qa.scatter'), run: () => insertChart(api, 'scatter') },
        { label: t('qa.moreCharts'), run: () => api.executeCommand('grillon.insert.chart.all') },
      ],
    },
    { id: 'totals', label: t('qa.totals'), items: [...totals('down', t('qa.below')), ...totals('right', t('qa.right'))] },
    {
      id: 'tables',
      label: t('qa.tables'),
      items: [
        { label: t('qa.table'), run: () => api.executeCommand('grillon.data.table') },
        { label: t('qa.pivot'), run: () => api.executeCommand('grillon.pivot.create') },
      ],
    },
    {
      id: 'sparklines',
      label: t('qa.sparklines'),
      items: [
        { label: t('spark.line'), run: () => quickSparklines(api, 'line') },
        { label: t('spark.column'), run: () => quickSparklines(api, 'column') },
        { label: t('spark.winloss'), run: () => quickSparklines(api, 'winloss') },
      ],
    },
  ];
}

let panel: HTMLElement | null = null;

function closePanel(): void {
  panel?.remove();
  panel = null;
  document.removeEventListener('pointerdown', onOutside, true);
  document.removeEventListener('keydown', onKey, true);
}
function onOutside(e: Event): void {
  if (panel && !panel.contains(e.target as Node)) closePanel();
}
function onKey(e: KeyboardEvent): void {
  if (e.key === 'Escape') closePanel();
}

/** Ouvre le panneau d'Analyse rapide près de l'élément `anchor`. */
export function openQuickAnalysis(api: FUniver, anchor: Element): void {
  closePanel();
  const all = tabs(api);
  const el = document.createElement('div');
  el.className = 'g-qa-panel';
  el.dataset.testid = 'quick-analysis';
  el.setAttribute('role', 'dialog');
  el.setAttribute('aria-label', t('qa.title'));
  const bar = document.createElement('div');
  bar.className = 'g-qa-tabs';
  bar.setAttribute('role', 'tablist');
  const list = document.createElement('div');
  list.className = 'g-qa-items';
  const show = (id: string) => {
    for (const b of bar.querySelectorAll('button')) {
      b.classList.toggle('selected', b.dataset.tab === id);
      b.setAttribute('aria-selected', String(b.dataset.tab === id));
    }
    list.replaceChildren(
      ...all
        .find((x) => x.id === id)!
        .items.map((item) => {
          const b = document.createElement('button');
          b.type = 'button';
          b.setAttribute('role', 'menuitem');
          b.textContent = item.label;
          b.addEventListener('click', () => {
            closePanel();
            void item.run();
          });
          return b;
        }),
    );
  };
  for (const tab of all) {
    const b = document.createElement('button');
    b.type = 'button';
    b.setAttribute('role', 'tab');
    b.dataset.tab = tab.id;
    b.textContent = tab.label;
    b.addEventListener('click', () => show(tab.id));
    bar.append(b);
  }
  el.append(bar, list);
  document.body.append(el);
  const rect = anchor.getBoundingClientRect();
  const w = el.offsetWidth;
  const h = el.offsetHeight;
  el.style.left = `${Math.max(8, Math.min(rect.left, window.innerWidth - w - 8))}px`;
  el.style.top = `${Math.max(8, Math.min(rect.bottom + 4, window.innerHeight - h - 8))}px`;
  panel = el;
  show(all[0]!.id);
  document.addEventListener('pointerdown', onOutside, true);
  document.addEventListener('keydown', onKey, true);
  (bar.querySelector('button') as HTMLButtonElement | null)?.focus();
}

export function installQuickAnalysis(api: FUniver): void {
  let popup: { dispose(): void } | null | undefined;
  let timer: ReturnType<typeof setTimeout> | undefined;
  const button = () =>
    createElement(
      'button',
      {
        type: 'button',
        className: 'g-qa-btn',
        title: t('qa.title'),
        'aria-label': t('qa.title'),
        'data-testid': 'quick-analysis-button',
        onPointerDown: (e: Event) => e.stopPropagation(),
        onClick: (e: { currentTarget: Element }) => openQuickAnalysis(api, e.currentTarget),
      },
      '⚡',
    );
  const refresh = () => {
    popup?.dispose();
    popup = null;
    const wb = api.getActiveWorkbook();
    const range = wb?.getActiveRange();
    const ws = wb?.getActiveSheet();
    if (!wb || !range || !ws) return;
    const rows = range.getLastRow() - range.getRow() + 1;
    const cols = range.getLastColumn() - range.getColumn() + 1;
    if (rows * cols < 2 || rows * cols > MAX_CELLS || rows >= ws.getMaxRows() || cols >= ws.getMaxColumns()) return;
    if (document.querySelector('.g-dialog-backdrop')) return;
    // Au moins une cellule non vide dans la sélection.
    const probe = range.getCellDataGrid();
    if (!probe.some((r) => r.some((c) => c && c.v !== undefined && c.v !== null && c.v !== ''))) return;
    popup = ws.getRange(range.getLastRow(), range.getLastColumn()).attachPopup({ componentKey: button as never, isReactComponent: true, direction: 'bottom', offset: [0, 4] } as never) as { dispose(): void } | null | undefined;
  };
  const schedule = () => {
    clearTimeout(timer);
    timer = setTimeout(refresh, 250);
  };
  // Toute sélection (souris, clavier, programme) passe par cette opération.
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  injector.get<ICommandService>(ICommandService).onCommandExecuted((info) => {
    if (info.id === 'sheet.operation.set-selections') schedule();
  });
  api.addEvent(api.Event.ActiveSheetChanged, schedule);
}

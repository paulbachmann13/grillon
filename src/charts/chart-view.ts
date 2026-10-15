/**
 * Graphiques dans la feuille : un « DOM flottant » Univer (déplaçable, redimensionnable,
 * enregistré avec le classeur) qui affiche un graphique ECharts. La définition du graphique
 * (ChartSpec) est stockée dans les données du DOM flottant.
 */
import { IconManager } from '@univerjs/preset-sheets-core';
import { CommandType, ICommandService, type FUniver } from '@univerjs/presets';
import { createElement, useEffect, useRef, useState, type ComponentType } from 'react';
import { t } from '../i18n';
import { chartDialog } from './chart-dialogs';
import { CHART_COMPONENT, insertChart, isDark, loadECharts, readRef, type ChartData } from './chart-store';
import { CHART_KINDS, KIND_GROUPS } from './kinds';
import { buildEChartsOption, dataTableOf, resolveSeries } from './spec';

export { CHART_COMPONENT, insertChart, readRef };
export type { ChartData };
export { removeChart, updateChart } from './chart-store';

type EChartsModule = typeof import('echarts');

export function createChartComponent(api: FUniver): ComponentType<{ data?: ChartData }> {
  return function ChartView(props: { data?: ChartData }) {
    const holder = useRef<HTMLDivElement>(null);
    const spec = props.data?.spec;
    const key = JSON.stringify(spec ?? null);
    const [table, setTable] = useState<ReturnType<typeof dataTableOf> | null>(null);
    useEffect(() => {
      if (!spec || !holder.current) return;
      let disposed = false;
      let chart: ReturnType<EChartsModule['init']> | null = null;
      let timer: ReturnType<typeof setTimeout> | undefined;
      const render = () => {
        if (!chart) return;
        const visibleOnly = spec.plotVisibleOnly !== false;
        const series = resolveSeries(spec, (ref, display) => readRef(api, ref, { display, visibleOnly }));
        chart.setOption(buildEChartsOption(spec, series, isDark()) as never, true);
        setTable(spec.dataTable ? dataTableOf(series) : null);
      };
      const schedule = () => {
        clearTimeout(timer);
        timer = setTimeout(render, 120);
      };
      void loadECharts().then((ec) => {
        if (disposed || !holder.current) return;
        chart = ec.init(holder.current, undefined, { renderer: 'canvas' });
        render();
      });
      // Les données changent : saisie, recalcul des formules, lignes insérées, masquées…
      const sub = api.addEvent(api.Event.CommandExecuted, (e) => {
        if (e.type === 2 && /set-range-values|formula-calculation-result|insert-row|remove-row|insert-col|remove-col|move-|set-row-visible|set-col-visible|set-rows-visible|set-cols-visible|set-row-hidden|set-col-hidden/.test(e.id)) schedule();
      });
      const ro = new ResizeObserver(() => chart?.resize());
      ro.observe(holder.current);
      const mq = matchMedia('(prefers-color-scheme: dark)');
      mq.addEventListener('change', schedule);
      return () => {
        disposed = true;
        clearTimeout(timer);
        sub.dispose();
        ro.disconnect();
        mq.removeEventListener('change', schedule);
        chart?.dispose();
      };
    }, [key]);
    return createElement(
      'div',
      { className: 'g-chart', 'data-testid': 'chart', style: { background: spec?.background, border: spec?.border === false ? 'none' : undefined } },
      createElement('div', { ref: holder, className: 'g-chart-canvas', onDoubleClick: () => props.data && void chartDialog(api, { chartId: props.data.chartId }) }),
      table &&
        createElement(
          'div',
          { className: 'g-chart-table', 'data-testid': 'chart-table' },
          createElement(
            'table',
            null,
            createElement('thead', null, createElement('tr', null, createElement('th', null), ...table.header.map((h, i) => createElement('th', { key: i }, h)))),
            createElement(
              'tbody',
              null,
              ...table.rows.map((r, i) =>
                createElement('tr', { key: i }, createElement('th', null, createElement('span', { className: 'g-swatch', style: { background: r.color } }), r.name), ...r.cells.map((c, k) => createElement('td', { key: k }, c))),
              ),
            ),
          ),
        ),
      createElement(
        'button',
        {
          type: 'button',
          className: 'g-chart-edit',
          title: t('charts.edit'),
          'aria-label': t('charts.edit'),
          onPointerDown: (e: Event) => e.stopPropagation(),
          onClick: () => props.data && void chartDialog(api, { chartId: props.data.chartId }),
        },
        '⚙',
      ),
    );
  };
}

/** Boîte « Insérer un graphique » (tous les types) pour la sélection courante. */
export function openInsertChartDialog(api: FUniver, kindId = 'column'): Promise<void> {
  return chartDialog(api, { kindId });
}

function icon(path: string): ComponentType<Record<string, unknown>> {
  return (props) =>
    createElement('svg', { viewBox: '0 0 16 16', width: '1em', height: '1em', fill: 'none', stroke: 'currentColor', strokeWidth: 1.3, strokeLinecap: 'round', strokeLinejoin: 'round', ...props }, createElement('path', { d: path }));
}

export function installCharts(api: FUniver): void {
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const commands = injector.get<ICommandService>(ICommandService);
  injector.get<IconManager>(IconManager).register({
    GrillonChartIcon: icon('M2 14h12M4 11V7M7 11V3M10 11V8M13 11V5'),
    GrillonChartsAllIcon: icon('M2 2h5v5H2zM9 2h5v5H9zM2 9h5v5H2zM9 9h5v5H9z'),
  } as never);
  const path = ['ribbon.insert', 'ribbon.insert.media'];
  const id = 'grillon.insert.chart';
  api.createSubmenu({ id, icon: 'GrillonChartIcon', title: t('charts.insert'), tooltip: t('charts.insert') }).appendTo(path);
  let order = 0;
  for (const kind of CHART_KINDS) {
    const itemId = `${id}.${kind.id}`;
    if (!commands.hasCommand(itemId)) commands.registerCommand({ id: itemId, type: CommandType.COMMAND, handler: () => insertChart(api, kind) !== null });
    // Le libellé reprend le groupe : le menu est plat dans le ruban en grille.
    const group = KIND_GROUPS.find((g) => g.id === kind.group)?.label ?? '';
    api.createMenu({ id: itemId, title: kind.label.startsWith(group) || group === 'Combo' ? kind.label : `${group} : ${kind.label}`, action: itemId, order: order++ }).appendTo([...path, id]);
  }
  const allId = `${id}.all`;
  commands.registerCommand({ id: allId, type: CommandType.COMMAND, handler: () => (void openInsertChartDialog(api), true) });
  api.createMenu({ id: allId, icon: 'GrillonChartsAllIcon', title: t('charts.allCharts'), action: allId, order: -1 }).appendTo([...path, id]);
}

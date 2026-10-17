/**
 * Boîtes de dialogue des graphiques : « Insérer un graphique » (tous les types, avec aperçu)
 * et « Format du graphique » (données, éléments, axes, séries, mise en forme), comme Excel.
 * L'aperçu à droite se met à jour à chaque réglage ; « Appliquer » enregistre dans la feuille.
 */
import type { FUniver } from '@univerjs/presets';
import type { FWorksheet } from '@univerjs/sheets/facade';
import { fr } from '../i18n/fr';
import { t } from '../i18n';
import { alertDialog, box, checkbox, field, group, openDialog, select, textInput } from '../ui/dialog';
import { findChart, insertChartSpec, isDark, loadECharts, readRef, removeChart, specFromSelection, updateChart, workbook } from './chart-store';
import { applyKind, CHART_KINDS, kindOf, KIND_GROUPS } from './kinds';
import {
  buildEChartsOption,
  isStatChart,
  DEFAULT_COLORS,
  PALETTES,
  parseSheetRef,
  resolveSeries,
  seriesFromRange,
  sheetRef,
  type AxisSpec,
  type CellScalar,
  type ChartSpec,
  type DataLabelsSpec,
  type LabelPosition,
  type MarkerShape,
  type SeriesSpec,
  type TrendlineSpec,
} from './spec';

const L = (key: keyof typeof fr.chartUi): string => t(`chartUi.${key}`);

const isPieLike = (spec: ChartSpec) => spec.type === 'pie' || spec.type === 'doughnut';
const isXY = (spec: ChartSpec) => spec.type === 'scatter' || spec.type === 'bubble';
const hasAxes = (spec: ChartSpec) => !isPieLike(spec) && spec.type !== 'radar';

/** Référence « Feuil1!$A$1:$B$3 » (ou « B2 ») → référence complète avec nom de feuille ; null si invalide. */
export function normalizeRef(value: string, defaultSheet: string): string | null {
  const p = parseSheetRef(value);
  return p ? sheetRef(p.sheet ?? defaultSheet, p.startRow, p.startColumn, p.endRow, p.endColumn) : null;
}

function shiftRef(ref: string | undefined, dc: number): string | undefined {
  const p = ref ? parseSheetRef(ref) : null;
  return p && ref ? sheetRef(p.sheet ?? '', p.startRow, p.startColumn + dc, p.endRow, p.endColumn + dc).replace(/^!/, '') : ref;
}

/** Retire les objets vides et les valeurs `undefined` avant d'enregistrer. */
export function pruneSpec<T>(value: T): T {
  const clean = (v: unknown): unknown => {
    if (Array.isArray(v)) return v.map(clean);
    if (v && typeof v === 'object') {
      const out: Record<string, unknown> = {};
      for (const [k, x] of Object.entries(v)) {
        const c = clean(x);
        if (c === undefined || (c && typeof c === 'object' && !Array.isArray(c) && !Object.keys(c).length)) continue;
        out[k] = c;
      }
      return out;
    }
    return v === '' ? undefined : v;
  };
  return clean(value) as T;
}

function numberInput(value: number | undefined, placeholder = ''): HTMLInputElement {
  const i = document.createElement('input');
  i.type = 'number';
  i.step = 'any';
  i.placeholder = placeholder;
  if (value !== undefined) i.value = String(value);
  return i;
}

function colorInput(value: string): HTMLInputElement {
  const i = document.createElement('input');
  i.type = 'color';
  i.value = value;
  return i;
}

export interface ChartDialogOptions {
  /** Graphique à modifier (sinon : insertion à partir de la sélection). */
  chartId?: string;
  kindId?: string;
}

export async function chartDialog(api: FUniver, opts: ChartDialogOptions): Promise<void> {
  const existing = opts.chartId ? findChart(api, opts.chartId) : null;
  const initial = existing ? structuredClone(existing.data.spec) : specFromSelection(api, opts.kindId ?? 'column');
  if (!initial) return alertDialog(L('insertTitle'), L('needData'));
  let work: ChartSpec = initial;
  const inserting = !existing;
  const activeSheet = () => workbook(api)?.getActiveSheet()?.getSheetName() ?? 'Feuil1';
  const defaultSheet = existing ? (parseSheetRef(work.series[0]?.valuesRef ?? '')?.sheet ?? activeSheet()) : activeSheet();

  // --- Aperçu ---
  const preview = document.createElement('div');
  preview.className = 'g-chart-preview';
  preview.dataset.testid = 'chart-preview';
  let chart: { setOption(o: never, notMerge: boolean): void; resize(): void; dispose(): void } | null = null;
  let timer: ReturnType<typeof setTimeout> | undefined;
  const renderPreview = () => {
    if (!chart) return;
    const visibleOnly = work.plotVisibleOnly !== false;
    const series = resolveSeries(work, (ref, display) => readRef(api, ref, { display, visibleOnly, defaultSheet }));
    chart.setOption(buildEChartsOption(work, series, isDark()) as never, true);
  };
  const refresh = () => {
    clearTimeout(timer);
    timer = setTimeout(renderPreview, 60);
  };
  void loadECharts().then((ec) => {
    chart = ec.init(preview, undefined, { renderer: 'canvas' }) as never;
    renderPreview();
  });

  // --- Constructeurs de champs liés à la spécification ---
  const text = (label: string, get: () => string | undefined, set: (v: string | undefined) => void, placeholder = '') => {
    const i = textInput(get() ?? '', placeholder);
    i.addEventListener('input', () => {
      set(i.value || undefined);
      refresh();
    });
    return field(label, i);
  };
  const num = (label: string, get: () => number | undefined, set: (v: number | undefined) => void, placeholder = '') => {
    const i = numberInput(get(), placeholder);
    i.addEventListener('input', () => {
      set(i.value === '' || !Number.isFinite(Number(i.value)) ? undefined : Number(i.value));
      refresh();
    });
    return field(label, i);
  };
  const check = (label: string, get: () => boolean | undefined, set: (v: boolean | undefined) => void, onChange?: () => void) => {
    const c = checkbox(label, !!get());
    c.input.addEventListener('change', () => {
      set(c.input.checked || undefined);
      refresh();
      onChange?.();
    });
    return c.el;
  };
  const choose = (label: string, options: [string, string][], get: () => string | undefined, set: (v: string | undefined) => void, onChange?: () => void) => {
    const s = select(options, get() ?? options[0]![0]);
    s.addEventListener('change', () => {
      set(s.value || undefined);
      refresh();
      onChange?.();
    });
    return field(label, s);
  };
  const color = (label: string, get: () => string | undefined, fallback: () => string, set: (v: string | undefined) => void, resetLabel?: string) => {
    const i = colorInput(get() ?? fallback());
    i.addEventListener('input', () => {
      set(i.value);
      refresh();
    });
    const wrap = document.createElement('div');
    wrap.className = 'g-row';
    wrap.append(i);
    if (resetLabel) {
      const reset = document.createElement('button');
      reset.type = 'button';
      reset.textContent = resetLabel;
      reset.addEventListener('click', () => {
        set(undefined);
        i.value = fallback();
        refresh();
      });
      wrap.append(reset);
    }
    return field(label, wrap);
  };
  const panelOf = (...children: HTMLElement[]): HTMLElement => {
    const d = document.createElement('div');
    d.className = 'g-chart-panel';
    d.append(...children);
    return d;
  };

  // --- Onglet Type ---
  const typePanel = (): HTMLElement => {
    const wrap = document.createElement('div');
    wrap.className = 'g-kinds';
    const selected = kindOf(work).id;
    for (const g of KIND_GROUPS) {
      const kinds = CHART_KINDS.filter((k) => k.group === g.id);
      const section = document.createElement('div');
      const heading = document.createElement('h3');
      heading.textContent = g.label;
      const buttons = document.createElement('div');
      buttons.className = 'g-kind-row';
      for (const kind of kinds) {
        const b = document.createElement('button');
        b.type = 'button';
        b.className = 'g-kind' + (kind.id === selected ? ' selected' : '');
        b.textContent = kind.label;
        b.dataset.kind = kind.id;
        b.setAttribute('aria-pressed', String(kind.id === selected));
        b.addEventListener('click', () => {
          work = applyKind(work, kind);
          showTab('type');
          refresh();
        });
        buttons.append(b);
      }
      section.append(heading, buttons);
      wrap.append(section);
    }
    return panelOf(wrap);
  };

  // --- Onglet Données ---
  const readGridOf = (rect: { r1: number; c1: number; r2: number; c2: number }, sheetName: string): CellScalar[][] => {
    const ws = workbook(api)?.getSheetByName(sheetName) as unknown as FWorksheet | null;
    if (!ws) return [];
    return ws
      .getRange(rect.r1, rect.c1, rect.r2 - rect.r1 + 1, rect.c2 - rect.c1 + 1)
      .getCellDataGrid()
      .map((row) => row.map((c) => (c?.v === undefined || c?.v === null || c?.v === '' ? null : c?.t === 2 && typeof c.v === 'string' ? Number(c.v) : (c.v as CellScalar))));
  };
  const switchRowsColumns = () => {
    const refs = work.series.flatMap((s) => [s.nameRef, s.valuesRef, s.categoriesRef, s.sizesRef]).filter((r): r is string => !!r);
    const parsed = refs.map((r) => parseSheetRef(r)).filter((p): p is NonNullable<typeof p> => !!p);
    if (!parsed.length) return;
    const sheetName = parsed[0]!.sheet ?? defaultSheet;
    const rect = {
      r1: Math.min(...parsed.map((p) => p.startRow)),
      c1: Math.min(...parsed.map((p) => p.startColumn)),
      r2: Math.max(...parsed.map((p) => p.endRow)),
      c2: Math.max(...parsed.map((p) => p.endColumn)),
    };
    const first = parseSheetRef(work.series[0]!.valuesRef);
    const currentlyByRows = !!first && first.startRow === first.endRow && first.startColumn !== first.endColumn;
    const next = seriesFromRange(sheetName, readGridOf(rect, sheetName), { row: rect.r1, col: rect.c1 }, work.type, { byRows: !currentlyByRows });
    if (!next.length) return;
    work.series = next.map((s, i) => {
      const old = work.series[i];
      return old ? { ...s, color: old.color, type: old.type, secondaryAxis: old.secondaryAxis } : s;
    });
  };
  const dataPanel = (): HTMLElement => {
    const list = document.createElement('div');
    list.className = 'g-series-list';
    const rebuild = () => {
      list.replaceChildren(
        ...work.series.map((s, i) => {
          const card = document.createElement('fieldset');
          const legend = document.createElement('legend');
          legend.textContent = `${L('seriesPick')} ${i + 1}`;
          const refInput = (label: string, get: () => string | undefined, set: (v: string | undefined) => void, required = false) => {
            const input = textInput(get() ?? '');
            input.addEventListener('input', () => {
              const v = input.value.trim();
              if (!v && !required) {
                input.removeAttribute('aria-invalid');
                set(undefined);
              } else {
                const ref = normalizeRef(v, defaultSheet);
                if (!ref) return void input.setAttribute('aria-invalid', 'true');
                input.removeAttribute('aria-invalid');
                set(ref);
              }
              refresh();
            });
            return field(label, input);
          };
          const name = textInput(s.nameRef ? `=${s.nameRef}` : (s.name ?? ''));
          name.addEventListener('input', () => {
            const v = name.value;
            const ref = v.startsWith('=') ? normalizeRef(v, defaultSheet) : null;
            if (v.startsWith('=')) {
              if (!ref) return void name.setAttribute('aria-invalid', 'true');
              name.removeAttribute('aria-invalid');
              s.nameRef = ref;
              delete s.name;
            } else {
              s.name = v || undefined;
              delete s.nameRef;
            }
            refresh();
          });
          const actions = document.createElement('div');
          actions.className = 'g-row';
          const action = (label: string, run: () => void, disabled = false, danger = false) => {
            const b = document.createElement('button');
            b.type = 'button';
            b.textContent = label;
            b.disabled = disabled;
            if (danger) b.className = 'g-danger';
            b.addEventListener('click', () => {
              run();
              rebuild();
              refresh();
            });
            return b;
          };
          actions.append(
            action('↑', () => void ([work.series[i - 1], work.series[i]] = [work.series[i]!, work.series[i - 1]!]), i === 0),
            action('↓', () => void ([work.series[i + 1], work.series[i]] = [work.series[i]!, work.series[i + 1]!]), i === work.series.length - 1),
            action(L('removeSeries'), () => work.series.splice(i, 1), work.series.length <= 1, true),
          );
          card.append(
            legend,
            field(L('seriesName'), name),
            refInput(L('seriesValues'), () => s.valuesRef, (v) => void (s.valuesRef = v ?? s.valuesRef), true),
            ...(isPieLike(work) && i > 0 ? [] : [refInput(L('seriesCategories'), () => s.categoriesRef, (v) => void (v ? (s.categoriesRef = v) : delete s.categoriesRef))]),
            ...(work.type === 'bubble' ? [refInput(L('seriesSizes'), () => s.sizesRef, (v) => void (v ? (s.sizesRef = v) : delete s.sizesRef))] : []),
            actions,
          );
          return card;
        }),
      );
    };
    rebuild();
    const add = document.createElement('button');
    add.type = 'button';
    add.textContent = L('addSeries');
    add.addEventListener('click', () => {
      const last = work.series[work.series.length - 1];
      if (!last) return;
      work.series.push({ ...structuredClone(last), name: undefined, color: undefined, nameRef: shiftRef(last.nameRef, 1), valuesRef: shiftRef(last.valuesRef, 1) ?? last.valuesRef, sizesRef: shiftRef(last.sizesRef, 1) });
      rebuild();
      refresh();
    });
    const swap = document.createElement('button');
    swap.type = 'button';
    swap.textContent = L('switchRowCol');
    swap.addEventListener('click', () => {
      switchRowsColumns();
      rebuild();
      refresh();
    });
    const buttons = document.createElement('div');
    buttons.className = 'g-row';
    buttons.append(add, swap);
    return panelOf(
      list,
      buttons,
      check(L('hiddenCells'), () => work.plotVisibleOnly === false, (v) => void (work.plotVisibleOnly = v ? false : undefined)),
      choose(L('emptyCells'), [['gap', L('emptyGap')], ['zero', L('emptyZero')], ['connect', L('emptyConnect')]], () => work.emptyCells ?? 'gap', (v) => void (work.emptyCells = v === 'gap' ? undefined : (v as ChartSpec['emptyCells']))),
    );
  };

  // --- Étiquettes de données (graphique ou série) ---
  const labelsEditor = (get: () => DataLabelsSpec | undefined, set: (v: DataLabelsSpec | undefined) => void): HTMLElement => {
    const wrap = document.createElement('div');
    wrap.className = 'g-labels';
    const draw = () => {
      const dl = get();
      const rows: HTMLElement[] = [check(L('dataLabels'), () => !!dl?.show, (v) => set(v ? { ...(get() ?? {}), show: true, value: get()?.value ?? true } : undefined), draw)];
      if (dl?.show) {
        const patch = (key: keyof DataLabelsSpec) => (v: unknown) => void set({ ...(get() ?? {}), [key]: v });
        const positions: [string, string][] = [['auto', L('posAuto')], ['center', L('posCenter')], ['insideEnd', L('posInsideEnd')], ['insideBase', L('posInsideBase')], ['outsideEnd', L('posOutsideEnd')], ['above', L('posAbove')], ['below', L('posBelow')], ['left', L('posLeft')], ['right', L('posRight')]];
        rows.push(
          check(L('labelsValue'), () => dl.value, patch('value')),
          check(L('labelsCategory'), () => dl.category, patch('category')),
          check(L('labelsSeries'), () => dl.series, patch('series')),
          check(L('labelsPercent'), () => dl.percent, patch('percent')),
          choose(L('labelsPosition'), positions, () => dl.position ?? 'auto', (v) => patch('position')(v as LabelPosition)),
          text(L('labelsFormat'), () => dl.numFmt, patch('numFmt')),
          choose(L('separator'), [['; ', L('sepSemicolon')], [', ', L('sepComma')], [' ', L('sepSpace')], ['\n', L('sepNewline')]], () => dl.separator ?? '; ', patch('separator')),
        );
      }
      wrap.replaceChildren(...rows);
    };
    draw();
    return wrap;
  };

  // --- Onglet Éléments ---
  const elementsPanel = (): HTMLElement => {
    const hasSecondary = work.series.some((s) => s.secondaryAxis);
    const gridChecks =
      hasAxes(work) && work.type !== 'radar'
        ? group(
            L('gridlines'),
            ...(isXY(work) || !isPieLike(work)
              ? [
                  check(L('gridYMajor'), () => work.axes?.y?.gridMajor ?? true, (v) => void (axis('y').gridMajor = v ?? false)),
                  check(L('gridYMinor'), () => work.axes?.y?.gridMinor, (v) => void (axis('y').gridMinor = v)),
                  check(L('gridXMajor'), () => work.axes?.x?.gridMajor ?? isXY(work), (v) => void (axis('x').gridMajor = v ?? false)),
                  check(L('gridXMinor'), () => work.axes?.x?.gridMinor, (v) => void (axis('x').gridMinor = v)),
                ]
              : []),
          )
        : null;
    return panelOf(
      text(L('chartTitle'), () => work.title, (v) => void (work.title = v)),
      check(L('titleOverlay'), () => work.titleOverlay, (v) => void (work.titleOverlay = v)),
      ...(hasAxes(work)
        ? [
            text(L('axisTitleX'), () => work.xAxisTitle, (v) => void (work.xAxisTitle = v)),
            text(L('axisTitleY'), () => work.yAxisTitle, (v) => void (work.yAxisTitle = v)),
            ...(hasSecondary ? [text(L('axisTitleY2'), () => work.y2AxisTitle, (v) => void (work.y2AxisTitle = v))] : []),
          ]
        : []),
      choose(L('legend'), [['bottom', t('charts.legendBottom')], ['top', t('charts.legendTop')], ['right', t('charts.legendRight')], ['left', 'À gauche'], ['none', t('charts.legendNone')]], () => work.legend ?? 'bottom', (v) => void (work.legend = v as ChartSpec['legend'])),
      check(L('legendOverlay'), () => work.legendOverlay, (v) => void (work.legendOverlay = v)),
      labelsEditor(() => work.dataLabels, (v) => void (work.dataLabels = v)),
      ...(gridChecks ? [gridChecks] : []),
      ...(hasAxes(work) ? [check(L('dataTable'), () => work.dataTable, (v) => void (work.dataTable = v))] : []),
    );
  };

  // --- Onglet Axes ---
  const axis = (key: 'x' | 'y' | 'y2'): AxisSpec => ((work.axes ??= {})[key] ??= {});
  const axisGroup = (key: 'x' | 'y' | 'y2', title: string, valueScale: boolean): HTMLElement => {
    const a = () => axis(key);
    const ticks: [string, string][] = [['out', L('tickOut')], ['in', L('tickIn')], ['cross', L('tickCross')], ['none', L('tickNone')]];
    return group(
      title,
      ...(valueScale
        ? [
            num(L('axMin'), () => work.axes?.[key]?.min, (v) => void (a().min = v), 'Auto'),
            num(L('axMax'), () => work.axes?.[key]?.max, (v) => void (a().max = v), 'Auto'),
            num(L('axUnit'), () => work.axes?.[key]?.majorUnit, (v) => void (a().majorUnit = v), 'Auto'),
            check(L('axLog'), () => work.axes?.[key]?.log, (v) => void (a().log = v)),
            text(L('axFormat'), () => work.axes?.[key]?.numFmt, (v) => void (a().numFmt = v), 'Standard'),
          ]
        : []),
      check(L('axReverse'), () => work.axes?.[key]?.reverse, (v) => void (a().reverse = v)),
      choose(L('axLabels'), [['nextTo', L('axLabelsNext')], ['none', L('axLabelsNone')]], () => work.axes?.[key]?.labels ?? 'nextTo', (v) => void (a().labels = v === 'none' ? 'none' : undefined)),
      num(L('axAngle'), () => work.axes?.[key]?.labelAngle, (v) => void (a().labelAngle = v), '0'),
      choose(L('axTicks'), ticks, () => work.axes?.[key]?.ticks ?? 'out', (v) => void (a().ticks = v === 'out' ? undefined : (v as AxisSpec['ticks']))),
      ...(valueScale ? [check(L('axMinorTicks'), () => work.axes?.[key]?.minorTicks, (v) => void (a().minorTicks = v))] : []),
      check(L('axHidden'), () => work.axes?.[key]?.hidden, (v) => void (a().hidden = v)),
    );
  };
  const axesPanel = (): HTMLElement => {
    if (isPieLike(work)) return panelOf(Object.assign(document.createElement('p'), { textContent: L('noAxes') }));
    if (work.type === 'radar') return panelOf(axisGroup('y', L('axisY'), true));
    const horizontal = work.type === 'bar';
    const hasSecondary = work.series.some((s) => s.secondaryAxis);
    return panelOf(
      axisGroup('x', horizontal ? L('axisY') : L('axisX'), isXY(work)),
      axisGroup('y', horizontal ? L('axisX') : L('axisY'), true),
      ...(hasSecondary ? [axisGroup('y2', L('axisY2'), true)] : []),
    );
  };

  // --- Onglet Séries ---
  let seriesIndex = 0;
  const trendOf = (s: SeriesSpec): TrendlineSpec | undefined => s.trendlines?.[0];
  const seriesPanel = (): HTMLElement => {
    seriesIndex = Math.min(seriesIndex, work.series.length - 1);
    const s = work.series[seriesIndex]!;
    const picker = select(work.series.map((x, i) => [String(i), x.name ?? (x.nameRef ? String(readRef(api, x.nameRef, { display: true, defaultSheet })[0] ?? `=${x.nameRef}`) : `${L('seriesPick')} ${i + 1}`)] as [string, string]), String(seriesIndex));
    picker.addEventListener('change', () => {
      seriesIndex = Number(picker.value);
      showTab('series');
    });
    const palette = PALETTES[work.palette ?? 'office']?.colors ?? DEFAULT_COLORS;
    const cartesian = hasAxes(work);
    const kind = s.type ?? work.type;
    const lineLike = kind === 'line' || kind === 'area' || kind === 'radar' || (kind === 'scatter' && (work.scatterStyle ?? 'markers') !== 'markers');
    const markerLike = kind === 'line' || kind === 'scatter' || kind === 'radar';
    const body: HTMLElement[] = [field(L('seriesPick'), picker), color(L('color'), () => s.color, () => palette[seriesIndex % palette.length]!, (v) => void (s.color = v), L('backgroundAuto'))];
    if (cartesian && !isXY(work) && work.type !== 'radar') {
      const types: [string, string][] = [['', t('charts.sameType')], ['column', t('charts.column')], ['line', t('charts.line')], ['area', t('charts.area')]];
      body.push(
        choose(L('seriesType'), types, () => s.type ?? '', (v) => void (s.type = (v || undefined) as SeriesSpec['type']), () => showTab('series')),
        check(L('secondary'), () => s.secondaryAxis, (v) => void (s.secondaryAxis = v), () => showTab('series')),
      );
    }
    if (lineLike) {
      body.push(
        num(L('lineWidth'), () => s.lineWidth, (v) => void (s.lineWidth = v), '2,25'),
        choose(L('dash'), [['solid', L('dashSolid')], ['dash', L('dashDash')], ['dot', L('dashDot')]], () => s.dash ?? 'solid', (v) => void (s.dash = v === 'solid' ? undefined : (v as SeriesSpec['dash']))),
      );
    }
    if (markerLike) {
      const shapes: [string, string][] = [['', L('markerAuto')], ['none', L('markerNone')], ['circle', L('markerCircle')], ['square', L('markerSquare')], ['diamond', L('markerDiamond')], ['triangle', L('markerTriangle')]];
      body.push(
        choose(L('marker'), shapes, () => s.marker?.shape ?? '', (v) => void (s.marker = { ...(s.marker ?? {}), shape: (v || undefined) as MarkerShape | undefined })),
        num(L('markerSize'), () => s.marker?.size, (v) => void (s.marker = { ...(s.marker ?? {}), size: v }), '6'),
      );
    }
    if (kind === 'line' || kind === 'area') body.push(check(L('smooth'), () => s.smooth, (v) => void (s.smooth = v)));
    if (kind === 'area' || (kind === 'radar' && work.radarFilled)) body.push(num(L('opacity'), () => (s.opacity === undefined ? undefined : Math.round(s.opacity * 100)), (v) => void (s.opacity = v === undefined ? undefined : v / 100), '50'));
    body.push(
      labelsEditor(() => s.dataLabels, (v) => void (s.dataLabels = v)),
    );

    if (cartesian && work.type !== 'bar') {
      const t0 = trendOf(s);
      const trendBox = document.createElement('div');
      const types: [string, string][] = [['', L('trendNone')], ['linear', L('trendLinear')], ['exponential', L('trendExp')], ['logarithmic', L('trendLog')], ['polynomial', L('trendPoly')], ['power', L('trendPower')], ['movingAverage', L('trendMA')]];
      const setTrend = (patch: Partial<TrendlineSpec> | null) => {
        if (!patch) delete s.trendlines;
        else s.trendlines = [{ ...(trendOf(s) ?? { type: 'linear' }), ...patch }, ...(s.trendlines?.slice(1) ?? [])];
      };
      const drawTrend = () => {
        const cur = trendOf(s);
        const rows: HTMLElement[] = [choose(L('trendline'), types, () => cur?.type ?? '', (v) => setTrend(v ? { type: v as TrendlineSpec['type'] } : null), drawTrend)];
        if (cur) {
          if (cur.type === 'polynomial') rows.push(num(L('trendOrder'), () => cur.order, (v) => setTrend({ order: v }), '2'));
          if (cur.type === 'movingAverage') rows.push(num(L('trendPeriod'), () => cur.period, (v) => setTrend({ period: v }), '2'));
          if (cur.type !== 'movingAverage') {
            rows.push(
              num(L('trendForward'), () => cur.forward, (v) => setTrend({ forward: v }), '0'),
              num(L('trendBackward'), () => cur.backward, (v) => setTrend({ backward: v }), '0'),
              check(L('trendEquation'), () => cur.showEquation, (v) => setTrend({ showEquation: v })),
              check(L('trendR2'), () => cur.showR2, (v) => setTrend({ showR2: v })),
            );
          }
        }
        trendBox.replaceChildren(...rows);
      };
      drawTrend();
      body.push(group(L('trendline'), trendBox));
      const errBox = document.createElement('div');
      const kinds: [string, string][] = [['', L('errNone')], ['fixed', L('errFixed')], ['percent', L('errPercent')], ['stddev', L('errStddev')]];
      const drawErr = () => {
        const cur = s.errorBars;
        const rows: HTMLElement[] = [choose(L('errorBars'), kinds, () => cur?.kind ?? '', (v) => void (v ? (s.errorBars = { ...(cur ?? {}), kind: v as 'fixed' | 'percent' | 'stddev' }) : delete s.errorBars), drawErr)];
        if (cur) {
          rows.push(
            num(L('errValue'), () => cur.value, (v) => void (cur.value = v), cur.kind === 'percent' ? '5' : '1'),
            choose(L('errSides'), [['both', L('errBoth')], ['plus', L('errPlus')], ['minus', L('errMinus')]], () => cur.sides ?? 'both', (v) => void (cur.sides = v === 'both' ? undefined : (v as 'plus' | 'minus'))),
            check(L('errNoCap'), () => cur.noCap, (v) => void (cur.noCap = v)),
          );
        }
        errBox.replaceChildren(...rows);
      };
      drawErr();
      body.push(group(L('errorBars'), errBox));
    }
    return panelOf(...body);
  };

  // --- Onglet Mise en forme ---
  const formatPanel = (): HTMLElement => {
    const rows: HTMLElement[] = [
      choose(L('palette'), Object.entries(PALETTES).map(([id, p]) => [id, p.label] as [string, string]), () => work.palette ?? 'office', (v) => void (work.palette = v === 'office' ? undefined : v)),
      color(L('background'), () => work.background, () => (isDark() ? '#1b1f23' : '#ffffff'), (v) => void (work.background = v), L('backgroundAuto')),
      check(L('border'), () => work.border !== false, (v) => void (work.border = v ? undefined : false)),
      num(L('fontSize'), () => work.fontSize, (v) => void (work.fontSize = v), '12'),
    ];
    if (work.type === 'column' || work.type === 'bar') {
      rows.push(num(L('gapWidth'), () => work.gapWidth, (v) => void (work.gapWidth = v), '150'), num(L('overlap'), () => work.overlap, (v) => void (work.overlap = v), '0'));
    }
    if (work.type === 'line') rows.push(check(L('markers'), () => work.markers, (v) => void (work.markers = v)), check(L('smoothAll'), () => work.smooth, (v) => void (work.smooth = v)));
    if (work.type === 'doughnut') rows.push(num(L('holeSize'), () => work.holeSize, (v) => void (work.holeSize = v), '50'));
    if (isPieLike(work)) rows.push(num(L('firstAngle'), () => work.firstSliceAngle, (v) => void (work.firstSliceAngle = v), '0'), num(L('explosion'), () => work.explosion, (v) => void (work.explosion = v), '0'));
    if (work.type === 'bubble') rows.push(num(L('bubbleScale'), () => work.bubbleScale, (v) => void (work.bubbleScale = v), '100'));
    if (work.type === 'scatter') {
      rows.push(choose('Style', [['markers', 'Marqueurs seuls'], ['straightMarkers', 'Lignes droites et marqueurs'], ['straight', 'Lignes droites'], ['smoothMarkers', 'Courbes lissées et marqueurs'], ['smooth', 'Courbes lissées']], () => work.scatterStyle ?? 'markers', (v) => void (work.scatterStyle = v === 'markers' ? undefined : (v as ChartSpec['scatterStyle']))));
    }
    if (work.type === 'waterfall') {
      // « Définir comme total » d'Excel : numéros des points (1 = premier).
      rows.push(
        text(L('waterfallTotals'), () => (work.waterfallTotals ?? []).map((i) => i + 1).join('; ') || undefined, (v) => {
          const idx = (v ?? '').split(/[;,\s]+/).map(Number).filter((n) => Number.isInteger(n) && n > 0).map((n) => n - 1);
          work.waterfallTotals = idx.length ? idx : undefined;
        }, '1; 6'),
      );
    }
    if (work.type === 'histogram') {
      const bins = () => (work.bins ??= {});
      rows.push(
        num(L('binWidth'), () => work.bins?.width, (v) => void (bins().width = v), L('automatic')),
        num(L('binCount'), () => work.bins?.count, (v) => void (bins().count = v), L('automatic')),
        num(L('binOverflow'), () => work.bins?.overflow, (v) => void (bins().overflow = v)),
        num(L('binUnderflow'), () => work.bins?.underflow, (v) => void (bins().underflow = v)),
      );
    }
    if (work.type === 'boxWhisker') {
      rows.push(
        check(L('boxMean'), () => work.boxShowMean !== false, (v) => void (work.boxShowMean = v ? undefined : false)),
        check(L('boxOutliers'), () => work.boxShowOutliers !== false, (v) => void (work.boxShowOutliers = v ? undefined : false)),
      );
    }
    if (work.series.length === 1 && !isXY(work) && work.type !== 'radar' && !isStatChart(work.type)) rows.push(check(L('varyColors'), () => work.varyColors, (v) => void (work.varyColors = v)));
    return panelOf(...rows);
  };

  // --- Onglets ---
  const tabs: { id: string; label: string; build: () => HTMLElement }[] = [
    { id: 'type', label: L('tabType'), build: typePanel },
    { id: 'data', label: L('tabData'), build: dataPanel },
    { id: 'elements', label: L('tabElements'), build: elementsPanel },
    { id: 'axes', label: L('tabAxes'), build: axesPanel },
    { id: 'series', label: L('tabSeries'), build: seriesPanel },
    { id: 'format', label: L('tabFormat'), build: formatPanel },
  ];
  const tabBar = document.createElement('div');
  tabBar.className = 'g-chart-tabs';
  tabBar.setAttribute('role', 'tablist');
  const panelHost = document.createElement('div');
  let currentTab = inserting ? 'type' : 'elements';
  function showTab(id: string): void {
    currentTab = id;
    for (const b of tabBar.querySelectorAll('button')) {
      const on = b.dataset.tab === id;
      b.setAttribute('aria-selected', String(on));
      b.classList.toggle('selected', on);
    }
    panelHost.replaceChildren(tabs.find((x) => x.id === id)!.build());
  }
  for (const tab of tabs) {
    const b = document.createElement('button');
    b.type = 'button';
    b.setAttribute('role', 'tab');
    b.dataset.tab = tab.id;
    b.textContent = tab.label;
    b.addEventListener('click', () => showTab(tab.id));
    tabBar.append(b);
  }
  showTab(currentTab);

  const left = document.createElement('div');
  left.className = 'g-chart-left';
  left.append(tabBar, panelHost);
  const layout = document.createElement('div');
  layout.className = 'g-chart-editor';
  layout.append(left, preview);

  const extra = document.createElement('div');
  if (existing) {
    const del = document.createElement('button');
    del.type = 'button';
    del.className = 'g-danger';
    del.textContent = t('charts.delete');
    let removed = false;
    del.addEventListener('click', () => {
      removed = true;
      removeChart(api, opts.chartId!);
      (del.closest('.g-dialog')?.querySelector('.g-dialog-actions button:not(.primary)') as HTMLButtonElement | null)?.click();
    });
    extra.append(del);
    Object.defineProperty(extra, 'removed', { get: () => removed });
  }

  const done = openDialog({
    title: inserting ? L('insertTitle') : L('editTitle'),
    testId: 'dialog-chart',
    xl: true,
    okLabel: inserting ? L('insertOk') : L('applyOk'),
    body: box(layout, extra),
    onOk: () => true,
  });
  const ok = await done;
  clearTimeout(timer);
  (chart as { dispose(): void } | null)?.dispose();
  if (!ok || (extra as unknown as { removed?: boolean }).removed) return;
  const final = pruneSpec(work);
  if (existing) updateChart(api, opts.chartId!, final);
  else insertChartSpec(api, final);
}

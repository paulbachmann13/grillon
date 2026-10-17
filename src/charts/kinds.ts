/**
 * Catalogue des types et sous-types de graphiques (la boîte « Insérer un graphique » d'Excel).
 * Chaque sous-type est un ensemble d'options appliqué à la spécification.
 */
import { groupingOf, type ChartSpec, type ChartType, type Grouping, type ScatterStyle } from './spec';

export interface ChartKind {
  id: string;
  group: string;
  label: string;
  type: ChartType;
  grouping?: Grouping;
  markers?: boolean;
  scatterStyle?: ScatterStyle;
  radarFilled?: boolean;
  /** Histogramme groupé + courbe sur axe secondaire (la dernière série devient une courbe). */
  combo?: 'line' | 'line-secondary';
}

export const KIND_GROUPS: { id: string; label: string }[] = [
  { id: 'column', label: 'Histogramme' },
  { id: 'line', label: 'Courbes' },
  { id: 'pie', label: 'Secteurs' },
  { id: 'bar', label: 'Barres' },
  { id: 'area', label: 'Aires' },
  { id: 'scatter', label: 'Nuage de points' },
  { id: 'bubble', label: 'Bulles' },
  { id: 'radar', label: 'Radar' },
  { id: 'combo', label: 'Combo' },
  { id: 'waterfall', label: 'Cascade' },
  { id: 'funnel', label: 'Entonnoir' },
  { id: 'statistic', label: 'Statistique' },
];

export const CHART_KINDS: ChartKind[] = [
  { id: 'column', group: 'column', label: 'Histogramme groupé', type: 'column' },
  { id: 'column-stacked', group: 'column', label: 'Histogramme empilé', type: 'column', grouping: 'stacked' },
  { id: 'column-percent', group: 'column', label: 'Histogramme empilé 100 %', type: 'column', grouping: 'percent' },
  { id: 'line', group: 'line', label: 'Courbes', type: 'line' },
  { id: 'line-stacked', group: 'line', label: 'Courbes empilées', type: 'line', grouping: 'stacked' },
  { id: 'line-percent', group: 'line', label: 'Courbes empilées 100 %', type: 'line', grouping: 'percent' },
  { id: 'line-markers', group: 'line', label: 'Courbes avec marques', type: 'line', markers: true },
  { id: 'line-stacked-markers', group: 'line', label: 'Courbes empilées avec marques', type: 'line', grouping: 'stacked', markers: true },
  { id: 'pie', group: 'pie', label: 'Secteurs', type: 'pie' },
  { id: 'doughnut', group: 'pie', label: 'Anneau', type: 'doughnut' },
  { id: 'bar', group: 'bar', label: 'Barres groupées', type: 'bar' },
  { id: 'bar-stacked', group: 'bar', label: 'Barres empilées', type: 'bar', grouping: 'stacked' },
  { id: 'bar-percent', group: 'bar', label: 'Barres empilées 100 %', type: 'bar', grouping: 'percent' },
  { id: 'area', group: 'area', label: 'Aires', type: 'area' },
  { id: 'area-stacked', group: 'area', label: 'Aires empilées', type: 'area', grouping: 'stacked' },
  { id: 'area-percent', group: 'area', label: 'Aires empilées 100 %', type: 'area', grouping: 'percent' },
  { id: 'scatter', group: 'scatter', label: 'Nuage de points', type: 'scatter', scatterStyle: 'markers' },
  { id: 'scatter-smooth-markers', group: 'scatter', label: 'Nuage avec courbes lissées et marqueurs', type: 'scatter', scatterStyle: 'smoothMarkers' },
  { id: 'scatter-smooth', group: 'scatter', label: 'Nuage avec courbes lissées', type: 'scatter', scatterStyle: 'smooth' },
  { id: 'scatter-lines-markers', group: 'scatter', label: 'Nuage avec lignes droites et marqueurs', type: 'scatter', scatterStyle: 'straightMarkers' },
  { id: 'scatter-lines', group: 'scatter', label: 'Nuage avec lignes droites', type: 'scatter', scatterStyle: 'straight' },
  { id: 'bubble', group: 'bubble', label: 'Bulles', type: 'bubble' },
  { id: 'radar', group: 'radar', label: 'Radar', type: 'radar' },
  { id: 'radar-markers', group: 'radar', label: 'Radar avec marqueurs', type: 'radar', markers: true },
  { id: 'radar-filled', group: 'radar', label: 'Radar plein', type: 'radar', radarFilled: true },
  { id: 'combo', group: 'combo', label: 'Histogramme groupé – Courbe', type: 'column', combo: 'line' },
  { id: 'combo-secondary', group: 'combo', label: 'Histogramme groupé – Courbe sur axe secondaire', type: 'column', combo: 'line-secondary' },
  { id: 'waterfall', group: 'waterfall', label: 'Cascade', type: 'waterfall' },
  { id: 'funnel', group: 'funnel', label: 'Entonnoir', type: 'funnel' },
  { id: 'histogram', group: 'statistic', label: 'Histogramme (répartition)', type: 'histogram' },
  { id: 'box-whisker', group: 'statistic', label: 'Zone et valeur (boîte à moustaches)', type: 'boxWhisker' },
];

export const kindById = (id: string): ChartKind => CHART_KINDS.find((k) => k.id === id) ?? CHART_KINDS[0]!;

/** Applique un sous-type à une spécification (les séries sont conservées). */
export function applyKind(spec: ChartSpec, kind: ChartKind): ChartSpec {
  const next: ChartSpec = { ...spec, type: kind.type };
  delete next.stacked;
  if (kind.grouping) next.grouping = kind.grouping;
  else delete next.grouping;
  if (kind.markers !== undefined) next.markers = kind.markers;
  else delete next.markers;
  if (kind.scatterStyle) next.scatterStyle = kind.scatterStyle;
  else delete next.scatterStyle;
  if (kind.radarFilled) next.radarFilled = true;
  else delete next.radarFilled;
  next.series = spec.series.map((s) => {
    const copy = { ...s };
    delete copy.type;
    delete copy.secondaryAxis;
    return copy;
  });
  if (kind.combo && next.series.length > 1) {
    const last = next.series[next.series.length - 1]!;
    last.type = 'line';
    if (kind.combo === 'line-secondary') last.secondaryAxis = true;
  }
  return next;
}

/** Sous-type correspondant à une spécification (pour mettre en évidence le choix actuel). */
export function kindOf(spec: ChartSpec): ChartKind {
  const grouping = groupingOf(spec);
  const sameGrouping = (k: ChartKind) => (k.grouping ?? 'standard') === grouping;
  const candidates = CHART_KINDS.filter((k) => k.type === spec.type && !k.combo);
  const last = spec.series[spec.series.length - 1];
  if (spec.type === 'column' && spec.series.length > 1 && last?.type === 'line') return kindById(last.secondaryAxis ? 'combo-secondary' : 'combo');
  const exact = candidates.find(
    (k) => sameGrouping(k) && (k.markers ?? false) === (spec.markers ?? false) && (k.scatterStyle ?? 'markers') === (spec.scatterStyle ?? 'markers') && (k.radarFilled ?? false) === (spec.radarFilled ?? false),
  );
  return exact ?? candidates.find(sameGrouping) ?? candidates[0] ?? CHART_KINDS[0]!;
}

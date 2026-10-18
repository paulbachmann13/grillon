import { describe, expect, it } from 'vitest';
import {
  cellMatches,
  clickItem,
  clickPeriod,
  filterOf,
  isOpen,
  periodKey,
  periodLabel,
  pivotInclude,
  serialToYmd,
  slicerItems,
  timelinePeriods,
  visibleRows,
  ymdToSerial,
  type ColumnCell,
  type SlicerFilter,
  type SlicerSpec,
} from '../../src/data/slicer';

const text = (label: string): ColumnCell => ({ label, value: label });
const date = (y: number, m: number, d: number): ColumnCell => ({ label: `${d}/${m}/${y}`, value: ymdToSerial(y, m, d) });

// Colonnes : 0 = région, 1 = date
const rows: ColumnCell[][] = [
  [text('Nord'), date(2024, 1, 15)],
  [text('Sud'), date(2024, 2, 3)],
  [text('Nord'), date(2024, 4, 20)],
  [text('Est'), date(2025, 1, 2)],
  [{ label: '', value: null }, { label: '', value: null }],
];

describe('dates', () => {
  it('numéros de série Excel', () => {
    expect(ymdToSerial(2024, 1, 1)).toBe(45292);
    expect(serialToYmd(45292)).toEqual({ y: 2024, m: 1, d: 1 });
    expect(serialToYmd(45291.75)).toEqual({ y: 2023, m: 12, d: 31 });
    expect(serialToYmd(ymdToSerial(2024, 2, 29))).toEqual({ y: 2024, m: 2, d: 29 });
  });
  it('clés et libellés de période', () => {
    const s = ymdToSerial(2024, 5, 9);
    expect(periodKey(s, 'years')).toBe('2024');
    expect(periodKey(s, 'quarters')).toBe('2024-T2');
    expect(periodKey(s, 'months')).toBe('2024-05');
    expect(periodKey(s, 'days')).toBe('2024-05-09');
    expect(['2024', '2024-T2', '2024-05', '2024-05-09'].map(periodLabel)).toEqual(['2024', 'T2 2024', 'mai 2024', '9 mai 2024']);
  });
  it('toutes les périodes entre la première et la dernière date, même sans donnée', () => {
    const dates = rows.map((r) => r[1].value);
    expect(timelinePeriods(dates, 'years')).toEqual(['2024', '2025']);
    expect(timelinePeriods(dates, 'quarters')).toEqual(['2024-T1', '2024-T2', '2024-T3', '2024-T4', '2025-T1']);
    expect(timelinePeriods(dates, 'months')).toHaveLength(13);
    expect(timelinePeriods(dates, 'months').slice(-2)).toEqual(['2024-12', '2025-01']);
    expect(timelinePeriods(['a', null], 'years')).toEqual([]);
    expect(timelinePeriods([ymdToSerial(2024, 12, 30), ymdToSerial(2025, 1, 2)], 'days')).toEqual(['2024-12-30', '2024-12-31', '2025-01-01', '2025-01-02']);
  });
});

describe('segments', () => {
  const region: SlicerFilter = { kind: 'slicer', col: 0, selected: ['Nord'] };
  it('retient les lignes des éléments choisis', () => {
    expect(visibleRows(rows, [region])).toEqual([true, false, true, false, false]);
    expect(visibleRows(rows, [{ kind: 'slicer', col: 0, selected: null }])).toEqual([true, true, true, true, true]);
    expect(isOpen({ kind: 'slicer', col: 0, selected: null })).toBe(true);
  });
  it('les cellules vides forment l’élément « (vide) »', () => {
    expect(cellMatches(rows[4][0], { kind: 'slicer', col: 0, selected: ['(vide)'] })).toBe(true);
    expect(cellMatches(undefined, { kind: 'slicer', col: 0, selected: ['Nord'] })).toBe(false);
  });
  it('éléments triés, vide en dernier, comptés selon les autres filtres', () => {
    const period: SlicerFilter = { kind: 'timeline', col: 1, level: 'years', from: '2024', to: '2024' };
    const items = slicerItems(rows, [{ kind: 'slicer', col: 0, selected: null }, period], 0);
    expect(items).toEqual([
      { label: 'Est', count: 0 },
      { label: 'Nord', count: 2 },
      { label: 'Sud', count: 1 },
      { label: '(vide)', count: 0 },
    ]);
  });
  it('clic : un seul élément, Ctrl pour ajouter ou retirer, retour à « tout »', () => {
    const all = ['Est', 'Nord', 'Sud'];
    expect(clickItem(null, 'Nord', all, false)).toEqual(['Nord']);
    expect(clickItem(['Nord'], 'Nord', all, false)).toBeNull();
    expect(clickItem(['Nord'], 'Sud', all, true)).toEqual(['Nord', 'Sud']);
    expect(clickItem(['Nord', 'Sud'], 'Nord', all, true)).toEqual(['Sud']);
    expect(clickItem(['Nord'], 'Nord', all, true)).toBeNull();
    expect(clickItem(['Nord', 'Sud'], 'Est', all, true)).toBeNull();
    expect(clickItem(null, 'Nord', all, true)).toEqual(['Est', 'Sud']);
  });
  it('filtre construit depuis la définition', () => {
    const spec: SlicerSpec = { id: 's', type: 'slicer', title: 'Région', target: { kind: 'table', sheet: 'F', ref: 'A1:B5', col: 0 }, selected: ['Sud'] };
    expect(filterOf(spec)).toEqual({ kind: 'slicer', col: 0, selected: ['Sud'] });
  });
});

describe('chronologies', () => {
  it('retient les dates dans l’intervalle de périodes', () => {
    const f: SlicerFilter = { kind: 'timeline', col: 1, level: 'quarters', from: '2024-T1', to: '2024-T1' };
    expect(visibleRows(rows, [f])).toEqual([true, true, false, false, false]);
    const open: SlicerFilter = { kind: 'timeline', col: 1, level: 'months', from: null, to: null };
    expect(visibleRows(rows, [open]).every(Boolean)).toBe(true);
  });
  it('plusieurs filtres se combinent (ET)', () => {
    const f1: SlicerFilter = { kind: 'slicer', col: 0, selected: ['Nord'] };
    const f2: SlicerFilter = { kind: 'timeline', col: 1, level: 'months', from: '2024-03', to: '2025-01' };
    expect(visibleRows(rows, [f1, f2])).toEqual([false, false, true, false, false]);
  });
  it('clic sur une période, Maj pour étendre, reclic pour effacer', () => {
    expect(clickPeriod({ from: null, to: null }, '2024-03', false)).toEqual({ from: '2024-03', to: '2024-03' });
    expect(clickPeriod({ from: '2024-03', to: '2024-03' }, '2024-06', true)).toEqual({ from: '2024-03', to: '2024-06' });
    expect(clickPeriod({ from: '2024-03', to: '2024-03' }, '2024-01', true)).toEqual({ from: '2024-01', to: '2024-03' });
    expect(clickPeriod({ from: '2024-03', to: '2024-03' }, '2024-03', false)).toEqual({ from: null, to: null });
  });
});

describe('tableau croisé dynamique', () => {
  const column = rows.map((r) => r[0]);
  it('libellés à garder pour un segment', () => {
    expect(pivotInclude(column, { kind: 'slicer', col: 0, selected: ['Nord', 'Est'] })).toEqual(['Nord', 'Est']);
    expect(pivotInclude(column, { kind: 'slicer', col: 0, selected: null })).toBeNull();
  });
  it('libellés à garder pour une chronologie', () => {
    const dates = rows.map((r) => r[1]);
    expect(pivotInclude(dates, { kind: 'timeline', col: 1, level: 'years', from: '2025', to: '2025' })).toEqual(['2/1/2025']);
  });
});

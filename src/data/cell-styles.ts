/**
 * Styles de cellule d'Excel (Accueil → Styles de cellule) et calcul des styles d'une plage
 * (bordures de contour / intérieures selon la position de chaque cellule).
 */
import type { IStyleData } from '../model/univer-types';

type Border = { s: number; cl: { rgb: string } };
const line = (s: number, rgb: string): Border => ({ s, cl: { rgb } });
const THIN = 1;
const MEDIUM = 8;
const THICK = 13;
const DOUBLE = 7;
/** Bordure retirée (style « aucun »). */
export const NO_BORDER: Border = { s: 0, cl: { rgb: '#000000' } };

export interface CellStyleEntry {
  id: string;
  label: string;
  group: string;
  /** null : style « Normal » (efface la mise en forme). */
  style: IStyleData | null;
}

const ACCENTS: [string, string][] = [
  ['#4472C4', 'Accent 1'],
  ['#ED7D31', 'Accent 2'],
  ['#A5A5A5', 'Accent 3'],
  ['#FFC000', 'Accent 4'],
  ['#5B9BD5', 'Accent 5'],
  ['#70AD47', 'Accent 6'],
];

/** Couleur éclaircie (teinte) : t = 0 → couleur, t = 1 → blanc. */
export function tint(hex: string, t: number): string {
  const n = Number.parseInt(hex.slice(1), 16);
  const ch = (shift: number) => Math.round(((n >> shift) & 255) + (255 - ((n >> shift) & 255)) * t);
  return `#${[16, 8, 0].map((s) => ch(s).toString(16).padStart(2, '0')).join('').toUpperCase()}`;
}

const TITLE = '#44546A';

export const CELL_STYLES: CellStyleEntry[] = [
  { id: 'normal', label: 'Normal', group: 'Bon, mauvais et neutre', style: null },
  { id: 'bad', label: 'Insatisfaisant', group: 'Bon, mauvais et neutre', style: { bg: { rgb: '#FFC7CE' }, cl: { rgb: '#9C0006' } } },
  { id: 'good', label: 'Satisfaisant', group: 'Bon, mauvais et neutre', style: { bg: { rgb: '#C6EFCE' }, cl: { rgb: '#006100' } } },
  { id: 'neutral', label: 'Neutre', group: 'Bon, mauvais et neutre', style: { bg: { rgb: '#FFEB9C' }, cl: { rgb: '#9C5700' } } },
  {
    id: 'calculation',
    label: 'Calcul',
    group: 'Données et modèle',
    style: { bg: { rgb: '#F2F2F2' }, cl: { rgb: '#FA7D00' }, bl: 1, bd: { t: line(THIN, '#7F7F7F'), b: line(THIN, '#7F7F7F'), l: line(THIN, '#7F7F7F'), r: line(THIN, '#7F7F7F') } },
  },
  {
    id: 'check',
    label: 'Cellule de vérification',
    group: 'Données et modèle',
    style: { bg: { rgb: '#A5A5A5' }, cl: { rgb: '#FFFFFF' }, bl: 1, bd: { t: line(DOUBLE, '#3F3F3F'), b: line(DOUBLE, '#3F3F3F'), l: line(DOUBLE, '#3F3F3F'), r: line(DOUBLE, '#3F3F3F') } },
  },
  { id: 'explanatory', label: 'Texte explicatif', group: 'Données et modèle', style: { it: 1, cl: { rgb: '#7F7F7F' } } },
  {
    id: 'input',
    label: 'Entrée',
    group: 'Données et modèle',
    style: { bg: { rgb: '#FFCC99' }, cl: { rgb: '#3F3F76' }, bd: { t: line(THIN, '#7F7F7F'), b: line(THIN, '#7F7F7F'), l: line(THIN, '#7F7F7F'), r: line(THIN, '#7F7F7F') } },
  },
  { id: 'linked', label: 'Cellule liée', group: 'Données et modèle', style: { cl: { rgb: '#FA7D00' }, bd: { b: line(DOUBLE, '#FF8001') } } },
  {
    id: 'note',
    label: 'Remarque',
    group: 'Données et modèle',
    style: { bg: { rgb: '#FFFFCC' }, bd: { t: line(THIN, '#B2B2B2'), b: line(THIN, '#B2B2B2'), l: line(THIN, '#B2B2B2'), r: line(THIN, '#B2B2B2') } },
  },
  {
    id: 'output',
    label: 'Sortie',
    group: 'Données et modèle',
    style: { bg: { rgb: '#F2F2F2' }, cl: { rgb: '#3F3F3F' }, bl: 1, bd: { t: line(THIN, '#3F3F3F'), b: line(THIN, '#3F3F3F'), l: line(THIN, '#3F3F3F'), r: line(THIN, '#3F3F3F') } },
  },
  { id: 'warning', label: 'Avertissement', group: 'Données et modèle', style: { cl: { rgb: '#FF0000' } } },
  { id: 'title', label: 'Titre', group: 'Titres et en-têtes', style: { fs: 18, cl: { rgb: TITLE } } },
  { id: 'heading1', label: 'Titre 1', group: 'Titres et en-têtes', style: { fs: 15, bl: 1, cl: { rgb: TITLE }, bd: { b: line(THICK, '#4472C4') } } },
  { id: 'heading2', label: 'Titre 2', group: 'Titres et en-têtes', style: { fs: 13, bl: 1, cl: { rgb: TITLE }, bd: { b: line(THICK, '#A2B8E1') } } },
  { id: 'heading3', label: 'Titre 3', group: 'Titres et en-têtes', style: { fs: 11, bl: 1, cl: { rgb: TITLE }, bd: { b: line(MEDIUM, '#8EA9DB') } } },
  { id: 'heading4', label: 'Titre 4', group: 'Titres et en-têtes', style: { bl: 1, cl: { rgb: TITLE } } },
  { id: 'total', label: 'Total', group: 'Titres et en-têtes', style: { bl: 1, bd: { t: line(THIN, '#4472C4'), b: line(DOUBLE, '#4472C4') } } },
  ...ACCENTS.flatMap(([hex, name], i): CellStyleEntry[] => [
    { id: `accent${i + 1}-20`, label: `20 % – ${name}`, group: 'Styles de cellule à thème', style: { bg: { rgb: tint(hex, 0.8) } } },
    { id: `accent${i + 1}-40`, label: `40 % – ${name}`, group: 'Styles de cellule à thème', style: { bg: { rgb: tint(hex, 0.6) } } },
    { id: `accent${i + 1}-60`, label: `60 % – ${name}`, group: 'Styles de cellule à thème', style: { bg: { rgb: tint(hex, 0.4) }, cl: { rgb: '#FFFFFF' } } },
    { id: `accent${i + 1}`, label: name, group: 'Styles de cellule à thème', style: { bg: { rgb: hex }, cl: { rgb: '#FFFFFF' } } },
  ]),
  { id: 'comma', label: 'Milliers', group: 'Format de nombre', style: { n: { pattern: '#,##0.00' } } },
  { id: 'comma0', label: 'Milliers [0]', group: 'Format de nombre', style: { n: { pattern: '#,##0' } } },
  { id: 'currency', label: 'Monétaire', group: 'Format de nombre', style: { n: { pattern: '#,##0.00 "€"' } } },
  { id: 'currency0', label: 'Monétaire [0]', group: 'Format de nombre', style: { n: { pattern: '#,##0 "€"' } } },
  { id: 'percent', label: 'Pourcentage', group: 'Format de nombre', style: { n: { pattern: '0%' } } },
];

export type BorderPart = 't' | 'b' | 'l' | 'r' | 'h' | 'v';

/**
 * Style de chaque cellule d'une plage : `style` partout, plus les bordures demandées selon la
 * position (t/b/l/r = contour, h/v = séparations intérieures horizontales / verticales).
 * `null` pour une bordure = la retirer.
 */
export function rangeStyles(rows: number, cols: number, style: IStyleData, borders?: Partial<Record<BorderPart, Border | null>>): IStyleData[][] {
  const out: IStyleData[][] = [];
  for (let r = 0; r < rows; r++) {
    const row: IStyleData[] = [];
    for (let c = 0; c < cols; c++) {
      const s: IStyleData = { ...style };
      if (borders) {
        const bd: Record<string, Border> = { ...((style.bd as Record<string, Border> | undefined) ?? {}) };
        const pick = (edge: 't' | 'b' | 'l' | 'r', outer: boolean, inner: BorderPart) => {
          const key = outer ? edge : inner;
          if (key in borders) bd[edge] = borders[key] ?? NO_BORDER;
        };
        pick('t', r === 0, 'h');
        pick('b', r === rows - 1, 'h');
        pick('l', c === 0, 'v');
        pick('r', c === cols - 1, 'v');
        s.bd = bd as IStyleData['bd'];
      }
      row.push(s);
    }
    out.push(row);
  }
  return out;
}

/**
 * Styles de tableau d'Excel (galerie « Mettre sous forme de tableau ») : Clair 1 à 7, Moyen 1 à 7
 * et Foncé 1 à 7, aux couleurs du thème Office. Le nom est celui d'Excel (TableStyleMedium2…),
 * écrit tel quel dans le .xlsx.
 */
import { tint } from './cell-styles';

type Color = { rgb: string };
type Line = { s: number; cl: Color };
export interface TableStyleItem {
  bg?: Color;
  cl?: Color;
  bl?: number;
  bd?: Partial<Record<'t' | 'b' | 'l' | 'r', Line>>;
}
export interface TableStyleJson {
  headerRowStyle?: TableStyleItem;
  firstRowStyle?: TableStyleItem;
  secondRowStyle?: TableStyleItem;
  lastRowStyle?: TableStyleItem;
  wholeStyle?: TableStyleItem;
}
export interface TableStyleEntry {
  name: string;
  label: string;
  family: 'light' | 'medium' | 'dark';
  style: TableStyleJson;
}

/** Excel : noir puis les six accents du thème Office. */
const COLORS = ['#000000', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];
export const DEFAULT_TABLE_STYLE = 'TableStyleMedium2';

const shade = (hex: string, f: number) => {
  const n = Number.parseInt(hex.slice(1), 16);
  return `#${[16, 8, 0].map((s) => Math.round(((n >> s) & 255) * (1 - f)).toString(16).padStart(2, '0')).join('').toUpperCase()}`;
};
const line = (rgb: string): Line => ({ s: 1, cl: { rgb } });

function light(c: string): TableStyleJson {
  const band = c === '#000000' ? '#D9D9D9' : tint(c, 0.8);
  const edge = line(c === '#000000' ? '#000000' : c);
  return {
    headerRowStyle: { bl: 1, bd: { t: edge, b: edge } },
    firstRowStyle: { bg: { rgb: band } },
    lastRowStyle: { bd: { b: edge } },
  };
}

function medium(c: string): TableStyleJson {
  const head = c === '#000000' ? '#000000' : c;
  const band = c === '#000000' ? '#D9D9D9' : tint(c, 0.8);
  const rule = line(c === '#000000' ? '#BFBFBF' : tint(c, 0.4));
  return {
    headerRowStyle: { bg: { rgb: head }, cl: { rgb: '#FFFFFF' }, bl: 1 },
    firstRowStyle: { bg: { rgb: band }, bd: { b: rule } },
    secondRowStyle: { bd: { b: rule } },
    lastRowStyle: { bd: { b: line(head) } },
  };
}

function dark(c: string): TableStyleJson {
  const base = c === '#000000' ? '#737373' : shade(c, 0.25);
  const band = c === '#000000' ? '#404040' : shade(c, 0.5);
  return {
    headerRowStyle: { bg: { rgb: '#000000' }, cl: { rgb: '#FFFFFF' }, bl: 1, bd: { b: line('#FFFFFF') } },
    firstRowStyle: { bg: { rgb: base }, cl: { rgb: '#FFFFFF' } },
    secondRowStyle: { bg: { rgb: band }, cl: { rgb: '#FFFFFF' } },
  };
}

const FAMILIES = [
  ['light', 'Clair', 'TableStyleLight', light],
  ['medium', 'Moyen', 'TableStyleMedium', medium],
  ['dark', 'Foncé', 'TableStyleDark', dark],
] as const;

export const TABLE_STYLES: TableStyleEntry[] = FAMILIES.flatMap(([family, label, prefix, make]) =>
  COLORS.map((c, i) => ({ name: `${prefix}${i + 1}`, label: `${label} ${i + 1}`, family, style: make(c) })),
);

export const isKnownTableStyle = (name: string | undefined): name is string => !!name && TABLE_STYLES.some((s) => s.name === name);

/**
 * Style de la galerie qui représente un style Excel : les styles 8 à 28 d'Excel sont des variantes
 * des sept couleurs, on prend donc la couleur correspondante.
 */
export function galleryStyleFor(excelName: string | undefined): string {
  const m = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(excelName ?? '');
  if (!m) return DEFAULT_TABLE_STYLE;
  return `TableStyle${m[1]}${((Number(m[2]) - 1) % 7) + 1}`;
}

/** Nom à écrire dans le .xlsx : le style d'origine s'il n'a pas été changé dans Grillon. */
export function excelStyleName(galleryName: string | undefined, original: string | undefined): string {
  if (!galleryName || !isKnownTableStyle(galleryName)) return original ?? DEFAULT_TABLE_STYLE;
  return original && galleryStyleFor(original) === galleryName ? original : galleryName;
}

/**
 * Tri « comme Excel » : quelle plage trier et la première ligne est-elle une ligne d'en-tête ?
 * Logique pure.
 */
export interface Rect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/** Cellule simplifiée : valeur, type (1 texte, 2 nombre, 3 logique) et quelques traits de mise en forme. */
export interface SortCell {
  v: unknown;
  t?: number;
  bold?: boolean;
  fill?: string;
}

const isEmpty = (c: SortCell | undefined) => !c || c.v === null || c.v === undefined || c.v === '';
const isText = (c: SortCell | undefined) => !isEmpty(c) && (c!.t === 1 || (c!.t === undefined && typeof c!.v === 'string'));
const isNumber = (c: SortCell | undefined) => !isEmpty(c) && (c!.t === 2 || typeof c!.v === 'number');

/**
 * Plage à trier, comme Excel : une seule cellule, ou une partie des colonnes d'un tableau, s'étend à tout le tableau
 * (sinon les lignes seraient défaites : les notes ne correspondraient plus aux élèves). Une sélection qui couvre déjà
 * toutes les colonnes du tableau est triée telle quelle.
 */
export function sortRange(selection: Rect, region: Rect): Rect {
  const single = selection.r1 === selection.r2 && selection.c1 === selection.c2;
  const coversColumns = selection.c1 <= region.c1 && selection.c2 >= region.c2;
  if (single || !coversColumns) return region;
  return selection;
}

/**
 * La première ligne est-elle un en-tête ? Comme Excel : oui si elle n'est faite que de texte et qu'elle se distingue des
 * lignes suivantes, soit par la mise en forme (gras, remplissage), soit par le type (texte au-dessus de nombres).
 */
export function hasHeaderRow(rows: SortCell[][]): boolean {
  if (rows.length < 2) return false;
  const [head, ...body] = rows;
  const filled = head.filter((c) => !isEmpty(c));
  if (!filled.length || !filled.every(isText)) return false;
  const first = body[0];
  const formatted = head.some((c, i) => !isEmpty(c) && ((c.bold && !first[i]?.bold) || (c.fill && c.fill !== first[i]?.fill)));
  if (formatted) return true;
  // Une colonne avec un texte en tête et des nombres dessous.
  return head.some((c, i) => isText(c) && body.some((r) => isNumber(r[i])));
}

/** Comparaison de textes à la française, comme Excel : sans tenir compte de la casse ni des tirets et apostrophes, É avec E. */
const collator = new Intl.Collator('fr', { sensitivity: 'accent' });
export function compareFrench(a: string, b: string): number {
  const clean = (s: string) => s.replace(/[-'’]/g, '');
  return collator.compare(clean(a), clean(b));
}

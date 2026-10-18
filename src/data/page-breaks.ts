/**
 * Sauts de page manuels (Mise en page → Sauts de page, Vue → Aperçu des sauts de page).
 * Un saut est l'indice (à partir de 0) de la première ligne / colonne de la nouvelle page :
 * « saut avant la ligne 10 » = 9. C'est aussi l'attribut `id` d'un <brk> du .xlsx.
 * Logique pure, sans DOM ni Univer.
 */

export interface PageBreaks {
  rowBreaks?: number[];
  colBreaks?: number[];
}

/** Plage sélectionnée (indices à partir de 0). */
export interface BreakSelection {
  startRow: number;
  startCol: number;
  /** Lignes entières ou colonnes entières sélectionnées (en-têtes cliqués). */
  wholeRow?: boolean;
  wholeCol?: boolean;
}

/** Liste triée, sans doublon, sans valeur invalide (un saut avant la première ligne n'a pas de sens). */
export function normalizeBreaks(list: readonly number[] | undefined): number[] {
  return [...new Set((list ?? []).filter((n) => Number.isInteger(n) && n > 0))].sort((a, b) => a - b);
}

const add = (list: readonly number[] | undefined, n: number) => normalizeBreaks([...(list ?? []), n]);
const drop = (list: readonly number[] | undefined, n: number) => normalizeBreaks(list).filter((x) => x !== n);

/** Insérer un saut de page : au-dessus de la ligne et à gauche de la colonne de la cellule active (comme Excel). */
export function insertBreaks(b: PageBreaks, sel: BreakSelection): PageBreaks {
  let rowBreaks = normalizeBreaks(b.rowBreaks);
  let colBreaks = normalizeBreaks(b.colBreaks);
  if (!sel.wholeCol) rowBreaks = add(rowBreaks, sel.startRow);
  if (!sel.wholeRow) colBreaks = add(colBreaks, sel.startCol);
  return { rowBreaks, colBreaks };
}

/** Supprimer le saut de page situé au-dessus / à gauche de la cellule active. */
export function removeBreaks(b: PageBreaks, sel: BreakSelection): PageBreaks {
  let rowBreaks = normalizeBreaks(b.rowBreaks);
  let colBreaks = normalizeBreaks(b.colBreaks);
  if (!sel.wholeCol) rowBreaks = drop(rowBreaks, sel.startRow);
  if (!sel.wholeRow) colBreaks = drop(colBreaks, sel.startCol);
  return { rowBreaks, colBreaks };
}

/** Un saut existe-t-il à la position de la sélection ? (pour activer / désactiver « Supprimer ») */
export function hasBreakAt(b: PageBreaks, sel: BreakSelection): boolean {
  return (!sel.wholeCol && normalizeBreaks(b.rowBreaks).includes(sel.startRow)) || (!sel.wholeRow && normalizeBreaks(b.colBreaks).includes(sel.startCol));
}

export function hasBreaks(b: PageBreaks | undefined): boolean {
  return !!(b?.rowBreaks?.length || b?.colBreaks?.length);
}

/** Copie de la mise en page avec ces sauts (les listes vides ne sont pas conservées). */
export function withBreaks<T extends PageBreaks>(page: T, b: PageBreaks): T {
  const out = { ...page };
  const rows = normalizeBreaks(b.rowBreaks);
  const cols = normalizeBreaks(b.colBreaks);
  if (rows.length) out.rowBreaks = rows;
  else delete out.rowBreaks;
  if (cols.length) out.colBreaks = cols;
  else delete out.colBreaks;
  return out;
}

export interface PageSpan {
  /** Première et dernière ligne (ou colonne) de la page, comprises. */
  start: number;
  end: number;
  /** La page commence par un saut manuel. */
  manual: boolean;
}

/**
 * Découpe les lignes (ou colonnes) `from..to` en pages : un saut automatique quand la page est
 * pleine (`capacity`, dans l'unité de `sizes`), un saut manuel où l'utilisateur en a mis un.
 * Un élément plus grand qu'une page occupe seul sa page.
 */
export function paginate(sizes: (i: number) => number, from: number, to: number, capacity: number, manual: readonly number[] = []): PageSpan[] {
  const forced = new Set(normalizeBreaks(manual));
  const pages: PageSpan[] = [];
  let start = from;
  let used = 0;
  let startedManual = false;
  for (let i = from; i <= to; i++) {
    const size = Math.max(0, sizes(i));
    const isManual = i > from && forced.has(i);
    if (i > start && (isManual || used + size > capacity + 1e-6)) {
      pages.push({ start, end: i - 1, manual: startedManual });
      start = i;
      used = 0;
      startedManual = isManual;
    }
    used += size;
  }
  if (to >= from) pages.push({ start, end: to, manual: startedManual });
  return pages;
}

export interface PageRect {
  /** Numéro de page (1, 2, …) : de haut en bas, puis de gauche à droite, comme Excel par défaut. */
  n: number;
  rows: PageSpan;
  cols: PageSpan;
}

export function pageGrid(rowPages: readonly PageSpan[], colPages: readonly PageSpan[]): PageRect[] {
  const out: PageRect[] = [];
  let n = 1;
  for (const cols of colPages) for (const rows of rowPages) out.push({ n: n++, rows, cols });
  return out;
}

// --- Fichier .xlsx : <rowBreaks> et <colBreaks> ---

/** Dernier indice de colonne / ligne d'Excel (valeur de l'attribut `max`). */
const ROW_BREAK_MAX = 16383;
const COL_BREAK_MAX = 1048575;

function breaksXml(tag: 'rowBreaks' | 'colBreaks', list: readonly number[] | undefined): string {
  const ids = normalizeBreaks(list);
  if (!ids.length) return '';
  const max = tag === 'rowBreaks' ? ROW_BREAK_MAX : COL_BREAK_MAX;
  return `<${tag} count="${ids.length}" manualBreakCount="${ids.length}">${ids.map((id) => `<brk id="${id}" max="${max}" man="1"/>`).join('')}</${tag}>`;
}

export const rowBreaksXml = (list: readonly number[] | undefined) => breaksXml('rowBreaks', list);
export const colBreaksXml = (list: readonly number[] | undefined) => breaksXml('colBreaks', list);

/** Sauts manuels (`man="1"`) d'un élément <rowBreaks> ou <colBreaks> ; les sauts automatiques enregistrés par Excel sont ignorés. */
export function readBreaks(sheetXml: string, tag: 'rowBreaks' | 'colBreaks'): number[] {
  const block = new RegExp(`<(?:\\w+:)?${tag}\\b[^>]*>([\\s\\S]*?)</(?:\\w+:)?${tag}>`).exec(sheetXml)?.[1];
  if (!block) return [];
  const ids: number[] = [];
  for (const m of block.matchAll(/<(?:\w+:)?brk\b([^>]*)\/?>/g)) {
    if (!/\bman="(1|true)"/.test(m[1])) continue;
    const id = /\bid="(\d+)"/.exec(m[1])?.[1];
    if (id !== undefined) ids.push(Number(id));
  }
  return normalizeBreaks(ids);
}

/**
 * Valeurs d'erreur des formules (Excel en français) et règles de « Vérification des erreurs ».
 * Univer calcule avec les codes anglais (#NAME?, #VALUE!…) ; Excel en français affiche #NOM?,
 * #VALEUR!… : la traduction ne sert qu'à l'affichage, le .xlsx garde les codes anglais.
 */

export interface ErrorInfo {
  fr: string;
  title: string;
  help: string;
}

export const FORMULA_ERRORS: Record<string, ErrorInfo> = {
  '#DIV/0!': { fr: '#DIV/0!', title: 'Erreur de division par zéro', help: 'La formule ou la fonction divise par zéro ou par des cellules vides.' },
  '#N/A': { fr: '#N/A', title: 'Erreur de valeur non disponible', help: 'Une valeur n’est pas disponible pour la formule ou la fonction (par exemple, une recherche sans résultat).' },
  '#NAME?': { fr: '#NOM?', title: 'Erreur de nom non valide', help: 'La formule contient du texte non reconnu : nom de fonction mal orthographié, nom défini inexistant ou texte sans guillemets.' },
  '#REF!': { fr: '#REF!', title: 'Erreur de référence non valide', help: 'La formule fait référence à des cellules qui n’existent plus (supprimées ou collées par-dessus).' },
  '#VALUE!': { fr: '#VALEUR!', title: 'Erreur de valeur', help: 'Une valeur utilisée dans la formule n’est pas du bon type (du texte à la place d’un nombre, par exemple).' },
  '#NUM!': { fr: '#NOMBRE!', title: 'Erreur de nombre', help: 'La formule contient un nombre non valide ou son résultat est trop grand ou trop petit.' },
  '#NULL!': { fr: '#NUL!', title: 'Erreur d’intersection nulle', help: 'Les deux plages indiquées n’ont aucune cellule en commun.' },
  '#SPILL!': { fr: '#EPARS!', title: 'Erreur de propagation', help: 'Le résultat de la formule matricielle ne peut pas se propager : des cellules de destination ne sont pas vides.' },
  '#CALC!': { fr: '#CALC!', title: 'Erreur de calcul', help: 'Le moteur de calcul a rencontré un cas qu’il ne sait pas traiter (par exemple, un tableau vide).' },
};

export const TEXT_NUMBER: ErrorInfo = {
  fr: '',
  title: 'Nombre stocké sous forme de texte',
  help: 'Le nombre de cette cellule est mis en forme en tant que texte ou précédé d’une apostrophe.',
};

export const isErrorValue = (v: unknown): v is string => typeof v === 'string' && v in FORMULA_ERRORS;

/** Code d'erreur affiché par Excel en français. */
export const frenchError = (v: string): string => FORMULA_ERRORS[v]?.fr ?? v;

/**
 * Nombre lisible dans un texte (règle « Nombre stocké sous forme de texte » d'Excel) :
 * « 123 », « -1,5 », « 1 234,5 », « 12% », « 1.5 ». Renvoie le nombre ou null.
 */
export function numberFromText(text: string): number | null {
  const s = text.trim().replace(/^'/, '').replace(/[   ]/g, '');
  if (!s || !/^[-+]?(\d+([.,]\d+)?|[.,]\d+)(E[-+]?\d+)?%?$/i.test(s)) return null;
  const pct = s.endsWith('%');
  const n = Number(s.replace('%', '').replace(',', '.'));
  if (!Number.isFinite(n)) return null;
  return pct ? n / 100 : n;
}

/** Littéral de formule pour une valeur calculée (Évaluer la formule). */
export function formulaLiteral(value: unknown): string {
  if (Array.isArray(value)) return `{${(value as unknown[][]).map((row) => (Array.isArray(row) ? row : [row]).map(formulaLiteral).join(',')).join(';')}}`;
  if (value === null || value === undefined || value === '') return '0';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'number') return String(value);
  if (isErrorValue(value)) return value;
  return `"${String(value).replace(/"/g, '""')}"`;
}

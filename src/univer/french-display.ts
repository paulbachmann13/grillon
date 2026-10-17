/**
 * Affichage « à la française » des cellules au format Standard.
 *
 * Univer formate en français les nombres qui ont un format (« 0,00 », dates…),
 * mais affiche les nombres au format Standard avec un point (« 1.5 ») et les
 * booléens en anglais. Excel en français affiche « 1,5 » et « VRAI »/« FAUX ».
 * Cet intercepteur ne change que l'affichage : la valeur stockée reste un nombre.
 */
import { INTERCEPTOR_POINT, SheetInterceptorService } from '@univerjs/preset-sheets-core';
import { InterceptorEffectEnum } from '@univerjs/presets';
import type { FUniver } from '@univerjs/presets';
import { frenchError, isErrorValue } from '../data/formula-errors';

const NUMBER = 2;
const BOOLEAN = 3;

/** Longueur maximale du format Standard d'Excel (colonne de largeur par défaut). */
const GENERAL_MAX_CHARS = 11;

/**
 * Format Standard d'Excel, à la française : au plus 11 caractères (signe et virgule compris),
 * décimales arrondies en conséquence ; notation scientifique pour les très grands/petits nombres.
 */
export function formatGeneralFr(n: number): string {
  if (!Number.isFinite(n)) return String(n);
  const abs = Math.abs(n);
  const sign = n < 0 ? '-' : '';
  if (abs !== 0 && (abs >= 1e11 || abs < 1e-9)) {
    const [mant, exp] = n.toExponential(5).split('e');
    const m = String(Number(mant)).replace('.', ',');
    const e = Number(exp);
    return `${m}E${e < 0 ? '-' : '+'}${String(Math.abs(e)).padStart(2, '0')}`;
  }
  // Supprime le bruit des flottants (0,1 + 0,2) avant d'arrondir à la largeur.
  const clean = Number(abs.toPrecision(15));
  const intDigits = Math.floor(clean) === 0 ? 1 : String(Math.floor(clean)).length;
  const decimals = Math.max(0, GENERAL_MAX_CHARS - sign.length - intDigits - 1);
  const text = clean.toFixed(Math.min(decimals, 20)).replace(/\.?0+$/, '');
  return sign + text.replace('.', ',');
}

function isDefaultPattern(pattern: string | undefined): boolean {
  return !pattern || pattern === 'General' || pattern === 'G/通用格式';
}

const installed = new WeakSet<object>();

/** À appeler après la création d'un classeur (les services de feuille sont enregistrés à ce moment-là). */
export function installFrenchDisplay(api: FUniver): void {
  if (installed.has(api)) return;
  installed.add(api);
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const interceptors = injector.get<SheetInterceptorService>(SheetInterceptorService);
  interceptors.intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
    effect: InterceptorEffectEnum.Value,
    // Priorité basse : on passe après le plugin de formats de nombre (priorité 10).
    priority: 1,
    handler: (cell, location, next) => {
      if (!cell || cell.v === null || cell.v === undefined) return next(cell);
      // Erreurs de formule : codes d'Excel en français (#NOM?, #VALEUR!…).
      if (isErrorValue(cell.v) && frenchError(cell.v) !== cell.v) return next({ ...cell, v: frenchError(cell.v) });
      const isBool = cell.t === BOOLEAN;
      const isNum = typeof cell.v === 'number' && (cell.t === NUMBER || cell.t === undefined || cell.t === null);
      if (!isBool && !isNum) return next(cell);
      const styleId = cell.s;
      const style = typeof styleId === 'string' ? location.workbook.getStyles().get(styleId) : styleId;
      if (!isDefaultPattern(style?.n?.pattern ?? undefined)) return next(cell);
      if (isBool) {
        const truthy = cell.v === true || cell.v === 1 || String(cell.v).toUpperCase() === 'TRUE';
        return next({ ...cell, v: truthy ? 'VRAI' : 'FAUX', t: BOOLEAN as never });
      }
      return next({ ...cell, v: formatGeneralFr(cell.v as number), t: NUMBER as never });
    },
  });
}

/**
 * Texte trop long : coupé au bord de la cellule (par défaut) ou débordant sur les cellules vides
 * voisines (comme Excel). C'est un réglage d'affichage mémorisé sur l'appareil : il ne modifie
 * pas le classeur. Une cellule réglée explicitement (Habillage du texte → Déborder, Couper,
 * Renvoyer à la ligne) garde son réglage.
 */
import { INTERCEPTOR_POINT, SheetInterceptorService } from '@univerjs/preset-sheets-core';
import { InterceptorEffectEnum, type FUniver } from '@univerjs/presets';

const STORAGE_KEY = 'grillon-text-overflow';
/** WrapStrategy d'Univer : 0 non défini, 1 déborder, 2 couper, 3 renvoyer à la ligne. */
const CLIP = 2;

let overflow = readSetting();

function readSetting(): boolean {
  try {
    return localStorage.getItem(STORAGE_KEY) === 'overflow';
  } catch {
    return false;
  }
}

export function textOverflowEnabled(): boolean {
  return overflow;
}

const installed = new WeakSet<object>();

export function installTextOverflow(api: FUniver): void {
  if (installed.has(api)) return;
  installed.add(api);
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  injector.get<SheetInterceptorService>(SheetInterceptorService).intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
    effect: InterceptorEffectEnum.Style,
    priority: 1,
    handler: (cell, location, next) => {
      if (overflow || !cell || typeof cell.v !== 'string' || cell.v === '') return next(cell);
      const style = typeof cell.s === 'string' ? location.workbook.getStyles().get(cell.s) : cell.s;
      if (style?.tb) return next(cell);
      return next({ ...cell, interceptorStyle: { ...cell.interceptorStyle, tb: CLIP as never } });
    },
  });
}

/** Active ou désactive le débordement, puis redessine la feuille. */
export function setTextOverflow(api: FUniver, enabled: boolean): void {
  overflow = enabled;
  try {
    if (enabled) localStorage.setItem(STORAGE_KEY, 'overflow');
    else localStorage.removeItem(STORAGE_KEY);
  } catch {
    /* stockage indisponible : le choix vaut pour la session */
  }
  (api.getActiveWorkbook()?.getActiveSheet() as unknown as { refreshCanvas?(): void } | undefined)?.refreshCanvas?.();
}

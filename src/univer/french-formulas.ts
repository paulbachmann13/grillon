/**
 * Formules à la française : on saisit =SOMME(A1:A3;B1) comme dans Excel en français, la barre de
 * formule affiche les noms français, et le classeur garde la forme anglaise (celle du .xlsx).
 * Deux intercepteurs d'Univer : avant l'édition (anglais → français) et après (français → anglais).
 */
import { BEFORE_CELL_EDIT, AFTER_CELL_EDIT } from '@univerjs/sheets';
import { SheetInterceptorService } from '@univerjs/preset-sheets-core';
import type { FUniver } from '@univerjs/presets';
import { IDescriptionService } from '@univerjs/engine-formula';
import { EN_TO_FR, englishToFrench, FR_TO_EN, frenchToEnglish } from '../data/formula-locale';

interface EditedCell {
  f?: string | null;
  v?: unknown;
  [key: string]: unknown;
}

/** Applique `translate` à la formule d'une cellule en cours d'édition (champ `f`, ou texte « =… » dans `v`). */
export function translateEditedCell<T extends EditedCell | null | undefined | void>(cell: T, translate: (formula: string) => string): T {
  if (!cell) return cell;
  const c = cell as EditedCell;
  let out: EditedCell = c;
  if (typeof c.f === 'string' && c.f.startsWith('=')) out = { ...out, f: translate(c.f) };
  if (typeof c.v === 'string' && c.v.startsWith('=')) out = { ...out, v: translate(c.v) };
  return out as T;
}

const installed = new WeakSet<object>();

interface Manager {
  fetchThroughInterceptors(name: unknown, filter?: unknown): (value: unknown, context: unknown) => unknown;
  intercept(name: unknown, interceptor: unknown): unknown;
}

export function installFrenchFormulas(api: FUniver): void {
  if (installed.has(api)) return;
  installed.add(api);
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const writes = injector.get<SheetInterceptorService>(SheetInterceptorService).writeCellInterceptor as unknown as Manager;

  // Validation : on traduit avant les autres intercepteurs (formats de nombre, liens…). `next` ne chaîne pas, il autorise la suite.
  type Next = (value: EditedCell | null | undefined) => EditedCell | null | undefined;
  const after = (value: EditedCell | null | undefined, _context: unknown, next: Next) => next(translateEditedCell(value, frenchToEnglish));
  writes.intercept(AFTER_CELL_EDIT, { priority: 100, handler: after });
  installFrenchFunctionNames(api);

  // Ouverture de l'éditeur : le plugin de formules pose la formule (anglaise) puis arrête la chaîne, aucun intercepteur ne peut
  // donc passer après lui. On traduit donc le résultat final de la chaîne BEFORE_CELL_EDIT.
  const original = writes.fetchThroughInterceptors.bind(writes);
  writes.fetchThroughInterceptors = (name, filter) => {
    const compose = original(name, filter);
    if (name !== BEFORE_CELL_EDIT) return compose;
    return (value, context) => translateEditedCell(compose(value, context) as EditedCell | null | undefined, englishToFrench);
  };
}

interface FunctionListItem {
  name: string;
  desc: string;
  functionType?: number;
}

interface Descriptions {
  getSearchListByName(text: string): FunctionListItem[];
  getSearchListByNameFirstLetter(text: string): FunctionListItem[];
  getSearchListByType(type: number): FunctionListItem[];
  getFunctionInfo(name: string): { functionName: string } | undefined;
  hasFunction(name: string): boolean;
  hasDescription(name: string): boolean;
}

/**
 * Autocomplétion, aide et liste « Insérer une fonction » : les noms de fonctions sont proposés en français
 * (SOMME, SI, RECHERCHEV…). Un nom anglais tapé reste reconnu. Les noms sans équivalent (ABS, MAX…) ne changent pas.
 */
export function installFrenchFunctionNames(api: FUniver): void {
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const d = injector.get<Descriptions>(IDescriptionService);
  if ((d as unknown as { __grillonFr?: boolean }).__grillonFr) return;
  (d as unknown as { __grillonFr?: boolean }).__grillonFr = true;
  const fr = (en: string) => EN_TO_FR.get(en.toUpperCase()) ?? en;
  const translate = (list: FunctionListItem[]) => list.map((item) => ({ ...item, name: fr(item.name) }));
  const search = (text: string, prefix: boolean, original: (t: string) => FunctionListItem[]) => {
    const q = text.toUpperCase().trim();
    const all = translate(original('')) as FunctionListItem[];
    // Le nom anglais est conservé en mémoire pour qu'un nom anglais tapé trouve aussi sa fonction.
    const english = original('');
    return all.filter((item, i) => (prefix ? item.name.toUpperCase().startsWith(q) || english[i].name.toUpperCase().startsWith(q) : item.name.toUpperCase().includes(q) || english[i].name.toUpperCase().includes(q)));
  };
  const byName = d.getSearchListByName.bind(d);
  const byFirst = d.getSearchListByNameFirstLetter.bind(d);
  const byType = d.getSearchListByType.bind(d);
  const info = d.getFunctionInfo.bind(d);
  const hasFunction = d.hasFunction.bind(d);
  const hasDescription = d.hasDescription.bind(d);
  d.getSearchListByName = (text) => search(text, false, byName);
  d.getSearchListByNameFirstLetter = (text) => search(text, true, byFirst);
  d.getSearchListByType = (type) => translate(byType(type));
  d.getFunctionInfo = (name) => {
    const en = FR_TO_EN.get(name.toUpperCase());
    const found = info(en ?? name);
    return found && en ? { ...found, functionName: name.toUpperCase() } : found;
  };
  d.hasFunction = (name) => hasFunction(name) || FR_TO_EN.has(name.toUpperCase());
  d.hasDescription = (name) => hasDescription(name) || FR_TO_EN.has(name.toUpperCase());
}

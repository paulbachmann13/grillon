/**
 * Saisie à la française dans une cellule (voir `src/data/french-entry.ts`) : espace des milliers, montants en euros,
 * point décimal, date sans année, VRAI / FAUX. Intercepteur de validation de cellule, passé avant celui des formats
 * d'Univer (qui relirait « 1.5 » comme 15) : quand la saisie est reconnue, la chaîne s'arrête ici.
 */
import { AFTER_CELL_EDIT } from '@univerjs/sheets';
import { SheetInterceptorService } from '@univerjs/preset-sheets-core';
import type { FUniver } from '@univerjs/presets';
import { parseFrenchEntry } from '../data/french-entry';

const NUMBER = 2;
const BOOLEAN = 3;
const FORCE_STRING = 4;

interface EditedCell {
  v?: unknown;
  t?: number;
  f?: string | null;
  si?: string | null;
  p?: { body?: { dataStream?: string } } | null;
  s?: string | Record<string, unknown> | null;
  [key: string]: unknown;
}

interface Context {
  row: number;
  col: number;
  workbook: { getStyles(): { get(id: string): Record<string, unknown> | undefined } };
  worksheet: { getCellRaw(row: number, col: number): EditedCell | null | undefined };
}

type Next = (value: EditedCell | null | undefined) => EditedCell | null | undefined;

const installed = new WeakSet<object>();

export function installFrenchEntry(api: FUniver): void {
  if (installed.has(api)) return;
  installed.add(api);
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const writes = injector.get<SheetInterceptorService>(SheetInterceptorService).writeCellInterceptor as unknown as { intercept(name: unknown, interceptor: unknown): unknown };
  const handler = (value: EditedCell | null | undefined, context: Context, next: Next) => {
    // Formule, ou texte forcé par une apostrophe (« '1.5 » : Univer a déjà retiré l'apostrophe et marqué la valeur).
    if (!value || value.f || value.si || value.t === FORCE_STRING) return next(value);
    const text = typeof value.v === 'string' ? value.v : value.p?.body?.dataStream?.replace(/\r?\n$/, '').replace(/\r$/, '');
    if (typeof text !== 'string') return next(value);
    const resolve = (s: EditedCell['s']) => (typeof s === 'string' ? (context.workbook.getStyles().get(s) ?? {}) : (s ?? {}));
    const current = resolve(value.s ?? context.worksheet.getCellRaw(context.row, context.col)?.s);
    const pattern = (current.n as { pattern?: string } | undefined)?.pattern;
    // Cellule au format Texte (@) : on garde ce qui est tapé.
    if (pattern === '@') return next(value);
    const parsed = parseFrenchEntry(text);
    if (!parsed) return next(value);
    const style = !pattern && parsed.kind !== 'boolean' && parsed.pattern ? { ...current, n: { pattern: parsed.pattern } } : current;
    const cell: EditedCell = { ...value, p: null, s: Object.keys(style).length ? style : value.s };
    if (parsed.kind === 'boolean') Object.assign(cell, { v: parsed.value ? 1 : 0, t: BOOLEAN });
    else Object.assign(cell, { v: parsed.value, t: NUMBER });
    // Pas de `next` : l'intercepteur des formats d'Univer relirait la valeur (et ferait de « 1.5 » le nombre 15).
    return cell;
  };
  writes.intercept(AFTER_CELL_EDIT, { priority: 90, handler });
}

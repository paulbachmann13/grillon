/**
 * Accueil → Format → Protéger la feuille / Ôter la protection (comme Excel).
 * L'état est gardé dans les métadonnées de la feuille (custom.grillon.protection), réappliqué à
 * l'ouverture et écrit dans le .xlsx. Une feuille protégée n'est plus modifiable.
 */
import type { FUniver } from '@univerjs/presets';
import type { FWorksheet } from '@univerjs/sheets/facade';
import { isRangeUnlocked, type Rect } from '../data/cell-locks';
import { checkPassword, hasPassword, makeProtection, type SheetProtection } from '../data/sheet-protection';
import { t } from '../i18n';
import { alertDialog, box, openDialog } from '../ui/dialog';

interface Custom {
  grillon?: { protection?: SheetProtection; unlocked?: Rect[] } & Record<string, unknown>;
}

type Permission = { protect(): Promise<unknown>; unprotect(): Promise<unknown>; setReadOnly(): Promise<void>; isProtected(): boolean };
const permissionOf = (ws: FWorksheet) => (ws as unknown as { getWorksheetPermission(): Permission }).getWorksheetPermission();

export function protectionOf(ws: FWorksheet): SheetProtection | null {
  return ((ws.getCustomMetadata() ?? {}) as Custom).grillon?.protection ?? null;
}

/** Cellules déverrouillées de la feuille (Format de cellule → Protection). */
export function unlockedOf(ws: FWorksheet): Rect[] {
  return ((ws.getCustomMetadata() ?? {}) as Custom).grillon?.unlocked ?? [];
}

export function storeUnlocked(ws: FWorksheet, rects: Rect[]): void {
  const custom = (ws.getCustomMetadata() ?? {}) as Custom;
  const grillon = { ...(custom.grillon ?? {}) };
  if (rects.length) grillon.unlocked = rects;
  else delete grillon.unlocked;
  ws.setCustomMetadata({ ...custom, grillon });
}

function store(ws: FWorksheet, p: SheetProtection | null): void {
  const custom = (ws.getCustomMetadata() ?? {}) as Custom;
  const grillon = { ...(custom.grillon ?? {}) };
  if (p) grillon.protection = p;
  else delete grillon.protection;
  ws.setCustomMetadata({ ...custom, grillon });
}

/**
 * Univer ne sait pas protéger une feuille en laissant des cellules modifiables : quand il y a des
 * cellules déverrouillées, c'est la garde de Grillon (installProtectionGuard) qui protège la feuille.
 */
async function lock(ws: FWorksheet): Promise<void> {
  if (unlockedOf(ws).length) return;
  const perm = permissionOf(ws);
  if (!perm.isProtected()) await perm.protect();
  await perm.setReadOnly();
}

/** Réapplique la protection des feuilles (à l'ouverture d'un classeur). */
export async function applyStoredProtection(api: FUniver): Promise<void> {
  const wb = api.getActiveWorkbook();
  if (!wb) return;
  for (const ws of wb.getSheets() as unknown as FWorksheet[]) {
    if (protectionOf(ws)) await lock(ws).catch((e) => console.warn('Protection non appliquée', e));
  }
}

function passwordInput(label: string): { el: HTMLLabelElement; input: HTMLInputElement } {
  const el = document.createElement('label');
  el.className = 'g-field';
  const span = document.createElement('span');
  span.textContent = label;
  const input = document.createElement('input');
  input.type = 'password';
  input.autocomplete = 'new-password';
  el.append(span, input);
  return { el, input };
}

export async function protectSheetDialog(api: FUniver, onChange: () => void): Promise<void> {
  const ws = api.getActiveWorkbook()?.getActiveSheet() as unknown as FWorksheet | undefined;
  if (!ws) return;
  if (protectionOf(ws)) return alertDialog(t('protection.title'), t('protection.already'));
  const pw = passwordInput(t('protection.password'));
  const pw2 = passwordInput(t('protection.confirm'));
  await openDialog({
    title: t('protection.title'),
    testId: 'dialog-protect',
    body: box(t('protection.help'), pw.el, pw2.el),
    onOk: async () => {
      if (pw.input.value !== pw2.input.value) return t('protection.mismatch');
      const p = await makeProtection(pw.input.value || undefined);
      store(ws, p);
      await lock(ws);
      onChange();
    },
  });
}

export async function unprotectSheetDialog(api: FUniver, onChange: () => void): Promise<void> {
  const ws = api.getActiveWorkbook()?.getActiveSheet() as unknown as FWorksheet | undefined;
  if (!ws) return;
  const p = protectionOf(ws);
  if (!p) return alertDialog(t('protection.unprotectTitle'), t('protection.notProtected'));
  const unlock = async () => {
    store(ws, null);
    if (permissionOf(ws).isProtected()) await permissionOf(ws).unprotect();
    onChange();
  };
  if (!hasPassword(p)) return unlock();
  const pw = passwordInput(t('protection.password'));
  await openDialog({
    title: t('protection.unprotectTitle'),
    testId: 'dialog-unprotect',
    body: box(pw.el),
    onOk: async () => {
      if (!(await checkPassword(p, pw.input.value))) return t('protection.wrong');
      await unlock();
    },
  });
}

// --- Garde des feuilles protégées qui ont des cellules déverrouillées ---

/** Commandes sans effet sur le contenu, permises sur une feuille protégée. */
const HARMLESS = /undo|redo|copy|activate|zoom|selection|scroll|header-(width|height)|frozen|find|search|show-gridlines/i;
/** Commandes de saisie, permises si toutes les cellules visées sont déverrouillées. */
const EDITS = new Set(['sheet.command.set-range-values', 'sheet.command.clear-selection-content', 'sheet.command.paste', 'univer.command.paste', 'sheet.command.paste-value', 'sheet.command.paste-formula']);
/** Ajustements automatiques déclenchés par une saisie. */
const FOLLOW_UPS = /auto-height|row-is-auto-height/i;

type R = { startRow: number; endRow: number; startColumn: number; endColumn: number };

function guarded(api: FUniver, subUnitId?: string): FWorksheet | null {
  const wb = api.getActiveWorkbook();
  const ws = (subUnitId ? wb?.getSheetBySheetId(subUnitId) : wb?.getActiveSheet()) as unknown as FWorksheet | null | undefined;
  return ws && protectionOf(ws) && unlockedOf(ws).length ? ws : null;
}

function targetRanges(api: FUniver, params: Record<string, unknown> | undefined): R[] {
  const p = params ?? {};
  if (p.range) return [p.range as R];
  if (Array.isArray(p.ranges)) return p.ranges as R[];
  const range = api.getActiveWorkbook()?.getActiveRange()?.getRange() as R | undefined;
  return range ? [range] : [];
}

export function installProtectionGuard(api: FUniver): void {
  let warned = 0;
  const warn = () => {
    if (Date.now() - warned < 1000) return;
    warned = Date.now();
    void alertDialog(t('protection.title'), t('protection.locked'));
  };
  api.addEvent(api.Event.BeforeCommandExecute, (e) => {
    if (e.type !== 0 || !e.id.startsWith('sheet.') && !e.id.startsWith('univer.command.paste')) return;
    const params = e.params as Record<string, unknown> | undefined;
    const ws = guarded(api, params?.subUnitId as string | undefined);
    if (!ws || HARMLESS.test(e.id) || FOLLOW_UPS.test(e.id)) return;
    const unlocked = unlockedOf(ws);
    if (EDITS.has(e.id) && targetRanges(api, params).every((r) => isRangeUnlocked(unlocked, r))) return;
    e.cancel = true;
    warn();
  });
  api.addEvent(api.Event.BeforeSheetEditStart, (e) => {
    const ws = guarded(api, e.worksheet.getSheetId());
    if (!ws) return;
    if (isRangeUnlocked(unlockedOf(ws), { startRow: e.row, endRow: e.row, startColumn: e.column, endColumn: e.column })) return;
    e.cancel = true;
    warn();
  });
}

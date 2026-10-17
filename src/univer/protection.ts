/**
 * Accueil → Format → Protéger la feuille / Ôter la protection (comme Excel).
 * L'état est gardé dans les métadonnées de la feuille (custom.grillon.protection), réappliqué à
 * l'ouverture et écrit dans le .xlsx. Une feuille protégée n'est plus modifiable.
 */
import type { FUniver } from '@univerjs/presets';
import type { FWorksheet } from '@univerjs/sheets/facade';
import { checkPassword, hasPassword, makeProtection, type SheetProtection } from '../data/sheet-protection';
import { t } from '../i18n';
import { alertDialog, box, openDialog } from '../ui/dialog';

interface Custom {
  grillon?: { protection?: SheetProtection } & Record<string, unknown>;
}

type Permission = { protect(): Promise<unknown>; unprotect(): Promise<unknown>; setReadOnly(): Promise<void>; isProtected(): boolean };
const permissionOf = (ws: FWorksheet) => (ws as unknown as { getWorksheetPermission(): Permission }).getWorksheetPermission();

export function protectionOf(ws: FWorksheet): SheetProtection | null {
  return ((ws.getCustomMetadata() ?? {}) as Custom).grillon?.protection ?? null;
}

function store(ws: FWorksheet, p: SheetProtection | null): void {
  const custom = (ws.getCustomMetadata() ?? {}) as Custom;
  const grillon = { ...(custom.grillon ?? {}) };
  if (p) grillon.protection = p;
  else delete grillon.protection;
  ws.setCustomMetadata({ ...custom, grillon });
}

async function lock(ws: FWorksheet): Promise<void> {
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
    await permissionOf(ws).unprotect();
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

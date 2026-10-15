/**
 * Mise en page de la feuille active : stockée dans les métadonnées de la feuille
 * (custom.grillon.page), donc enregistrée avec le classeur et exportée dans le .xlsx.
 */
import type { FUniver } from '@univerjs/presets';
import type { PageSetup } from '../data/page-setup';

interface SheetCustom {
  grillon?: { page?: PageSetup } & Record<string, unknown>;
}

export function getPageSetup(api: FUniver): PageSetup {
  const ws = api.getActiveWorkbook()?.getActiveSheet();
  const custom = ws?.getCustomMetadata() as SheetCustom | undefined;
  return structuredClone(custom?.grillon?.page ?? {});
}

export function savePageSetup(api: FUniver, page: PageSetup): void {
  const ws = api.getActiveWorkbook()?.getActiveSheet();
  if (!ws) return;
  const custom = (ws.getCustomMetadata() ?? {}) as SheetCustom;
  const grillon = { ...(custom.grillon ?? {}) };
  if (Object.keys(page).length) grillon.page = page;
  else delete grillon.page;
  ws.setCustomMetadata({ ...custom, grillon });
}

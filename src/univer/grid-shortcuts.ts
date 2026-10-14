/**
 * Raccourcis d'Excel absents d'Univer :
 * - Ctrl+; : date du jour ; Ctrl+: (ou Ctrl+Maj+;) : heure actuelle.
 * Les autres (Ctrl+C/V/X/Z/Y/B/I/U, F2, F4, Ctrl+flèches, Ctrl+Maj+flèches, Alt+=, Ctrl+D…)
 * sont fournis par Univer ; Ctrl+E (remplissage instantané) par l'onglet Données.
 */
import { IEditorBridgeService } from '@univerjs/preset-sheets-core';
import type { FUniver } from '@univerjs/presets';
import { dateToSerial } from '../io/xlsx/units';

export function installGridShortcuts(api: FUniver): void {
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const editor = injector.get<IEditorBridgeService>(IEditorBridgeService);
  window.addEventListener(
    'keydown',
    (e) => {
      if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
      const isDate = e.key === ';' && !e.shiftKey;
      const isTime = e.key === ':' || (e.key === ';' && e.shiftKey);
      if (!isDate && !isTime) return;
      const active = document.activeElement as HTMLElement | null;
      const typing = !!active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA' || active.isContentEditable);
      if (typing && active?.id !== '__editor___INTERNAL_EDITOR__DOCS_NORMAL') return;
      if (editor.isVisible().visible) return;
      const range = api.getActiveWorkbook()?.getActiveRange();
      if (!range) return;
      e.preventDefault();
      e.stopPropagation();
      const now = new Date();
      if (isDate) {
        range.setValue(dateToSerial(new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()))));
        range.setNumberFormat('dd/mm/yyyy');
      } else {
        range.setValue((now.getHours() * 60 + now.getMinutes()) / 1440);
        range.setNumberFormat('hh:mm');
      }
    },
    { capture: true },
  );
}

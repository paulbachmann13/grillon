/**
 * Raccourcis de niveau application (fichier). Les raccourcis d'édition (Ctrl+C/V/Z/B…)
 * sont gérés par Univer.
 */
import type { DocumentController } from './document';

export function installShortcuts(doc: DocumentController): void {
  window.addEventListener(
    'keydown',
    (e) => {
      const mod = e.ctrlKey || e.metaKey;
      if (!mod || e.altKey) return;
      const key = e.key.toLowerCase();
      if (key === 's') {
        e.preventDefault();
        e.stopPropagation();
        void (e.shiftKey ? doc.saveAs() : doc.save());
      } else if (key === 'p' && !e.shiftKey) {
        e.preventDefault();
        e.stopPropagation();
        void doc.print();
      } else if (key === 'o' && !e.shiftKey) {
        e.preventDefault();
        e.stopPropagation();
        void doc.open();
      }
    },
    { capture: true },
  );
}

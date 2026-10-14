/**
 * Création de l'instance Univer avec uniquement les presets open source (Apache-2.0).
 * Aucun paquet @univerjs-pro n'est utilisé.
 */
import { createUniver, LocaleType, mergeLocales, type FUniver } from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import sheetsCoreFr from '@univerjs/preset-sheets-core/locales/fr-FR';
import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting';
import cfFr from '@univerjs/preset-sheets-conditional-formatting/locales/fr-FR';
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation';
import dvFr from '@univerjs/preset-sheets-data-validation/locales/fr-FR';
import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter';
import filterFr from '@univerjs/preset-sheets-filter/locales/fr-FR';
import { UniverSheetsSortPreset } from '@univerjs/preset-sheets-sort';
import sortFr from '@univerjs/preset-sheets-sort/locales/fr-FR';
import { UniverSheetsFindReplacePreset } from '@univerjs/preset-sheets-find-replace';
import findFr from '@univerjs/preset-sheets-find-replace/locales/fr-FR';
import { UniverSheetsNotePreset } from '@univerjs/preset-sheets-note';
import noteFr from '@univerjs/preset-sheets-note/locales/fr-FR';
import { UniverSheetsHyperLinkPreset } from '@univerjs/preset-sheets-hyper-link';
import linkFr from '@univerjs/preset-sheets-hyper-link/locales/fr-FR';
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing';
import drawingFr from '@univerjs/preset-sheets-drawing/locales/fr-FR';
import { UniverSheetsTablePreset } from '@univerjs/preset-sheets-table';
import tableFr from '@univerjs/preset-sheets-table/locales/fr-FR';

import '@univerjs/preset-sheets-core/lib/index.css';
import '@univerjs/preset-sheets-conditional-formatting/lib/index.css';
import '@univerjs/preset-sheets-data-validation/lib/index.css';
import '@univerjs/preset-sheets-filter/lib/index.css';
import '@univerjs/preset-sheets-sort/lib/index.css';
import '@univerjs/preset-sheets-find-replace/lib/index.css';
import '@univerjs/preset-sheets-note/lib/index.css';
import '@univerjs/preset-sheets-hyper-link/lib/index.css';
import '@univerjs/preset-sheets-drawing/lib/index.css';
import '@univerjs/preset-sheets-table/lib/index.css';

export function setupUniver(container: HTMLElement): FUniver {
  const { univerAPI } = createUniver({
    locale: LocaleType.FR_FR,
    // Annuler/rétablir sans limite (Univer en garde 50 par défaut).
    undoRedoHistoryLimit: Number.POSITIVE_INFINITY,
    locales: {
      [LocaleType.FR_FR]: mergeLocales(sheetsCoreFr, cfFr, dvFr, filterFr, sortFr, findFr, noteFr, linkFr, drawingFr, tableFr),
    },
    presets: [
      UniverSheetsCorePreset({
        container,
        // Ruban en grille : groupes sur deux rangées, comme le ruban d'Excel.
        ribbonType: 'grid',
        workerURL: new Worker(new URL('./worker.ts', import.meta.url), { type: 'module' }),
      }),
      UniverSheetsConditionalFormattingPreset(),
      UniverSheetsDataValidationPreset(),
      UniverSheetsFilterPreset(),
      UniverSheetsSortPreset(),
      UniverSheetsFindReplacePreset(),
      UniverSheetsNotePreset(),
      UniverSheetsHyperLinkPreset(),
      UniverSheetsDrawingPreset(),
      UniverSheetsTablePreset(),
    ],
  });
  return univerAPI;
}

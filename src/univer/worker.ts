/** Worker de calcul des formules Univer (le recalcul ne bloque pas l'interface). */
import { UniverSheetsCoreWorkerPreset } from '@univerjs/preset-sheets-core/worker';
import { UniverSheetsFilterWorkerPreset } from '@univerjs/preset-sheets-filter/worker';
import { createUniver, LocaleType } from '@univerjs/presets';
import { CUSTOM_FUNCTIONS } from './custom-functions';

createUniver({
  // Sans langue, le moteur prend celle de Chine : VPM affichait des ¥.
  locale: LocaleType.FR_FR,
  presets: [UniverSheetsCoreWorkerPreset({ formula: { function: CUSTOM_FUNCTIONS as never } }), UniverSheetsFilterWorkerPreset()],
});

/** Worker de calcul des formules Univer (le recalcul ne bloque pas l'interface). */
import { UniverSheetsCoreWorkerPreset } from '@univerjs/preset-sheets-core/worker';
import { UniverSheetsFilterWorkerPreset } from '@univerjs/preset-sheets-filter/worker';
import { createUniver } from '@univerjs/presets';

createUniver({ presets: [UniverSheetsCoreWorkerPreset(), UniverSheetsFilterWorkerPreset()] });

/**
 * Types et constantes du modèle de données Univer (IWorkbookData).
 *
 * On n'importe que des *types* depuis @univerjs/core : la couche de conversion
 * (src/io) doit pouvoir tourner dans Node (tests Vitest) et dans un Web Worker
 * sans charger le moteur de rendu. Les valeurs d'énumération dont on a besoin
 * sont donc recopiées ici (elles sont stables dans le format de données Univer).
 */
import type {
  IBorderData,
  IBorderStyleData,
  ICellData,
  IColumnData,
  IRange,
  IRowData,
  IStyleData,
  IWorkbookData,
  IWorksheetData,
} from '@univerjs/core';

export type {
  IBorderData,
  IBorderStyleData,
  ICellData,
  IColumnData,
  IRange,
  IRowData,
  IStyleData,
  IWorkbookData,
  IWorksheetData,
};

export const BooleanNumber = { FALSE: 0, TRUE: 1 } as const;

export const CellValueType = { STRING: 1, NUMBER: 2, BOOLEAN: 3, FORCE_STRING: 4 } as const;

export const HorizontalAlign = {
  UNSPECIFIED: 0,
  LEFT: 1,
  CENTER: 2,
  RIGHT: 3,
  JUSTIFIED: 4,
  BOTH: 5,
  DISTRIBUTED: 6,
} as const;

export const VerticalAlign = { UNSPECIFIED: 0, TOP: 1, MIDDLE: 2, BOTTOM: 3 } as const;

export const WrapStrategy = { UNSPECIFIED: 0, OVERFLOW: 1, CLIP: 2, WRAP: 3 } as const;

export const BorderStyleTypes = {
  NONE: 0,
  THIN: 1,
  HAIR: 2,
  DOTTED: 3,
  DASHED: 4,
  DASH_DOT: 5,
  DASH_DOT_DOT: 6,
  DOUBLE: 7,
  MEDIUM: 8,
  MEDIUM_DASHED: 9,
  MEDIUM_DASH_DOT: 10,
  MEDIUM_DASH_DOT_DOT: 11,
  SLANT_DASH_DOT: 12,
  THICK: 13,
} as const;

/** Valeur de LocaleType.FR_FR. */
export const LOCALE_FR = 'frFR';

/** Nom de la ressource Univer qui stocke les noms définis. */
export const DEFINED_NAME_RESOURCE = 'SHEET_DEFINED_NAME_PLUGIN';

/** Taille par défaut d'une feuille Univer (Excel en a beaucoup plus, mais la grille s'étend à la demande). */
export const DEFAULT_ROW_COUNT = 1000;
export const DEFAULT_COLUMN_COUNT = 26;

/** Clé de `custom` dans laquelle on garde les infos xlsx sans équivalent Univer. */
export const GRILLON_CUSTOM_KEY = 'grillon';

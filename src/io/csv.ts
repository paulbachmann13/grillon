/**
 * Import / export CSV.
 *
 * Par défaut on suit les conventions d'Excel en français : séparateur « ; »,
 * virgule décimale, UTF-8 avec BOM (sinon Excel lit mal les accents).
 * À l'import, le séparateur et le format des nombres sont détectés.
 */
import Papa from 'papaparse';
import { CellValueType, DEFAULT_COLUMN_COUNT, DEFAULT_ROW_COUNT, LOCALE_FR, type ICellData, type IWorkbookData, type IWorksheetData } from '../model/univer-types';
import { dateToSerial } from './xlsx/units';

export interface CsvExportOptions {
  delimiter?: string;
  decimal?: '.' | ',';
  bom?: boolean;
}

export interface CsvImportOptions {
  delimiter?: string;
  /** Interpréter « 1,5 » comme 1,5 (par défaut : détecté). */
  decimal?: '.' | ',';
  /** Première ligne = en-têtes : en gras et figée. */
  header?: boolean;
}

export type CsvEncoding = 'utf-8' | 'windows-1252';

/**
 * Décode les octets d'un fichier texte. Sans encodage imposé : UTF-8 si le fichier est de l'UTF-8
 * valide (ou commence par un BOM), sinon Windows-1252 (« ANSI », celui des CSV d'Excel sous Windows).
 */
export function decodeCsvBytes(bytes: Uint8Array, encoding?: CsvEncoding): { text: string; encoding: CsvEncoding } {
  if (!encoding) {
    try {
      return { text: new TextDecoder('utf-8', { fatal: true }).decode(bytes), encoding: 'utf-8' };
    } catch {
      encoding = 'windows-1252';
    }
  }
  return { text: new TextDecoder(encoding).decode(bytes), encoding };
}

/** Séparateur et séparateur décimal détectés, pour pré-remplir l'assistant d'import. */
export function detectCsvFormat(text: string): { delimiter: string; decimal: '.' | ',' } {
  const { rows, delimiter } = csvToRows(text);
  return { delimiter, decimal: detectDecimal(rows, delimiter) };
}

const DATE_PATTERN = 'dd/mm/yyyy';

/** Convertit une chaîne CSV en valeur de cellule typée. */
export function parseCsvValue(raw: string, decimal: '.' | ','): ICellData | null {
  const s = raw.trim();
  if (s === '') return null;
  if (s.startsWith('=') && s.length > 1) return { f: s };
  const upper = s.toUpperCase();
  if (upper === 'TRUE' || upper === 'VRAI') return { v: 1, t: CellValueType.BOOLEAN };
  if (upper === 'FALSE' || upper === 'FAUX') return { v: 0, t: CellValueType.BOOLEAN };

  // Nombres : « 1234.5 », « 1 234,5 », « -12,5 % », « 1e3 ».
  let num = s.replace(/[\s  ]/g, '');
  let percent = false;
  if (num.endsWith('%')) {
    percent = true;
    num = num.slice(0, -1);
  }
  if (decimal === ',') num = num.replace(/\./g, '').replace(',', '.');
  else num = num.replace(/,/g, '');
  if (/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(num) && !/^0\d/.test(num.replace(/^[-+]/, ''))) {
    const n = Number(num);
    if (Number.isFinite(n)) {
      return percent ? { v: n / 100, t: CellValueType.NUMBER, s: { n: { pattern: '0.00%' } } } : { v: n, t: CellValueType.NUMBER };
    }
  }

  // Dates JJ/MM/AAAA et AAAA-MM-JJ.
  let m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s);
  if (m) {
    const d = new Date(Date.UTC(Number(m[3]), Number(m[2]) - 1, Number(m[1])));
    if (d.getUTCDate() === Number(m[1])) return { v: dateToSerial(d), t: CellValueType.NUMBER, s: { n: { pattern: DATE_PATTERN } } };
  }
  m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (m) {
    const d = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
    if (d.getUTCDate() === Number(m[3])) return { v: dateToSerial(d), t: CellValueType.NUMBER, s: { n: { pattern: 'yyyy-mm-dd' } } };
  }
  return { v: raw, t: CellValueType.STRING };
}

/** Devine le séparateur décimal : si on voit des « 1,5 » sans « 1.5 », c'est la virgule. */
function detectDecimal(rows: string[][], delimiter: string): '.' | ',' {
  if (delimiter === ',') return '.';
  let comma = 0;
  let dot = 0;
  for (const row of rows.slice(0, 200)) {
    for (const c of row) {
      if (/^-?\d+,\d+$/.test(c.trim())) comma++;
      else if (/^-?\d+\.\d+$/.test(c.trim())) dot++;
    }
  }
  return comma >= dot ? ',' : '.';
}

export function csvToRows(text: string, delimiter?: string): { rows: string[][]; delimiter: string } {
  const res = Papa.parse<string[]>(text.replace(/^﻿/, ''), {
    delimiter: delimiter ?? '',
    delimitersToGuess: [';', ',', '\t', '|'],
    skipEmptyLines: false,
  });
  const rows = res.data;
  // Papa ajoute une ligne vide finale si le fichier se termine par un saut de ligne.
  while (rows.length && rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === '') rows.pop();
  return { rows, delimiter: res.meta.delimiter };
}

/** CSV → feuille Univer (sans identifiant de classeur). */
export function csvToSheet(text: string, name: string, options: CsvImportOptions = {}): IWorksheetData {
  const { rows, delimiter } = csvToRows(text, options.delimiter);
  const decimal = options.decimal ?? detectDecimal(rows, delimiter);
  const cellData: Record<number, Record<number, ICellData>> = {};
  let maxCol = 0;
  rows.forEach((row, r) => {
    row.forEach((raw, c) => {
      const cell = parseCsvValue(raw, decimal);
      if (!cell) return;
      (cellData[r] ??= {})[c] = cell;
      if (c > maxCol) maxCol = c;
    });
  });
  const sheet = {
    id: `sheet-${Math.random().toString(36).slice(2, 10)}`,
    name,
    cellData,
    rowCount: Math.max(DEFAULT_ROW_COUNT, rows.length + 100),
    columnCount: Math.max(DEFAULT_COLUMN_COUNT, maxCol + 6),
  } as IWorksheetData;
  if (options.header && rows.length) {
    for (const cell of Object.values(cellData[0] ?? {})) cell.s = { ...(cell.s as object | undefined), bl: 1 };
    sheet.freeze = { xSplit: 0, ySplit: 1, startRow: 1, startColumn: -1 };
  }
  return sheet;
}

export function csvToWorkbook(text: string, name: string, options: CsvImportOptions = {}): IWorkbookData {
  const sheet = csvToSheet(text, name.replace(/\.[^.]+$/, '').slice(0, 31) || 'Feuil1', options);
  return {
    id: `wb-${Math.random().toString(36).slice(2, 10)}`,
    name,
    appVersion: '1.0.0',
    locale: LOCALE_FR as IWorkbookData['locale'],
    styles: {},
    sheetOrder: [sheet.id],
    sheets: { [sheet.id]: sheet },
  };
}

function formatNumberForCsv(n: number, decimal: '.' | ','): string {
  const s = String(n);
  return decimal === ',' ? s.replace('.', ',') : s;
}

/**
 * Feuille → CSV. On écrit les valeurs (résultats des formules), comme Excel.
 * `display` permet de fournir le texte affiché (dates formatées, etc.) quand on l'a.
 */
export function sheetToCsv(
  sheet: Partial<IWorksheetData>,
  options: CsvExportOptions = {},
  display?: (row: number, col: number, cell: ICellData) => string | undefined,
): string {
  const delimiter = options.delimiter ?? ';';
  const decimal = options.decimal ?? ',';
  const cellData = sheet.cellData ?? {};
  const rowKeys = Object.keys(cellData).map(Number);
  if (!rowKeys.length) return options.bom === false ? '' : '﻿';
  const maxRow = Math.max(...rowKeys);
  let maxCol = 0;
  for (const r of rowKeys) for (const c of Object.keys(cellData[r] ?? {})) {
    const cell = cellData[r]![Number(c)];
    if (cell && (cell.v !== undefined && cell.v !== null && cell.v !== '')) maxCol = Math.max(maxCol, Number(c));
  }
  const lines: string[][] = [];
  for (let r = 0; r <= maxRow; r++) {
    const row: string[] = [];
    for (let c = 0; c <= maxCol; c++) {
      const cell = cellData[r]?.[c];
      if (!cell || cell.v === undefined || cell.v === null) {
        row.push('');
        continue;
      }
      const shown = display?.(r, c, cell);
      if (shown !== undefined) row.push(shown);
      else if (cell.t === CellValueType.BOOLEAN) row.push(cell.v ? 'VRAI' : 'FAUX');
      else if (typeof cell.v === 'number') row.push(formatNumberForCsv(cell.v, decimal));
      else row.push(String(cell.v));
    }
    lines.push(row);
  }
  const csv = Papa.unparse(lines, { delimiter, newline: '\r\n', quotes: false });
  return (options.bom === false ? '' : '﻿') + csv;
}

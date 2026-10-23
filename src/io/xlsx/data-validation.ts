/**
 * Validation de données : XML OOXML (<dataValidations>, et la variante x14 d'Excel 2010+)
 * → règles Univer (ressource SHEET_DATA_VALIDATION_PLUGIN = { [sheetId]: règles[] }),
 * et règles Univer → ExcelJS.
 */
import type ExcelJS from 'exceljs';
import type { IRange } from '../../model/univer-types';
import { extractAll, getWorksheetElement, parseXml } from './ooxml';
import { dateToSerial, rangeFromA1, rangeToA1, serialToDate } from './units';

export const DV_RESOURCE = 'SHEET_DATA_VALIDATION_PLUGIN';

export interface UniverDvRule {
  uid: string;
  ranges: (IRange & { sheetId?: string })[];
  type: string;
  operator?: string;
  formula1?: string;
  formula2?: string;
  allowBlank?: boolean;
  showDropDown?: boolean;
  showErrorMessage?: boolean;
  error?: string;
  errorStyle?: number;
  errorTitle?: string;
  showInputMessage?: boolean;
  prompt?: string;
  promptTitle?: string;
}

export type UniverDvResource = Record<string, UniverDvRule[]>;

const ERROR_STYLE_TO_UNIVER: Record<string, number> = { information: 0, stop: 1, warning: 2 };
const ERROR_STYLE_FROM_UNIVER: Record<number, 'information' | 'stop' | 'warning'> = { 0: 'information', 1: 'stop', 2: 'warning' };
const SUPPORTED_TYPES = new Set(['list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom']);

let counter = 0;
function newUid(): string {
  counter = (counter + 1) % 1_000_000;
  return `dv${Date.now().toString(36)}${counter.toString(36)}`;
}

interface RawDv {
  type?: string;
  operator?: string;
  allowBlank?: string;
  showDropDown?: string;
  showInputMessage?: string;
  showErrorMessage?: string;
  errorStyle?: string;
  errorTitle?: string;
  error?: string;
  promptTitle?: string;
  prompt?: string;
  sqref?: string;
  formula1?: string | { f?: string };
  formula2?: string | { f?: string };
}

function formulaText(f: RawDv['formula1']): string | undefined {
  if (f === undefined || f === null) return undefined;
  if (typeof f === 'string') return f;
  if (typeof f === 'object' && typeof f.f === 'string') return f.f;
  return undefined;
}

const truthy = (v: string | undefined) => v === '1' || v === 'true';

function decodeXmlEntities(s: string): string {
  return s
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');
}

/** Bornes de date : numéro de série Excel ⇄ « AAAA-MM-JJ » (format attendu par Univer). */
function serialToIso(text: string): string {
  const n = Number(text);
  if (!Number.isFinite(n)) return text.startsWith('=') ? text : `=${text}`;
  const d = serialToDate(n);
  const iso = d.toISOString();
  return n % 1 ? `${iso.slice(0, 10)} ${iso.slice(11, 19)}` : iso.slice(0, 10);
}

function isoToSerial(text: string): string {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?$/.exec(text.trim());
  if (!m) return text.replace(/^=/, '');
  const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] ?? 0), +(m[5] ?? 0), +(m[6] ?? 0)));
  return String(dateToSerial(d));
}

/** Découpe une liste Excel « "a,b,c" » (guillemets doublés échappés). */
function parseLiteralList(formula: string): string[] | null {
  const m = /^"((?:[^"]|"")*)"$/.exec(formula.trim());
  if (!m) return null;
  return m[1]
    .replace(/""/g, '"')
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);
}

function toUniver(raw: RawDv, sheetId: string, sqref: string): UniverDvRule | null {
  const type = raw.type ?? 'none';
  if (!SUPPORTED_TYPES.has(type)) return null;
  const ranges = sqref
    .split(/\s+/)
    .filter(Boolean)
    .flatMap((ref) => {
      // Plage devenue « #REF! » (lignes supprimées dans Excel) : ignorée.
      try {
        return [{ ...rangeFromA1(ref), sheetId }];
      } catch {
        return [];
      }
    });
  if (!ranges.length) return null;
  const f1 = formulaText(raw.formula1);
  const f2 = formulaText(raw.formula2);
  const rule: UniverDvRule = { uid: newUid(), ranges, type };
  if (type === 'list') {
    const items = f1 !== undefined ? parseLiteralList(decodeXmlEntities(f1)) : null;
    if (items) rule.formula1 = JSON.stringify(items);
    else if (f1) {
      rule.formula1 = `=${decodeXmlEntities(f1)}`;
      rule.formula2 = '1';
    }
    // Attention : dans OOXML, showDropDown="1" signifie… masquer la flèche.
    rule.showDropDown = !truthy(raw.showDropDown);
  } else {
    rule.operator = raw.operator ?? 'between';
    const conv = (f: string) => (type === 'custom' ? `=${decodeXmlEntities(f)}` : type === 'date' ? serialToIso(decodeXmlEntities(f)) : decodeXmlEntities(f));
    if (f1 !== undefined) rule.formula1 = conv(f1);
    if (f2 !== undefined) rule.formula2 = conv(f2);
  }
  if (truthy(raw.allowBlank)) rule.allowBlank = true;
  if (truthy(raw.showErrorMessage)) rule.showErrorMessage = true;
  if (raw.error) rule.error = raw.error;
  if (raw.errorTitle) rule.errorTitle = raw.errorTitle;
  rule.errorStyle = ERROR_STYLE_TO_UNIVER[raw.errorStyle ?? 'stop'] ?? 1;
  if (truthy(raw.showInputMessage)) rule.showInputMessage = true;
  if (raw.prompt) rule.prompt = raw.prompt;
  if (raw.promptTitle) rule.promptTitle = raw.promptTitle;
  return rule;
}

/** Lit les validations d'une feuille directement dans son XML. */
export function readDataValidations(sheetXml: string, sheetId: string): UniverDvRule[] {
  const rules: UniverDvRule[] = [];
  const main = getWorksheetElement(sheetXml, 'dataValidations');
  if (main) {
    const doc = parseXml<{ dataValidations?: { dataValidation?: RawDv[] } }>(main);
    for (const raw of doc.dataValidations?.dataValidation ?? []) {
      const rule = raw.sqref ? toUniver(raw, sheetId, raw.sqref) : null;
      if (rule) rules.push(rule);
    }
  }
  // Excel 2010+ range dans <extLst> les listes qui pointent vers une autre feuille (x14).
  const ext = getWorksheetElement(sheetXml, 'extLst');
  if (ext) {
    for (const el of extractAll(ext, 'dataValidations')) {
      const doc = parseXml<{ dataValidations?: { dataValidation?: (RawDv & { sqref?: string })[] } }>(el.xml);
      for (const raw of doc.dataValidations?.dataValidation ?? []) {
        // En x14, la plage est dans <xm:sqref> (élément) et les formules dans <xm:f>.
        const sqref = typeof raw.sqref === 'string' ? raw.sqref : undefined;
        const rule = sqref ? toUniver(raw, sheetId, sqref) : null;
        if (rule) rules.push(rule);
      }
    }
  }
  return rules;
}

/** Écrit les validations Univer d'une feuille via ExcelJS (une entrée par plage). */
export function exportDataValidations(ws: ExcelJS.Worksheet, rules: UniverDvRule[]): void {
  const dvs = (ws as unknown as { dataValidations: { add(address: string, v: ExcelJS.DataValidation): void } }).dataValidations;
  for (const rule of rules) {
    if (!SUPPORTED_TYPES.has(rule.type) && rule.type !== 'listMultiple') continue;
    const type = (rule.type === 'listMultiple' ? 'list' : rule.type) as ExcelJS.DataValidation['type'];
    const formulae: string[] = [];
    if (type === 'list') {
      if (rule.formula2 === '1' || (rule.formula1 ?? '').startsWith('=')) {
        formulae.push((rule.formula1 ?? '').replace(/^=/, ''));
      } else {
        let items: string[] = [];
        try {
          const parsed = JSON.parse(rule.formula1 ?? '[]');
          if (Array.isArray(parsed)) items = parsed.map(String);
        } catch {
          items = (rule.formula1 ?? '').split(',').map((s) => s.trim());
        }
        formulae.push(`"${items.filter(Boolean).join(',').replace(/"/g, '""')}"`);
      }
    } else {
      // ExcelJS convertit lui-même les bornes de date : il attend un objet Date.
      const conv = (f: string): string => {
        if (type !== 'date') return f.replace(/^=/, '');
        const serial = isoToSerial(f);
        return (/^-?\d+(\.\d+)?$/.test(serial) ? serialToDate(Number(serial)) : serial) as unknown as string;
      };
      if (rule.formula1 !== undefined && rule.formula1 !== '') formulae.push(conv(rule.formula1));
      if (rule.formula2 !== undefined && rule.formula2 !== '' && (rule.operator === 'between' || rule.operator === 'notBetween')) {
        formulae.push(conv(rule.formula2));
      }
    }
    const dv: ExcelJS.DataValidation = {
      type,
      allowBlank: !!rule.allowBlank,
      formulae,
      showErrorMessage: !!rule.showErrorMessage,
      showInputMessage: !!rule.showInputMessage,
    };
    if (type !== 'list' && rule.operator) dv.operator = rule.operator as ExcelJS.DataValidationOperator;
    if (rule.error) dv.error = rule.error;
    if (rule.errorTitle) dv.errorTitle = rule.errorTitle;
    if (rule.errorStyle !== undefined && rule.errorStyle !== 1) dv.errorStyle = ERROR_STYLE_FROM_UNIVER[rule.errorStyle];
    if (rule.prompt) dv.prompt = rule.prompt;
    if (rule.promptTitle) dv.promptTitle = rule.promptTitle;
    if (type === 'list' && rule.showDropDown === false) (dv as { showDropDown?: boolean }).showDropDown = true;
    for (const r of rule.ranges ?? []) dvs.add(rangeToA1(r), dv);
  }
}

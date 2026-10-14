/**
 * Mises en forme conditionnelles : ExcelJS (conditionalFormattings) ⇄ Univer
 * (ressource SHEET_CONDITIONAL_FORMATTING_PLUGIN = { [sheetId]: règles[] }).
 *
 * Ordre : dans Univer, la règle d'indice 0 est la plus prioritaire, comme la priorité 1 d'Excel.
 * Les règles qu'ExcelJS ne sait pas écrire (doublons, « commence par »…) sont exportées
 * sous forme de règle « formule » équivalente, qu'Excel et LibreOffice comprennent.
 */
import type ExcelJS from 'exceljs';
import type { IRange, IStyleData } from '../../model/univer-types';
import { cssToArgb } from './colors';
import { styleFromUniver, styleToUniver } from './styles';
import { columnToLetters, rangeFromA1, rangeToA1, toA1 } from './units';

export const CF_RESOURCE = 'SHEET_CONDITIONAL_FORMATTING_PLUGIN';

// --- Modèle Univer (sous-ensemble utilisé, cf. @univerjs/sheets-conditional-formatting) ---

type ValueType = 'num' | 'min' | 'max' | 'percent' | 'percentile' | 'formula';
interface ValueConfig {
  type: ValueType;
  value?: number | string;
}
type NumberOperator = 'greaterThan' | 'greaterThanOrEqual' | 'lessThan' | 'lessThanOrEqual' | 'notBetween' | 'between' | 'equal' | 'notEqual';

export type UniverCfRule =
  | { type: 'highlightCell'; subType: 'number'; operator: NumberOperator; value?: number | [number, number]; style: IStyleData }
  | { type: 'highlightCell'; subType: 'text'; operator: string; value?: string; style: IStyleData }
  | { type: 'highlightCell'; subType: 'timePeriod'; operator: string; style: IStyleData }
  | { type: 'highlightCell'; subType: 'formula'; value: string; style: IStyleData; operator?: string }
  | { type: 'highlightCell'; subType: 'rank'; isBottom: boolean; isPercent: boolean; value: number; style: IStyleData }
  | { type: 'highlightCell'; subType: 'average'; operator: NumberOperator; style: IStyleData }
  | { type: 'highlightCell'; subType: 'duplicateValues' | 'uniqueValues'; style: IStyleData; operator?: string; value?: string }
  | {
      type: 'dataBar';
      isShowValue: boolean;
      config: { min: ValueConfig; max: ValueConfig; isGradient: boolean; positiveColor: string; nativeColor: string };
    }
  | { type: 'colorScale'; config: { index: number; color: string; value: ValueConfig }[] }
  | {
      type: 'iconSet';
      isShowValue: boolean;
      config: { operator: NumberOperator; value: ValueConfig; iconType: string; iconId: string }[];
    };

export interface UniverCfEntry {
  cfId: string;
  ranges: IRange[];
  stopIfTrue: boolean;
  rule: UniverCfRule;
}

export type UniverCfResource = Record<string, UniverCfEntry[]>;

// --- Utilitaires ---

const NUMBER_OPERATORS: NumberOperator[] = ['greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual', 'notBetween', 'between', 'equal', 'notEqual'];

const OPERATOR_SYMBOL: Record<string, string> = {
  greaterThan: '>',
  greaterThanOrEqual: '>=',
  lessThan: '<',
  lessThanOrEqual: '<=',
  equal: '=',
  notEqual: '<>',
};

const EXCEL_VALUE_TYPES: Record<string, ValueType> = {
  num: 'num',
  min: 'min',
  max: 'max',
  percent: 'percent',
  percentile: 'percentile',
  formula: 'formula',
  autoMin: 'min',
  autoMax: 'max',
};

let idCounter = 0;
function newCfId(): string {
  idCounter = (idCounter + 1) % 1_000_000;
  return `cf${Date.now().toString(36)}${idCounter.toString(36)}`;
}

function absRef(row: number, col: number): string {
  return `$${columnToLetters(col)}$${row + 1}`;
}

function absRange(r: IRange): string {
  const a = absRef(r.startRow, r.startColumn);
  const b = absRef(r.endRow, r.endColumn);
  return a === b ? a : `${a}:${b}`;
}

function quote(text: string): string {
  return `"${text.replace(/"/g, '""')}"`;
}

function isNumeric(s: string | undefined): boolean {
  return s !== undefined && /^\s*[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?\s*$/i.test(s);
}

/** Style d'un format différentiel (dxf). Dans un dxf, la couleur d'un remplissage plein est dans bgColor. */
function dxfToUniver(style: Partial<ExcelJS.Style> | undefined, theme: string[]): IStyleData {
  if (!style) return {};
  const fill = style.fill as ExcelJS.FillPattern | undefined;
  const fixed: Partial<ExcelJS.Style> = { ...style };
  if (fill?.type === 'pattern') {
    fixed.fill = { ...fill, pattern: 'solid', fgColor: fill.bgColor ?? fill.fgColor };
  }
  const s = styleToUniver(fixed, theme) ?? {};
  // Les formats conditionnels n'ont pas de police/taille propre dans Univer.
  delete s.ff;
  delete s.fs;
  return s;
}

function dxfFromUniver(style: IStyleData | undefined): Partial<ExcelJS.Style> | undefined {
  const s = styleFromUniver(style);
  if (!s) return undefined;
  const argb = cssToArgb(style?.bg?.rgb);
  if (argb) s.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb }, bgColor: { argb } };
  return s;
}

function cfvoToUniver(cfvo: { type?: string; value?: number } | undefined, fallback: ValueType): ValueConfig {
  const type = EXCEL_VALUE_TYPES[cfvo?.type ?? ''] ?? fallback;
  const v = cfvo?.value;
  return type === 'min' || type === 'max' || v === undefined || Number.isNaN(v) ? { type } : { type, value: v };
}

function cfvoFromUniver(v: ValueConfig): { type: string; value?: number } {
  if (v.type === 'min' || v.type === 'max') return { type: v.type };
  if (v.type === 'formula') {
    const n = Number(String(v.value ?? '').replace(/^=/, ''));
    return Number.isFinite(n) ? { type: 'num', value: n } : { type: 'num', value: 0 };
  }
  const n = Number(v.value ?? 0);
  return { type: v.type, value: Number.isFinite(n) ? n : 0 };
}

function colorToCss(c: { argb?: string } | undefined, fallback: string): string {
  const argb = c?.argb;
  return argb && /^[0-9A-Fa-f]{8}$/.test(argb) ? `#${argb.slice(2).toUpperCase()}` : fallback;
}

/** Texte d'une règle « contient » déduit de sa formule (ExcelJS ne lit pas l'attribut text). */
function textFromFormula(formula: string | undefined): string | undefined {
  if (!formula) return undefined;
  const m = /SEARCH\("((?:[^"]|"")*)"/i.exec(formula) ?? /LEN\("((?:[^"]|"")*)"\)/i.exec(formula);
  return m ? m[1].replace(/""/g, '"') : undefined;
}

// --- Import ---

type ExcelRule = ExcelJS.ConditionalFormattingRule & {
  formulae?: string[];
  style?: Partial<ExcelJS.Style>;
  operator?: string;
  text?: string;
  timePeriod?: string;
  rank?: number;
  percent?: boolean;
  bottom?: boolean;
  aboveAverage?: boolean;
  cfvo?: { type: string; value?: number }[];
  color?: { argb?: string } | { argb?: string }[];
  iconSet?: string;
  reverse?: boolean;
  showValue?: boolean;
  gradient?: boolean;
  negativeFillColor?: { argb?: string };
  priority?: number;
  stopIfTrue?: boolean;
};

const STR = '"((?:[^"]|"")*)"';
const REF = '\\$?[A-Z]{1,3}\\$?\\d+';

/**
 * Reconnaît les formules qu'écrit l'export (et Excel lui-même pour certaines règles)
 * afin de retrouver le type de règle d'origine.
 */
function recognizeExpression(formula: string, topLeft: string, style: IStyleData): UniverCfRule | null {
  const tl = topLeft.replace(/\$/g, '');
  const unq = (s: string) => s.replace(/""/g, '"');
  let m = new RegExp(`^COUNTIF\\(${REF}(?::${REF})?,(${REF})\\)(>1|=1)$`, 'i').exec(formula);
  if (m && m[1].replace(/\$/g, '') === tl) {
    return { type: 'highlightCell', subType: m[2] === '>1' ? 'duplicateValues' : 'uniqueValues', style };
  }
  m = new RegExp(`^LEFT\\((${REF}),LEN\\(${STR}\\)\\)=${STR}$`, 'i').exec(formula);
  if (m && m[1] === tl && m[2] === m[3]) return { type: 'highlightCell', subType: 'text', operator: 'beginsWith', value: unq(m[2]), style };
  m = new RegExp(`^RIGHT\\((${REF}),LEN\\(${STR}\\)\\)=${STR}$`, 'i').exec(formula);
  if (m && m[1] === tl && m[2] === m[3]) return { type: 'highlightCell', subType: 'text', operator: 'endsWith', value: unq(m[2]), style };
  m = new RegExp(`^ISERROR\\(SEARCH\\(${STR},(${REF})\\)\\)$`, 'i').exec(formula);
  if (m && m[2] === tl) return { type: 'highlightCell', subType: 'text', operator: 'notContainsText', value: unq(m[1]), style };
  return null;
}

function ruleToUniver(rule: ExcelRule, topLeft: string, theme: string[]): UniverCfRule | null {
  const style = () => dxfToUniver(rule.style, theme);
  const f = rule.formulae ?? [];
  switch (rule.type as string) {
    case 'expression':
      if (!f[0]) return null;
      return recognizeExpression(f[0], topLeft, style()) ?? { type: 'highlightCell', subType: 'formula', value: `=${f[0]}`, style: style() };
    case 'cellIs': {
      const op = rule.operator as NumberOperator;
      if (!NUMBER_OPERATORS.includes(op)) return null;
      if ((op === 'between' || op === 'notBetween') && isNumeric(f[0]) && isNumeric(f[1])) {
        return { type: 'highlightCell', subType: 'number', operator: op, value: [Number(f[0]), Number(f[1])], style: style() };
      }
      if (op !== 'between' && op !== 'notBetween' && isNumeric(f[0])) {
        return { type: 'highlightCell', subType: 'number', operator: op, value: Number(f[0]), style: style() };
      }
      const text = /^"((?:[^"]|"")*)"$/.exec(f[0] ?? '')?.[1];
      if ((op === 'equal' || op === 'notEqual') && text !== undefined) {
        return { type: 'highlightCell', subType: 'text', operator: op, value: text.replace(/""/g, '"'), style: style() };
      }
      // Comparaison à une référence ou une formule : règle « formule » équivalente.
      if (op === 'between' || op === 'notBetween') {
        const inside = `AND(${topLeft}>=${f[0]},${topLeft}<=${f[1]})`;
        return { type: 'highlightCell', subType: 'formula', value: `=${op === 'between' ? inside : `NOT(${inside})`}`, style: style() };
      }
      return { type: 'highlightCell', subType: 'formula', value: `=${topLeft}${OPERATOR_SYMBOL[op]}${f[0]}`, style: style() };
    }
    case 'top10':
      return { type: 'highlightCell', subType: 'rank', isBottom: !!rule.bottom, isPercent: !!rule.percent, value: rule.rank ?? 10, style: style() };
    case 'aboveAverage':
      return { type: 'highlightCell', subType: 'average', operator: rule.aboveAverage === false ? 'lessThan' : 'greaterThan', style: style() };
    case 'containsText': {
      const op = rule.operator ?? 'containsText';
      if (op === 'containsText') {
        const text = rule.text ?? textFromFormula(f[0]);
        if (text !== undefined) return { type: 'highlightCell', subType: 'text', operator: 'containsText', value: text, style: style() };
        return f[0] ? { type: 'highlightCell', subType: 'formula', value: `=${f[0]}`, style: style() } : null;
      }
      return { type: 'highlightCell', subType: 'text', operator: op, style: style() };
    }
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith': {
      const text = rule.text ?? textFromFormula(f[0]);
      if (text !== undefined) return { type: 'highlightCell', subType: 'text', operator: rule.type as string, value: text, style: style() };
      return f[0] ? { type: 'highlightCell', subType: 'formula', value: `=${f[0]}`, style: style() } : null;
    }
    case 'timePeriod':
      return rule.timePeriod ? { type: 'highlightCell', subType: 'timePeriod', operator: rule.timePeriod, style: style() } : null;
    case 'duplicateValues':
    case 'uniqueValues':
      return { type: 'highlightCell', subType: rule.type as 'duplicateValues' | 'uniqueValues', style: style() };
    case 'dataBar': {
      const color = Array.isArray(rule.color) ? rule.color[0] : rule.color;
      return {
        type: 'dataBar',
        isShowValue: rule.showValue !== false,
        config: {
          min: cfvoToUniver(rule.cfvo?.[0], 'min'),
          max: cfvoToUniver(rule.cfvo?.[1], 'max'),
          isGradient: rule.gradient !== false,
          positiveColor: colorToCss(color, '#638EC6'),
          nativeColor: colorToCss(rule.negativeFillColor, '#FF0000'),
        },
      };
    }
    case 'colorScale': {
      const colors = (Array.isArray(rule.color) ? rule.color : []) as { argb?: string }[];
      const cfvo = rule.cfvo ?? [];
      if (cfvo.length < 2) return null;
      return {
        type: 'colorScale',
        config: cfvo.map((c, i) => ({ index: i, color: colorToCss(colors[i], '#FFFFFF'), value: cfvoToUniver(c, i === 0 ? 'min' : 'max') })),
      };
    }
    case 'iconSet': {
      const cfvo = rule.cfvo ?? [];
      const n = cfvo.length;
      if (n < 2) return null;
      const iconType = rule.iconSet ?? '3TrafficLights1';
      // Univer teste les seuils du plus haut au plus bas ; l'icône 0 d'Univer est la meilleure,
      // l'icône 0 d'Excel la plus basse (sauf si « reverse »).
      const config = [];
      for (let k = 0; k < n; k++) {
        const excelIndex = n - 1 - k;
        const iconIndex = rule.reverse ? n - 1 - k : k;
        config.push({
          operator: 'greaterThanOrEqual' as NumberOperator,
          value: cfvoToUniver(cfvo[excelIndex], 'num'),
          iconType,
          iconId: String(rule.reverse ? n - 1 - iconIndex : iconIndex),
        });
      }
      return { type: 'iconSet', isShowValue: rule.showValue !== false, config };
    }
    default:
      return null;
  }
}

/** Mises en forme conditionnelles d'une feuille ExcelJS → règles Univer. */
export function importConditionalFormats(ws: ExcelJS.Worksheet, sheetId: string, theme: string[]): UniverCfEntry[] {
  const groups = (ws as unknown as { conditionalFormattings?: { ref: string; rules: ExcelRule[] }[] }).conditionalFormattings ?? [];
  const collected: { priority: number; entry: UniverCfEntry }[] = [];
  let seq = 0;
  for (const group of groups) {
    const ranges = group.ref
      .split(/\s+/)
      .filter(Boolean)
      .map((ref) => ({ ...rangeFromA1(ref), sheetId }));
    if (!ranges.length) continue;
    const topLeft = toA1(ranges[0].startRow, ranges[0].startColumn);
    for (const rule of group.rules ?? []) {
      const converted = ruleToUniver(rule, topLeft, theme);
      seq++;
      if (!converted) continue;
      collected.push({
        priority: rule.priority ?? 1_000_000 + seq,
        entry: { cfId: newCfId(), ranges, stopIfTrue: !!rule.stopIfTrue, rule: converted },
      });
    }
  }
  return collected.sort((a, b) => a.priority - b.priority).map((c) => c.entry);
}

// --- Export ---

function ruleFromUniver(entry: UniverCfEntry, priority: number): ExcelRule | null {
  const r = entry.rule;
  const first = entry.ranges[0];
  const tl = toA1(first.startRow, first.startColumn);
  const base = { priority } as ExcelRule;
  const expr = (formula: string, style: IStyleData): ExcelRule =>
    ({ ...base, type: 'expression', formulae: [formula.replace(/^=/, '')], style: dxfFromUniver(style) }) as ExcelRule;

  if (r.type === 'highlightCell') {
    switch (r.subType) {
      case 'number': {
        const values = Array.isArray(r.value) ? r.value : [r.value ?? 0];
        return { ...base, type: 'cellIs', operator: r.operator, formulae: values.map(String), style: dxfFromUniver(r.style) } as ExcelRule;
      }
      case 'text': {
        const v = r.value ?? '';
        switch (r.operator) {
          case 'containsText':
            return { ...base, type: 'containsText', operator: 'containsText', text: v, formulae: [`NOT(ISERROR(SEARCH(${quote(v)},${tl})))`], style: dxfFromUniver(r.style) } as ExcelRule;
          case 'notContainsText':
            return expr(`ISERROR(SEARCH(${quote(v)},${tl}))`, r.style);
          case 'beginsWith':
            return expr(`LEFT(${tl},LEN(${quote(v)}))=${quote(v)}`, r.style);
          case 'endsWith':
            return expr(`RIGHT(${tl},LEN(${quote(v)}))=${quote(v)}`, r.style);
          case 'equal':
          case 'notEqual':
            return { ...base, type: 'cellIs', operator: r.operator, formulae: [quote(v)], style: dxfFromUniver(r.style) } as ExcelRule;
          case 'containsBlanks':
          case 'notContainsBlanks':
          case 'containsErrors':
          case 'notContainsErrors':
            return { ...base, type: 'containsText', operator: r.operator, style: dxfFromUniver(r.style) } as ExcelRule;
          default:
            return null;
        }
      }
      case 'timePeriod':
        return { ...base, type: 'timePeriod', timePeriod: r.operator, style: dxfFromUniver(r.style) } as ExcelRule;
      case 'formula':
        return r.value ? expr(r.value, r.style) : null;
      case 'rank':
        return { ...base, type: 'top10', rank: r.value, percent: r.isPercent, bottom: r.isBottom, style: dxfFromUniver(r.style) } as ExcelRule;
      case 'average': {
        if (r.operator === 'greaterThan' || r.operator === 'greaterThanOrEqual') return { ...base, type: 'aboveAverage', aboveAverage: true, style: dxfFromUniver(r.style) } as ExcelRule;
        if (r.operator === 'lessThan' || r.operator === 'lessThanOrEqual') return { ...base, type: 'aboveAverage', aboveAverage: false, style: dxfFromUniver(r.style) } as ExcelRule;
        const all = entry.ranges.map(absRange).join(',');
        return expr(`${tl}${OPERATOR_SYMBOL[r.operator] ?? '='}AVERAGE(${all})`, r.style);
      }
      case 'duplicateValues':
      case 'uniqueValues': {
        const all = absRange(first);
        return expr(`COUNTIF(${all},${tl})${r.subType === 'duplicateValues' ? '>1' : '=1'}`, r.style);
      }
      default:
        return null;
    }
  }
  if (r.type === 'dataBar') {
    return {
      ...base,
      type: 'dataBar',
      cfvo: [cfvoFromUniver(r.config.min), cfvoFromUniver(r.config.max)],
      color: { argb: cssToArgb(r.config.positiveColor) ?? 'FF638EC6' },
      negativeFillColor: { argb: cssToArgb(r.config.nativeColor) ?? 'FFFF0000' },
      gradient: r.config.isGradient,
      showValue: r.isShowValue,
    } as unknown as ExcelRule;
  }
  if (r.type === 'colorScale') {
    const sorted = [...r.config].sort((a, b) => a.index - b.index);
    return {
      ...base,
      type: 'colorScale',
      cfvo: sorted.map((c) => cfvoFromUniver(c.value)),
      color: sorted.map((c) => ({ argb: cssToArgb(c.color) ?? 'FFFFFFFF' })),
    } as unknown as ExcelRule;
  }
  if (r.type === 'iconSet') {
    const n = r.config.length;
    if (n < 2) return null;
    const iconType = r.config[0].iconType;
    // Inverse de l'import : config[k] → seuil de l'icône Excel n-1-k.
    const cfvo = new Array(n);
    for (let k = 0; k < n; k++) cfvo[n - 1 - k] = cfvoFromUniver(r.config[k].value);
    cfvo[0] = { type: 'percent', value: 0 };
    const reverse = r.config[0].iconId !== '0';
    return { ...base, type: 'iconSet', iconSet: iconType, cfvo, reverse, showValue: r.isShowValue } as unknown as ExcelRule;
  }
  return null;
}

/** Règles Univer d'une feuille → ExcelJS. */
export function exportConditionalFormats(ws: ExcelJS.Worksheet, entries: UniverCfEntry[]): void {
  entries.forEach((entry, i) => {
    if (!entry.ranges?.length) return;
    const rule = ruleFromUniver(entry, i + 1);
    if (!rule) return;
    if (entry.stopIfTrue) rule.stopIfTrue = true;
    ws.addConditionalFormatting({ ref: entry.ranges.map((r) => rangeToA1(r)).join(' '), rules: [rule] });
  });
}

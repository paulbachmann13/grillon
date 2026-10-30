/**
 * Fonctions Excel absentes du moteur Univer, ajoutées au moteur de formules.
 *
 * Le calcul a lieu dans le Web Worker : les classes de fonction sont passées à la config
 * du preset côté worker (calcul) ET côté page (analyse des formules, autocomplétion).
 * La logique pure (isoCeiling, trimRange) est testée à part.
 */
import { numfmt } from '@univerjs/core';
import {
  ArrayValueObject,
  BaseFunction,
  ErrorValueObject,
  FunctionType,
  functionDate,
  functionLookup,
  functionMath,
  functionStatistical,
  functionText,
  NumberValueObject,
  StringValueObject,
  type BaseValueObject,
  type IFunctionInfo,
  type IFunctionNames,
} from '@univerjs/engine-formula';
import { englishToFrench } from '../data/formula-locale';
import { amorDegrc, fixedText, frenchCriteria, frenchFormatToEnglish, frenchNumberText, parseFrenchDate, parseFrenchNumber, searchWithWildcards } from '../data/french-values';

type Scalar = number | string | boolean | null;

/** ISO.CEILING(nombre; [précision]) : arrondi à l'excès, quel que soit le signe de la précision. */
export function isoCeiling(x: number, sig = 1): number {
  const s = Math.abs(sig);
  if (s === 0) return 0;
  return Math.ceil(x / s) * s;
}

/** TRIMRANGE(plage) : retire les lignes et colonnes vides en bordure. */
export function trimRange(grid: Scalar[][]): Scalar[][] {
  const empty = (v: Scalar) => v === null || v === '';
  let top = 0;
  let bottom = grid.length - 1;
  while (top <= bottom && grid[top].every(empty)) top++;
  while (bottom >= top && grid[bottom].every(empty)) bottom--;
  if (top > bottom) return [['']];
  const rows = grid.slice(top, bottom + 1);
  const width = Math.max(...rows.map((r) => r.length));
  let left = 0;
  let right = width - 1;
  while (left <= right && rows.every((r) => empty(r[left] ?? null))) left++;
  while (right >= left && rows.every((r) => empty(r[right] ?? null))) right--;
  return rows.map((r) => r.slice(left, right + 1).map((v) => v ?? ''));
}

function scalar(v: BaseValueObject): Scalar {
  if (v.isNull()) return null;
  return v.getValue() as Scalar;
}

class IsoCeiling extends BaseFunction {
  override minParams = 1;
  override maxParams = 2;

  override calculate(number: BaseValueObject, significance?: BaseValueObject) {
    if (number.isError()) return number;
    if (significance?.isError()) return significance;
    const x = Number(scalar(number) ?? 0);
    const s = significance ? Number(scalar(significance) ?? 1) : 1;
    if (!Number.isFinite(x) || !Number.isFinite(s)) return ErrorValueObject.create('#VALUE!' as never);
    return NumberValueObject.create(isoCeiling(x, s));
  }
}

class TrimRange extends BaseFunction {
  override minParams = 1;
  override maxParams = 3;
  override needsReferenceObject = false;

  override calculate(range: BaseValueObject) {
    if (range.isError()) return range;
    if (!range.isArray()) return range;
    const grid = (range as ArrayValueObject).getArrayValue().map((row) => row.map((c) => (c ? scalar(c) : null)));
    return ArrayValueObject.createByArray(trimRange(grid));
  }
}

class AmorDegrc extends BaseFunction {
  override minParams = 6;
  override maxParams = 7;

  override calculate(...args: BaseValueObject[]) {
    const bad = args.find((a) => a.isError());
    if (bad) return bad;
    const n = args.map((a) => Number(scalar(a) ?? 0));
    const [cost, purchased, first, salvage, period, rate, basis] = n as [number, number, number, number, number, number, number?];
    if (n.some((x) => !Number.isFinite(x))) return ErrorValueObject.create('#VALUE!' as never);
    if (rate <= 0 || cost < salvage || period < 0) return ErrorValueObject.create('#NUM!' as never);
    return NumberValueObject.create(amorDegrc(cost, purchased, first, salvage, period, rate, basis ?? 0));
  }
}

// --- Fonctions du moteur corrigées ---
// Le moteur enregistre les fonctions de CUSTOM_FUNCTIONS après les siennes : un même nom remplace la sienne. À revérifier à
// chaque mise à jour d'Univer (classes internes atteintes par les listes functionText, functionLookup… ; test : excel-reference).

type Fn = new (name: IFunctionNames) => BaseFunction;
type Args = BaseValueObject[];
const builtin = (list: unknown, name: string): Fn => {
  const entry = (list as [Fn, string][]).find(([, n]) => n === name);
  if (!entry) throw new Error(`Fonction du moteur introuvable : ${name}`);
  return entry[0];
};
type Calc = { calculate(...a: Args): BaseValueObject };
/** Sous-classe d'une fonction du moteur dont les arguments sont d'abord transformés. */
function withArgs(list: unknown, name: string, map: (args: Args) => Args): [typeof BaseFunction, IFunctionNames] {
  const Base = builtin(list, name) as unknown as new (n: IFunctionNames) => Calc;
  class Fixed extends Base {
    override calculate(...args: Args) {
      return super.calculate(...map(args));
    }
  }
  return [Fixed as unknown as typeof BaseFunction, name as IFunctionNames];
}
/** Applique `fn` à une valeur simple, ou à chaque valeur d'une plage. */
function mapValue(v: BaseValueObject, fn: (x: BaseValueObject) => BaseValueObject): BaseValueObject {
  // Certaines fonctions (NB.SI, SOMME.SI…) reçoivent une référence et non une valeur : laissée telle quelle.
  if ((v as unknown as { isReferenceObject(): boolean }).isReferenceObject()) return v;
  return v.isArray() ? (v as ArrayValueObject).map((c) => fn(c as BaseValueObject)) : fn(v);
}

/** Nombre ou logique → texte à la française (« &a&1,5&VRAI »). */
const frenchText = (v: BaseValueObject) => mapValue(v, (x) => {
  if (x.isBoolean()) return StringValueObject.create(x.getValue() ? 'VRAI' : 'FAUX');
  if (x.isNumber()) return StringValueObject.create(frenchNumberText(Number(x.getValue())));
  return x;
});
/** Critère de NB.SI, SOMME.SI… écrit à la française. */
const frenchCriterion = (v: BaseValueObject) => mapValue(v, (x) => (x.isString() ? StringValueObject.create(frenchCriteria(String(x.getValue()))) : x));
/** Plages de MOYENNEA, ECARTYPEA… : un texte compte pour 0, même s'il ressemble à un nombre. */
const textIsZero = (v: BaseValueObject) => (v.isArray() ? mapValue(v, (x) => (x.isString() ? NumberValueObject.create(0) : x)) : v);
/** HEURE, MINUTE, SECONDE : Excel arrondit d'abord à la seconde la plus proche. */
const roundedSecond = (v: BaseValueObject) => mapValue(v, (x) => (x.isNumber() ? NumberValueObject.create(Math.round(Number(x.getValue()) * 86_400) / 86_400) : x));

/** Le moteur trie la plage elle-même : TRIER faussait toutes les autres formules qui lisent cette plage. On trie une copie. */
const SortBase = builtin(functionLookup, 'SORT') as unknown as new (n: IFunctionNames) => { _getResult(array: ArrayValueObject, ...rest: unknown[]): BaseValueObject };
class SafeSort extends SortBase {
  override _getResult(array: ArrayValueObject, ...rest: unknown[]) {
    const copy = ArrayValueObject.create({ calculateValueList: array.getArrayValue().map((row) => [...row]), rowCount: array.getRowCount(), columnCount: array.getColumnCount(), unitId: '', sheetId: '', row: -1, column: -1 });
    return super._getResult(copy, ...rest);
  }
}

/** Valeur d'une formule TEXTE, avec les codes français et les noms de mois en français. */
function frenchTextFormat(value: BaseValueObject, format: string): BaseValueObject {
  if (value.isError()) return value;
  if (value.isBoolean()) return StringValueObject.create(value.getValue() ? 'VRAI' : 'FAUX');
  let v: number | string = value.isNull() ? 0 : (value.getValue() as number | string);
  if (typeof v === 'string') v = parseFrenchNumber(v) ?? parseFrenchDate(v) ?? v;
  try {
    return StringValueObject.create(numfmt.format(frenchFormatToEnglish(format), v, { locale: 'fr' }));
  } catch {
    return ErrorValueObject.create('#VALUE!' as never);
  }
}
class FrenchTextFn extends BaseFunction {
  override minParams = 2;
  override maxParams = 2;
  override isArgumentsIgnoreNumberPattern() {
    return true;
  }
  override calculate(value: BaseValueObject, format: BaseValueObject) {
    if (value.isError()) return value;
    if (format.isError()) return format;
    const code = format.isArray() ? (format as ArrayValueObject).get(0, 0) : format;
    if (code?.isBoolean()) return ErrorValueObject.create('#VALUE!' as never);
    const text = String(code?.getValue() ?? '');
    const result = mapValue(value, (x) => frenchTextFormat(x, text));
    if (!result.isArray()) return result;
    const array = result as ArrayValueObject;
    return array.getRowCount() === 1 && array.getColumnCount() === 1 ? (array.get(0, 0) as BaseValueObject) : array;
  }
}

/** CNUM : « 12,5 », « 1 000 », « 12,5 % », dates et heures à la française. */
const ValueBase = builtin(functionText, 'VALUE') as unknown as new (n: IFunctionNames) => Calc;
class FrenchValue extends ValueBase {
  override calculate(...args: Args) {
    return mapValue(args[0]!, (x) => {
      if (x.isString()) {
        const text = String(x.getValue());
        const n = parseFrenchNumber(text) ?? parseFrenchDate(text);
        if (n !== null) return NumberValueObject.create(n);
      }
      return super.calculate(x);
    });
  }
}
const DateValueBase = builtin(functionDate, 'DATEVALUE') as unknown as new (n: IFunctionNames) => Calc;
class FrenchDateValue extends DateValueBase {
  override calculate(...args: Args) {
    return mapValue(args[0]!, (x) => {
      if (x.isString()) {
        const n = parseFrenchDate(String(x.getValue()));
        if (n !== null) return NumberValueObject.create(n);
      }
      return super.calculate(x);
    });
  }
}

/** CTXT et EURO : séparateurs français, symbole €. */
class FrenchFixed extends BaseFunction {
  override minParams = 1;
  override maxParams = 3;
  override calculate(number: BaseValueObject, decimals?: BaseValueObject, noCommas?: BaseValueObject) {
    if (number.isError()) return number;
    if (decimals?.isError()) return decimals;
    const d = decimals && !decimals.isNull() ? Math.trunc(Number(scalar(decimals) ?? 2)) : 2;
    const raw = Number(scalar(number) ?? 0);
    if (!Number.isFinite(raw)) return ErrorValueObject.create('#VALUE!' as never);
    return StringValueObject.create(fixedText(raw, d, !!noCommas && !!scalar(noCommas)));
  }
}
class FrenchDollar extends BaseFunction {
  override minParams = 1;
  override maxParams = 2;
  override calculate(number: BaseValueObject, decimals?: BaseValueObject) {
    if (number.isError()) return number;
    if (decimals?.isError()) return decimals;
    const d = decimals && !decimals.isNull() ? Math.trunc(Number(scalar(decimals) ?? 2)) : 2;
    const raw = Number(scalar(number) ?? 0);
    if (!Number.isFinite(raw)) return ErrorValueObject.create('#VALUE!' as never);
    return StringValueObject.create(`${fixedText(raw, d)}\u00a0€`);
  }
}

/** CHERCHE avec les jokers ? et *. */
const SearchBase = builtin(functionText, 'SEARCH') as unknown as new (n: IFunctionNames) => { _handleSingleObject(find: BaseValueObject, within: BaseValueObject, start: BaseValueObject): BaseValueObject };
class WildcardSearch extends SearchBase {
  override _handleSingleObject(find: BaseValueObject, within: BaseValueObject, start: BaseValueObject) {
    const text = String(find.getValue() ?? '');
    if (!/[?*~]/.test(text) || within.isNull()) return super._handleSingleObject(find, within, start);
    const pos = searchWithWildcards(text, String(within.getValue()), Math.floor(Number(start.getValue()) || 1));
    return pos ? NumberValueObject.create(pos) : ErrorValueObject.create('#VALUE!' as never);
  }
}

/** CHOISIR({1;2};…) (recherche vers la gauche) : le moteur ne répète pas l'indice sur toutes les lignes de la plage. */
const ChooseBase = builtin(functionLookup, 'CHOOSE') as unknown as new (n: IFunctionNames) => Calc;
class RepeatingChoose extends ChooseBase {
  override calculate(index: BaseValueObject, ...values: Args) {
    if (!index.isArray()) return super.calculate(index, ...values);
    const idx = index as ArrayValueObject;
    const sizes = values.map((v) => {
      const a = (v as unknown as { isReferenceObject(): boolean }).isReferenceObject() ? (v as unknown as { toArrayValueObject(): ArrayValueObject }).toArrayValueObject() : (v as ArrayValueObject);
      return v.isArray() || a !== (v as unknown) ? [a.getRowCount(), a.getColumnCount()] : [1, 1];
    });
    const rows = Math.max(idx.getRowCount(), ...sizes.map((x) => x[0]!));
    const cols = Math.max(idx.getColumnCount(), ...sizes.map((x) => x[1]!));
    if ((idx.getRowCount() !== 1 && idx.getColumnCount() !== 1) || (rows === idx.getRowCount() && cols === idx.getColumnCount())) return super.calculate(index, ...values);
    const list = Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_c, c) => idx.get(idx.getRowCount() === 1 ? 0 : r, idx.getColumnCount() === 1 ? 0 : c)));
    return super.calculate(ArrayValueObject.create({ calculateValueList: list as never, rowCount: rows, columnCount: cols, unitId: '', sheetId: '', row: -1, column: -1 }), ...values);
  }
}

/** FORMULETEXTE : la formule telle qu'on la voit, en français. */
const FormulaTextBase = builtin(functionLookup, 'FORMULATEXT') as unknown as new (n: IFunctionNames) => Calc;
class FrenchFormulaText extends FormulaTextBase {
  override calculate(...args: Args) {
    return mapValue(super.calculate(...args), (x) => (x.isString() ? StringValueObject.create(englishToFrench(String(x.getValue()))) : x));
  }
}

// Formats automatiques du moteur : à l'américaine (2024/03/01, 2:30 PM, (€106,07)). Excel en français : 01/03/2024, 14:30, -106,07 €.
const AUTO_PATTERNS: Record<string, string> = { 'yyyy/mm/dd;@': 'dd/mm/yyyy;@', 'yyyy/mm/dd hh:mm': 'dd/mm/yyyy hh:mm', 'h:mm A/P': 'hh:mm' };
const CURRENCY = /^"€"#,##0(\.0+)?_\);\[Red\]\("€"#,##0(\.0+)?\)$/;
const createNumber = NumberValueObject.create.bind(NumberValueObject);
NumberValueObject.create = (value: number, pattern = '') => {
  const money = CURRENCY.exec(pattern);
  return createNumber(value, money ? `#,##0${money[1] ?? ''} "€";[Red]-#,##0${money[1] ?? ''} "€"` : (AUTO_PATTERNS[pattern] ?? pattern));
};

// Texte converti en nombre dans un calcul (« "1,5"+1 ») : à la française aussi.
const toNumber = StringValueObject.prototype.convertToNumberObjectValue;
StringValueObject.prototype.convertToNumberObjectValue = function (this: StringValueObject) {
  const text = String(this.getValue());
  if (/[,\s\u00a0\u202f]/.test(text)) {
    const n = parseFrenchNumber(text);
    if (n !== null) return NumberValueObject.create(n);
  }
  return toNumber.call(this);
};

const CRITERIA_AT: [string, number[]][] = [['COUNTIF', [1]], ['COUNTIFS', [1, 3, 5, 7, 9, 11, 13, 15]], ['SUMIF', [1]], ['SUMIFS', [2, 4, 6, 8, 10, 12, 14]], ['AVERAGEIF', [1]], ['AVERAGEIFS', [2, 4, 6, 8, 10, 12, 14]], ['MAXIFS', [2, 4, 6, 8, 10, 12, 14]], ['MINIFS', [2, 4, 6, 8, 10, 12, 14]]];
const allFunctions = [...(functionStatistical as unknown as unknown[]), ...(functionMath as unknown as unknown[])];

export const CUSTOM_FUNCTIONS: Array<[typeof BaseFunction, IFunctionNames]> = [
  [IsoCeiling as unknown as typeof BaseFunction, 'ISO.CEILING' as IFunctionNames],
  [TrimRange as unknown as typeof BaseFunction, 'TRIMRANGE' as IFunctionNames],
  [AmorDegrc as unknown as typeof BaseFunction, 'AMORDEGRC' as IFunctionNames],
  [SafeSort as unknown as typeof BaseFunction, 'SORT' as IFunctionNames],
  [FrenchTextFn as unknown as typeof BaseFunction, 'TEXT' as IFunctionNames],
  [FrenchValue as unknown as typeof BaseFunction, 'VALUE' as IFunctionNames],
  [FrenchDateValue as unknown as typeof BaseFunction, 'DATEVALUE' as IFunctionNames],
  [FrenchFixed as unknown as typeof BaseFunction, 'FIXED' as IFunctionNames],
  [FrenchDollar as unknown as typeof BaseFunction, 'DOLLAR' as IFunctionNames],
  [WildcardSearch as unknown as typeof BaseFunction, 'SEARCH' as IFunctionNames],
  [RepeatingChoose as unknown as typeof BaseFunction, 'CHOOSE' as IFunctionNames],
  [FrenchFormulaText as unknown as typeof BaseFunction, 'FORMULATEXT' as IFunctionNames],
  ...['CONCATENATE', 'CONCAT', 'TEXTJOIN'].map((name) => withArgs(functionText, name, name === 'TEXTJOIN' ? (a) => [a[0]!, a[1]!, ...a.slice(2).map(frenchText)] : (a) => a.map(frenchText))),
  ...CRITERIA_AT.map(([name, at]) => withArgs(allFunctions, name, (a) => a.map((v, i) => (at.includes(i) ? frenchCriterion(v) : v)))),
  ...['AVERAGEA', 'STDEVA', 'STDEVPA', 'VARA', 'VARPA'].map((name) => withArgs(functionStatistical, name, (a) => a.map(textIsZero))),
  ...['HOUR', 'MINUTE', 'SECOND'].map((name) => withArgs(functionDate, name, (a) => a.map(roundedSecond))),
];

/** Descriptions pour l'autocomplétion et l'aide sur les arguments (en français). */
export const CUSTOM_FUNCTION_DESCRIPTIONS: IFunctionInfo[] = [
  {
    functionName: 'ISO.CEILING',
    functionType: FunctionType.Math,
    description: 'Arrondit un nombre à l’excès au multiple de la précision le plus proche, quel que soit son signe.',
    abstract: 'Arrondit à l’excès (norme ISO)',
    functionParameter: [
      { name: 'nombre', detail: 'Valeur à arrondir.', example: '4,3', require: 1, repeat: 0 },
      { name: 'précision', detail: 'Multiple auquel arrondir (1 par défaut).', example: '2', require: 0, repeat: 0 },
    ],
  },
  {
    functionName: 'TRIMRANGE',
    functionType: FunctionType.Lookup,
    description: 'Renvoie la plage sans les lignes et colonnes vides en bordure.',
    abstract: 'Retire les bords vides d’une plage',
    functionParameter: [{ name: 'plage', detail: 'Plage à réduire.', example: 'A1:D100', require: 1, repeat: 0 }],
  },
] as IFunctionInfo[];

/**
 * Fonctions Excel absentes du moteur Univer, ajoutées au moteur de formules.
 *
 * Le calcul a lieu dans le Web Worker : les classes de fonction sont passées à la config
 * du preset côté worker (calcul) ET côté page (analyse des formules, autocomplétion).
 * La logique pure (isoCeiling, trimRange) est testée à part.
 */
import {
  ArrayValueObject,
  BaseFunction,
  ErrorValueObject,
  FunctionType,
  NumberValueObject,
  type BaseValueObject,
  type IFunctionInfo,
  type IFunctionNames,
} from '@univerjs/engine-formula';

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

export const CUSTOM_FUNCTIONS: Array<[typeof BaseFunction, IFunctionNames]> = [
  [IsoCeiling as unknown as typeof BaseFunction, 'ISO.CEILING' as IFunctionNames],
  [TrimRange as unknown as typeof BaseFunction, 'TRIMRANGE' as IFunctionNames],
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

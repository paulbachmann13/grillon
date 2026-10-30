/**
 * Fonctions récentes d'Excel dans le fichier .xlsx. Le fichier écrit « _xlfn.IFS(...) », « _xlfn.STDEV.S(...) »,
 * « _xlfn._xlws.SORT(...) », et les paramètres de LET et LAMBDA « _xlpm.x ». Sans ces préfixes, Excel affiche #NOM? (voire répare
 * le fichier pour SORT et FILTER). Grillon garde en mémoire la forme sans préfixe : ajoutés à l'écriture, retirés à la lecture.
 * Liste relevée dans Excel (tests/fixtures/excel-reference.ts) et dans la documentation de Microsoft (MS-XLSX, « future functions »).
 */
const FUTURE = new Set(
  `ACOT ACOTH AGGREGATE ARABIC ARRAYTOTEXT BASE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV BITAND BITLSHIFT BITOR BITRSHIFT BITXOR BYCOL BYROW
  CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST CHOOSECOLS CHOOSEROWS COMBINA CONCAT CONFIDENCE.NORM CONFIDENCE.T COT COTH
  COVARIANCE.P COVARIANCE.S CSC CSCH DAYS DECIMAL DROP ECMA.CEILING ERF.PRECISE ERFC.PRECISE EXPAND EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FILTERXML FLOOR.MATH
  FLOOR.PRECISE FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR FORMULATEXT GAMMA GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS
  GROUPBY HSTACK HYPGEOM.DIST IFNA IFS IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN ISFORMULA ISO.CEILING ISOMITTED ISOWEEKNUM LAMBDA LET LOGNORM.DIST LOGNORM.INV MAKEARRAY MAP
  MAXIFS MINIFS MODE.MULT MODE.SNGL MUNIT NEGBINOM.DIST NETWORKDAYS.INTL NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC
  PERCENTRANK.EXC PERCENTRANK.INC PERMUTATIONA PHI PIVOTBY POISSON.DIST QUARTILE.EXC QUARTILE.INC QUERYSTRING RANDARRAY RANK.AVG RANK.EQ REDUCE REGEXEXTRACT REGEXREPLACE REGEXTEST RRI SCAN SEC SECH SEQUENCE SHEET SHEETS SINGLE
  SKEW.P SORTBY STDEV.P STDEV.S SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TAKE TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TOCOL TOROW TRIMRANGE UNICHAR UNICODE UNIQUE
  VALUETOTEXT VAR.P VAR.S VSTACK WEBSERVICE WEIBULL.DIST WORKDAY.INTL WRAPCOLS WRAPROWS XLOOKUP XMATCH XOR Z.TEST`.split(/\s+/),
);
/** Ces deux noms existaient autrefois dans une autre feuille de macros : double préfixe. */
const WORKSHEET_FUTURE = new Set(['SORT', 'FILTER']);

/** Fonctions dont le résultat est un tableau : leur formule doit être marquée « matricielle dynamique » dans le fichier. */
const ARRAY_RESULT = new Set(
  `SORT SORTBY FILTER UNIQUE SEQUENCE RANDARRAY TEXTSPLIT HSTACK VSTACK TAKE DROP EXPAND CHOOSECOLS CHOOSEROWS TOCOL TOROW WRAPROWS WRAPCOLS MAP BYROW BYCOL SCAN MAKEARRAY GROUPBY PIVOTBY
  TRANSPOSE MMULT MINVERSE MUNIT FREQUENCY LINEST LOGEST TREND GROWTH MODE.MULT`.split(/\s+/),
);

interface Part {
  text: string;
  string: boolean;
}

/** Sépare une formule en morceaux : chaînes "…" et noms de feuille '…' à part (on n'y touche jamais). */
function split(formula: string): Part[] {
  const out: Part[] = [];
  let i = 0;
  let start = 0;
  while (i < formula.length) {
    const q = formula[i];
    if (q === '"' || q === "'") {
      if (i > start) out.push({ text: formula.slice(start, i), string: false });
      let j = i + 1;
      while (j < formula.length) {
        if (formula[j] === q) {
          if (formula[j + 1] === q) j += 2;
          else break;
        } else j++;
      }
      out.push({ text: formula.slice(i, j + 1), string: true });
      i = j + 1;
      start = i;
    } else i++;
  }
  if (start < formula.length) out.push({ text: formula.slice(start), string: false });
  return out;
}

const FUNCTION_NAME = /(^|[^A-Za-z0-9_.$!])([A-Za-z][A-Za-z0-9_.]*)(?=\()/g;

/** Noms des paramètres de LET et LAMBDA présents dans la formule. */
function lambdaParams(formula: string): string[] {
  const names = new Set<string>();
  const parts = split(formula);
  const masked = parts.map((p) => (p.string ? ' '.repeat(p.text.length) : p.text)).join('');
  for (const m of masked.matchAll(/(^|[^A-Za-z0-9_.])(LET|LAMBDA)\(/gi)) {
    const open = m.index! + m[0].length;
    const args: string[] = [];
    let depth = 0;
    let start = open;
    for (let i = open; i < masked.length; i++) {
      const ch = masked[i];
      if (ch === '(' || ch === '{') depth++;
      else if (ch === ')' || ch === '}') {
        if (depth === 0) {
          args.push(masked.slice(start, i));
          break;
        }
        depth--;
      } else if (ch === ',' && depth === 0) {
        args.push(masked.slice(start, i));
        start = i + 1;
      }
    }
    const isLet = m[2]!.toUpperCase() === 'LET';
    args.forEach((a, k) => {
      const name = a.trim();
      if ((isLet ? k % 2 === 0 && k < args.length - 1 : k < args.length - 1) && /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name)) names.add(name);
    });
  }
  return [...names];
}

/** Formule en mémoire (sans « = ») → forme du fichier. */
export function toFileFormula(formula: string): string {
  const params = lambdaParams(formula);
  const paramRe = params.length ? new RegExp(`(^|[^A-Za-z0-9_.$!])(${params.map((p) => p.replace(/\./g, '\\.')).join('|')})(?![A-Za-z0-9_.(])`, 'g') : null;
  return split(formula)
    .map((p) => {
      if (p.string) return p.text;
      let t = p.text.replace(FUNCTION_NAME, (m, before: string, name: string) => {
        const upper = name.toUpperCase();
        if (WORKSHEET_FUTURE.has(upper)) return `${before}_xlfn._xlws.${upper}`;
        return FUTURE.has(upper) ? `${before}_xlfn.${upper}` : m;
      });
      if (paramRe) t = t.replace(paramRe, (_m, before: string, name: string) => `${before}_xlpm.${name}`);
      return t;
    })
    .join('');
}

/** Forme du fichier → formule en mémoire. */
export const fromFileFormula = (formula: string): string => formula.replace(/_xlfn\.|_xlws\.|_xlpm\./g, '');

/** La formule doit-elle être marquée matricielle dans le fichier (sans cela Excel n'en calcule qu'une valeur, ou #VALEUR!) ? */
export function needsArrayFlag(formula: string): boolean {
  const code = split(formula).map((p) => (p.string ? (p.text[0] === "'" ? 'S' : '""') : p.text)).join('');
  for (const m of code.matchAll(FUNCTION_NAME)) if (ARRAY_RESULT.has(m[2]!.toUpperCase())) return true;
  if (code.includes('{')) return true;
  // Plage dans un calcul : (A1:A5="x")*B1:B5, A1:A5+1…
  const range = String.raw`(?:[A-Za-z0-9_.]+!)?\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+`;
  return new RegExp(`(?:[=<>+\\-*/^&]\\s*${range})|(?:${range}\\s*[=<>+\\-*/^&%])`).test(code);
}

const unescapeXml = (s: string) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const escapeXml = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/**
 * XML d'une feuille : formules de cellules, de mise en forme conditionnelle et de validation dans la forme du fichier ; les
 * formules matricielles sont marquées (cm="1" renvoie à xl/metadata.xml, t="array" à la plage). Renvoie aussi si une l'a été.
 */
export function rewriteSheetFormulas(xml: string): { xml: string; hasArray: boolean } {
  let hasArray = false;
  let out = xml.replace(/<c r="([A-Z]+\d+)"([^>]*)><f>([^<]*)<\/f>/g, (_m, ref: string, attrs: string, text: string) => {
    const formula = unescapeXml(text);
    const file = escapeXml(toFileFormula(formula));
    if (!needsArrayFlag(formula)) return `<c r="${ref}"${attrs}><f>${file}</f>`;
    hasArray = true;
    return `<c r="${ref}"${attrs} cm="1"><f t="array" ref="${ref}">${file}</f>`;
  });
  out = out.replace(/<(formula[12]?)>([^<]*)<\/\1>/g, (_m, tag: string, text: string) => `<${tag}>${escapeXml(toFileFormula(unescapeXml(text)))}</${tag}>`);
  return { xml: out, hasArray };
}

/** Partie xl/metadata.xml : déclare le type « matrice dynamique » auquel renvoie cm="1". */
export const DYNAMIC_ARRAY_METADATA =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">' +
  '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>' +
  '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>' +
  '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>';
export const CT_SHEET_METADATA = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml';
export const REL_SHEET_METADATA = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata';

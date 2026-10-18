/**
 * Formules « à la française », comme Excel en français : noms de fonctions français (SOMME, SI,
 * RECHERCHEV…), « ; » entre les arguments, virgule décimale, VRAI / FAUX.
 *
 * Le classeur garde toujours la forme anglaise (c'est celle du fichier .xlsx et du moteur de
 * calcul) : on traduit à la saisie (français → anglais) et à l'affichage dans la barre de
 * formule (anglais → français). Logique pure, sans DOM ni Univer.
 */

/** [anglais, français] ; les noms identiques dans les deux langues ne sont pas listés. */
const PAIRS: [string, string][] = [
  // Maths et trigonométrie
  ['AGGREGATE', 'AGREGAT'], ['ARABIC', 'CHIFFRE.ARABE'], ['CEILING', 'PLAFOND'], ['CEILING.MATH', 'PLAFOND.MATH'], ['CEILING.PRECISE', 'PLAFOND.PRECIS'],
  ['DEGREES', 'DEGRES'], ['EVEN', 'PAIR'], ['FLOOR', 'PLANCHER'], ['FLOOR.MATH', 'PLANCHER.MATH'], ['FLOOR.PRECISE', 'PLANCHER.PRECIS'], ['GCD', 'PGCD'],
  ['INT', 'ENT'], ['ISO.CEILING', 'ISO.PLAFOND'], ['LCM', 'PPCM'], ['MDETERM', 'DETERMAT'], ['MINVERSE', 'INVERSEMAT'], ['MMULT', 'PRODUITMAT'],
  ['MROUND', 'ARRONDI.AU.MULTIPLE'], ['MULTINOMIAL', 'MULTINOMIALE'], ['MUNIT', 'MATRICE.UNITAIRE'], ['ODD', 'IMPAIR'], ['POWER', 'PUISSANCE'],
  ['PRODUCT', 'PRODUIT'], ['RAND', 'ALEA'], ['RANDARRAY', 'TABLEAU.ALEA'], ['RANDBETWEEN', 'ALEA.ENTRE.BORNES'], ['ROMAN', 'ROMAIN'], ['ROUND', 'ARRONDI'],
  ['ROUNDDOWN', 'ARRONDI.INF'], ['ROUNDUP', 'ARRONDI.SUP'], ['SERIESSUM', 'SOMME.SERIES'], ['SIGN', 'SIGNE'], ['SQRT', 'RACINE'], ['SQRTPI', 'RACINE.PI'],
  ['SUBTOTAL', 'SOUS.TOTAL'], ['SUM', 'SOMME'], ['SUMIF', 'SOMME.SI'], ['SUMIFS', 'SOMME.SI.ENS'], ['SUMPRODUCT', 'SOMMEPROD'], ['SUMSQ', 'SOMME.CARRES'],
  ['SUMX2MY2', 'SOMME.X2MY2'], ['SUMX2PY2', 'SOMME.X2PY2'], ['SUMXMY2', 'SOMME.XMY2'], ['TRUNC', 'TRONQUE'],
  // Logique
  ['AND', 'ET'], ['FALSE', 'FAUX'], ['IF', 'SI'], ['IFERROR', 'SIERREUR'], ['IFNA', 'SI.NON.DISP'], ['IFS', 'SI.CONDITIONS'], ['NOT', 'NON'], ['OR', 'OU'],
  ['SWITCH', 'SI.MULTIPLE'], ['TRUE', 'VRAI'], ['XOR', 'OUX'], ['BYCOL', 'PARCOL'], ['BYROW', 'PARLIGNE'],
  // Statistiques
  ['AVEDEV', 'ECART.MOYEN'], ['AVERAGE', 'MOYENNE'], ['AVERAGEIF', 'MOYENNE.SI'], ['AVERAGEIFS', 'MOYENNE.SI.ENS'], ['BETA.DIST', 'LOI.BETA.N'],
  ['BETA.INV', 'BETA.INVERSE.N'], ['BINOM.DIST', 'LOI.BINOMIALE.N'], ['BINOM.DIST.RANGE', 'LOI.BINOMIALE.SERIE'], ['BINOM.INV', 'LOI.BINOMIALE.INVERSE'],
  ['CHISQ.DIST', 'LOI.KHIDEUX.N'], ['CHISQ.DIST.RT', 'LOI.KHIDEUX.DROITE'], ['CHISQ.INV', 'LOI.KHIDEUX.INVERSE'], ['CHISQ.INV.RT', 'LOI.KHIDEUX.INVERSE.DROITE'],
  ['CHISQ.TEST', 'TEST.KHIDEUX'], ['CONFIDENCE.NORM', 'INTERVALLE.CONFIANCE.NORMAL'], ['CONFIDENCE.T', 'INTERVALLE.CONFIANCE.STUDENT'],
  ['CORREL', 'COEFFICIENT.CORRELATION'], ['COUNT', 'NB'], ['COUNTA', 'NBVAL'], ['COUNTBLANK', 'NB.VIDE'], ['COUNTIF', 'NB.SI'], ['COUNTIFS', 'NB.SI.ENS'],
  ['COVARIANCE.P', 'COVARIANCE.PEARSON'], ['COVARIANCE.S', 'COVARIANCE.STANDARD'], ['DEVSQ', 'SOMME.CARRES.ECARTS'], ['EXPON.DIST', 'LOI.EXPONENTIELLE.N'],
  ['F.DIST', 'LOI.F.N'], ['F.DIST.RT', 'LOI.F.DROITE'], ['F.INV', 'INVERSE.LOI.F.N'], ['F.INV.RT', 'INVERSE.LOI.F.DROITE'], ['F.TEST', 'TEST.F'],
  ['FISHERINV', 'FISHER.INVERSE'], ['FORECAST', 'PREVISION'], ['FORECAST.LINEAR', 'PREVISION.LINEAIRE'], ['FREQUENCY', 'FREQUENCE'],
  ['GAMMA.DIST', 'LOI.GAMMA.N'], ['GAMMA.INV', 'LOI.GAMMA.INVERSE.N'], ['GAMMALN', 'LNGAMMA'], ['GAMMALN.PRECISE', 'LNGAMMA.PRECIS'],
  ['GEOMEAN', 'MOYENNE.GEOMETRIQUE'], ['GROWTH', 'CROISSANCE'], ['HARMEAN', 'MOYENNE.HARMONIQUE'], ['HYPGEOM.DIST', 'LOI.HYPERGEOMETRIQUE.N'],
  ['INTERCEPT', 'ORDONNEE.ORIGINE'], ['KURT', 'KURTOSIS'], ['LARGE', 'GRANDE.VALEUR'], ['LINEST', 'DROITEREG'], ['LOGEST', 'LOGREG'],
  ['LOGNORM.DIST', 'LOI.LOGNORMALE.N'], ['LOGNORM.INV', 'LOI.LOGNORMALE.INVERSE.N'], ['MAXIFS', 'MAX.SI.ENS'], ['MEDIAN', 'MEDIANE'], ['MINIFS', 'MIN.SI.ENS'],
  ['MODE.MULT', 'MODE.MULTIPLE'], ['MODE.SNGL', 'MODE.SIMPLE'], ['NEGBINOM.DIST', 'LOI.BINOMIALE.NEG.N'], ['NORM.DIST', 'LOI.NORMALE.N'],
  ['NORM.INV', 'LOI.NORMALE.INVERSE.N'], ['NORM.S.DIST', 'LOI.NORMALE.STANDARD.N'], ['NORM.S.INV', 'LOI.NORMALE.STANDARD.INVERSE.N'],
  ['PERCENTILE', 'CENTILE'], ['PERCENTILE.EXC', 'CENTILE.EXCLURE'], ['PERCENTILE.INC', 'CENTILE.INCLURE'], ['PERCENTRANK', 'RANG.POURCENTAGE'],
  ['PERCENTRANK.EXC', 'RANG.POURCENTAGE.EXCLURE'], ['PERCENTRANK.INC', 'RANG.POURCENTAGE.INCLURE'], ['PERMUT', 'PERMUTATION'],
  ['POISSON.DIST', 'LOI.POISSON.N'], ['PROB', 'PROBABILITE'], ['QUARTILE.EXC', 'QUARTILE.EXCLURE'], ['QUARTILE.INC', 'QUARTILE.INCLURE'], ['RANK', 'RANG'],
  ['RANK.AVG', 'MOYENNE.RANG'], ['RANK.EQ', 'RANG.EGALITE'], ['RSQ', 'COEFFICIENT.DETERMINATION'], ['SKEW', 'COEFFICIENT.ASYMETRIE'],
  ['SKEW.P', 'COEFFICIENT.ASYMETRIE.P'], ['SLOPE', 'PENTE'], ['SMALL', 'PETITE.VALEUR'], ['STANDARDIZE', 'CENTREE.REDUITE'], ['STDEV', 'ECARTYPE'],
  ['STDEV.P', 'ECARTYPE.PEARSON'], ['STDEV.S', 'ECARTYPE.STANDARD'], ['STDEVP', 'ECARTYPEP'], ['STEYX', 'ERREUR.TYPE.XY'], ['T.DIST', 'LOI.STUDENT.N'],
  ['T.DIST.2T', 'LOI.STUDENT.BILATERALE'], ['T.DIST.RT', 'LOI.STUDENT.DROITE'], ['T.INV', 'LOI.STUDENT.INVERSE.N'], ['T.INV.2T', 'LOI.STUDENT.INVERSE.BILATERALE'],
  ['T.TEST', 'TEST.STUDENT'], ['TREND', 'TENDANCE'], ['TRIMMEAN', 'MOYENNE.REDUITE'], ['VAR.P', 'VAR.P.N'], ['VARP', 'VAR.P'], ['WEIBULL.DIST', 'LOI.WEIBULL.N'],
  ['Z.TEST', 'TEST.Z'],
  // Fonctions de compatibilité
  ['BETADIST', 'LOI.BETA'], ['BINOMDIST', 'LOI.BINOMIALE'], ['CHIDIST', 'LOI.KHIDEUX'], ['CHIINV', 'KHIDEUX.INVERSE'], ['CHITEST', 'TEST.KHIDEUX'],
  ['CONFIDENCE', 'INTERVALLE.CONFIANCE'], ['COVAR', 'COVARIANCE'], ['CRITBINOM', 'CRITERE.LOI.BINOMIALE'], ['EXPONDIST', 'LOI.EXPONENTIELLE'],
  ['FDIST', 'LOI.F'], ['FINV', 'INVERSE.LOI.F'], ['FTEST', 'TEST.F'], ['GAMMADIST', 'LOI.GAMMA'], ['GAMMAINV', 'LOI.GAMMA.INVERSE'],
  ['HYPGEOMDIST', 'LOI.HYPERGEOMETRIQUE'], ['LOGINV', 'LOI.LOGNORMALE.INVERSE'], ['LOGNORMDIST', 'LOI.LOGNORMALE'], ['NEGBINOMDIST', 'LOI.BINOMIALE.NEG'],
  ['NORMDIST', 'LOI.NORMALE'], ['NORMINV', 'LOI.NORMALE.INVERSE'], ['NORMSDIST', 'LOI.NORMALE.STANDARD'], ['NORMSINV', 'LOI.NORMALE.STANDARD.INVERSE'],
  ['POISSON', 'LOI.POISSON'], ['TDIST', 'LOI.STUDENT'], ['TINV', 'LOI.STUDENT.INVERSE'], ['TTEST', 'TEST.STUDENT'], ['WEIBULL', 'LOI.WEIBULL'], ['ZTEST', 'TEST.Z'],
  // Texte
  ['CHAR', 'CAR'], ['CLEAN', 'EPURAGE'], ['CONCATENATE', 'CONCATENER'], ['DOLLAR', 'DEVISE'], ['FIND', 'TROUVE'], ['FINDB', 'TROUVERB'], ['FIXED', 'CTXT'],
  ['LEFT', 'GAUCHE'], ['LEFTB', 'GAUCHEB'], ['LEN', 'NBCAR'], ['LOWER', 'MINUSCULE'], ['MID', 'STXT'], ['MIDB', 'STXTB'], ['NUMBERVALUE', 'VALEURNOMBRE'],
  ['PROPER', 'NOMPROPRE'], ['REPLACE', 'REMPLACER'], ['REPLACEB', 'REMPLACERB'], ['RIGHT', 'DROITE'], ['RIGHTB', 'DROITEB'], ['SEARCH', 'CHERCHE'],
  ['SEARCHB', 'CHERCHERB'], ['SUBSTITUTE', 'SUBSTITUE'], ['TEXT', 'TEXTE'], ['TEXTJOIN', 'JOINDRE.TEXTE'], ['TRIM', 'SUPPRESPACE'], ['UNICHAR', 'UNICAR'],
  ['UPPER', 'MAJUSCULE'], ['VALUE', 'CNUM'], ['TEXTAFTER', 'TEXTE.APRES'], ['TEXTBEFORE', 'TEXTE.AVANT'], ['TEXTSPLIT', 'FRACTIONNER.TEXTE'],
  ['VALUETOTEXT', 'VALEURENTEXTE'], ['ARRAYTOTEXT', 'TABLEAUENTEXTE'],
  // Recherche et référence
  ['ADDRESS', 'ADRESSE'], ['AREAS', 'ZONES'], ['CHOOSE', 'CHOISIR'], ['CHOOSECOLS', 'CHOISIRCOLS'], ['CHOOSEROWS', 'CHOISIRLIGNES'], ['COLUMN', 'COLONNE'],
  ['COLUMNS', 'COLONNES'], ['DROP', 'EXCLURE'], ['EXPAND', 'ETENDRE'], ['FILTER', 'FILTRE'], ['FORMULATEXT', 'FORMULETEXTE'], ['HLOOKUP', 'RECHERCHEH'],
  ['HSTACK', 'ASSEMB.H'], ['HYPERLINK', 'LIEN_HYPERTEXTE'], ['LOOKUP', 'RECHERCHE'], ['MATCH', 'EQUIV'], ['OFFSET', 'DECALER'], ['ROW', 'LIGNE'], ['ROWS', 'LIGNES'],
  ['SORT', 'TRIER'], ['SORTBY', 'TRIERPAR'], ['TAKE', 'PRENDRE'], ['VLOOKUP', 'RECHERCHEV'], ['VSTACK', 'ASSEMB.V'], ['XLOOKUP', 'RECHERCHEX'], ['XMATCH', 'EQUIVX'],
  // Date et heure
  ['DATEVALUE', 'DATEVAL'], ['DAY', 'JOUR'], ['DAYS', 'JOURS'], ['DAYS360', 'JOURS360'], ['EDATE', 'MOIS.DECALER'], ['EOMONTH', 'FIN.MOIS'], ['HOUR', 'HEURE'],
  ['ISOWEEKNUM', 'NO.SEMAINE.ISO'], ['MONTH', 'MOIS'], ['NETWORKDAYS', 'NB.JOURS.OUVRES'], ['NETWORKDAYS.INTL', 'NB.JOURS.OUVRES.INTL'], ['NOW', 'MAINTENANT'],
  ['SECOND', 'SECONDE'], ['TIME', 'TEMPS'], ['TIMEVALUE', 'TEMPSVAL'], ['TODAY', 'AUJOURDHUI'], ['WEEKDAY', 'JOURSEM'], ['WEEKNUM', 'NUMSEMAINE'],
  ['WORKDAY', 'SERIE.JOUR.OUVRE'], ['WORKDAY.INTL', 'SERIE.JOUR.OUVRE.INTL'], ['YEAR', 'ANNEE'], ['YEARFRAC', 'FRACTION.ANNEE'],
  // Finances
  ['ACCRINT', 'INTERET.ACC'], ['ACCRINTM', 'INTERET.ACC.MAT'], ['COUPDAYBS', 'NB.JOURS.COUPON.PREC'], ['COUPDAYS', 'NB.JOURS.COUPON'],
  ['COUPDAYSNC', 'NB.JOURS.COUPON.SUIV'], ['COUPNCD', 'DATE.COUPON.SUIV'], ['COUPNUM', 'NB.COUPONS'], ['COUPPCD', 'DATE.COUPON.PREC'], ['CUMIPMT', 'CUMUL.INTER'],
  ['CUMPRINC', 'CUMUL.PRINCPER'], ['DISC', 'TAUX.ESCOMPTE'], ['DOLLARDE', 'PRIX.DEC'], ['DOLLARFR', 'PRIX.FRAC'], ['DURATION', 'DUREE'], ['EFFECT', 'TAUX.EFFECTIF'],
  ['FV', 'VC'], ['FVSCHEDULE', 'VC.PAIEMENTS'], ['INTRATE', 'TAUX.INTERET'], ['IPMT', 'INTPER'], ['IRR', 'TRI'], ['MDURATION', 'DUREE.MODIFIEE'], ['MIRR', 'TRIM'],
  ['NOMINAL', 'TAUX.NOMINAL'], ['NPER', 'NPM'], ['NPV', 'VAN'], ['PDURATION', 'PDUREE'], ['PMT', 'VPM'], ['PPMT', 'PRINCPER'], ['PRICE', 'PRIX.TITRE'],
  ['PRICEDISC', 'PRIX.TITRE.ESCOMPTE'], ['PRICEMAT', 'PRIX.TITRE.ECHEANCE'], ['PV', 'VA'], ['RATE', 'TAUX'], ['RECEIVED', 'VALEUR.ENCAISSEMENT'],
  ['RRI', 'TAUX.INT.EQUIV'], ['SLN', 'AMORLIN'], ['XIRR', 'TRI.PAIEMENTS'], ['XNPV', 'VAN.PAIEMENTS'], ['YIELD', 'RENDEMENT.TITRE'], ['YIELDDISC', 'RENDEMENT.SIMPLE'],
  ['YIELDMAT', 'RENDEMENT.TITRE.ECHEANCE'],
  // Information
  ['CELL', 'CELLULE'], ['ERROR.TYPE', 'TYPE.ERREUR'], ['INFO', 'INFORMATIONS'], ['ISBLANK', 'ESTVIDE'], ['ISERR', 'ESTERR'], ['ISERROR', 'ESTERREUR'],
  ['ISEVEN', 'EST.PAIR'], ['ISFORMULA', 'ESTFORMULE'], ['ISLOGICAL', 'ESTLOGIQUE'], ['ISNA', 'ESTNA'], ['ISNONTEXT', 'ESTNONTEXTE'], ['ISNUMBER', 'ESTNUM'],
  ['ISODD', 'EST.IMPAIR'], ['ISREF', 'ESTREF'], ['ISTEXT', 'ESTTEXTE'], ['SHEET', 'FEUILLE'], ['SHEETS', 'FEUILLES'],
  // Base de données
  ['DAVERAGE', 'BDMOYENNE'], ['DCOUNT', 'BDNB'], ['DCOUNTA', 'BDNBVAL'], ['DGET', 'BDLIRE'], ['DMAX', 'BDMAX'], ['DMIN', 'BDMIN'], ['DPRODUCT', 'BDPRODUIT'],
  ['DSTDEV', 'BDECARTYPE'], ['DSTDEVP', 'BDECARTYPEP'], ['DSUM', 'BDSOMME'], ['DVAR', 'BDVAR'], ['DVARP', 'BDVARP'],
  // Ingénierie
  ['BIN2DEC', 'BINDEC'], ['BIN2HEX', 'BINHEX'], ['BIN2OCT', 'BINOCT'], ['BITAND', 'BITET'], ['BITLSHIFT', 'BITDECALG'], ['BITOR', 'BITOU'], ['BITRSHIFT', 'BITDECALD'],
  ['BITXOR', 'BITOUEXCLUSIF'], ['COMPLEX', 'COMPLEXE'], ['DEC2BIN', 'DECBIN'], ['DEC2HEX', 'DECHEX'], ['DEC2OCT', 'DECOCT'], ['GESTEP', 'SUP.SEUIL'],
  ['HEX2BIN', 'HEXBIN'], ['HEX2DEC', 'HEXDEC'], ['HEX2OCT', 'HEXOCT'], ['IMABS', 'COMPLEXE.MODULE'], ['IMAGINARY', 'COMPLEXE.IMAGINAIRE'], ['IMREAL', 'COMPLEXE.REEL'],
  ['OCT2BIN', 'OCTBIN'], ['OCT2DEC', 'OCTDEC'], ['OCT2HEX', 'OCTHEX'],
];

export const EN_TO_FR = new Map<string, string>(PAIRS.map(([en, fr]) => [en, fr]));
/** Noms français → anglais. Si deux noms anglais ont le même nom français, le premier gagne. */
export const FR_TO_EN = new Map<string, string>();
for (const [en, fr] of PAIRS) if (!FR_TO_EN.has(fr)) FR_TO_EN.set(fr, en);

const ERRORS_EN_FR: [string, string][] = [['#NAME?', '#NOM?'], ['#VALUE!', '#VALEUR!'], ['#NUM!', '#NOMBRE!'], ['#NULL!', '#NUL!'], ['#SPILL!', '#DEVERSEMENT!']];
const ERR_FR_TO_EN = new Map(ERRORS_EN_FR.map(([en, fr]) => [fr, en]));
const ERR_EN_TO_FR = new Map(ERRORS_EN_FR);

const TABLE_KEYS_EN_FR: [string, string][] = [['#Data', '#Données'], ['#Headers', '#En-têtes'], ['#Totals', '#Totaux'], ['#All', '#Tout'], ['#This Row', '#Cette ligne']];

type Direction = 'toEnglish' | 'toFrench';

const isIdentStart = (c: string) => /[A-Za-z_]/.test(c);
const isIdentChar = (c: string) => /[A-Za-z0-9_.]/.test(c);

/** Parcourt la formule (sans le « = » initial) et traduit chaque élément ; chaînes et noms de feuille entre apostrophes sont conservés. */
function translateBody(body: string, dir: Direction): string {
  const fn = dir === 'toEnglish' ? FR_TO_EN : EN_TO_FR;
  const bool = dir === 'toEnglish' ? new Map([['VRAI', 'TRUE'], ['FAUX', 'FALSE']]) : new Map([['TRUE', 'VRAI'], ['FALSE', 'FAUX']]);
  const errors = dir === 'toEnglish' ? ERR_FR_TO_EN : ERR_EN_TO_FR;
  let out = '';
  let braces = 0;
  let i = 0;
  while (i < body.length) {
    const c = body[i];
    // Chaîne de texte "…" (guillemets doublés pour un guillemet).
    if (c === '"') {
      let j = i + 1;
      while (j < body.length && (body[j] !== '"' || body[j + 1] === '"')) j += body[j] === '"' ? 2 : 1;
      out += body.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    // Nom de feuille entre apostrophes.
    if (c === "'") {
      let j = i + 1;
      while (j < body.length && (body[j] !== "'" || body[j + 1] === "'")) j += body[j] === "'" ? 2 : 1;
      out += body.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    // Référence structurée Tableau[[#Data],[Colonne]] : seuls les mots-clés # sont traduits.
    if (c === '[') {
      let depth = 0;
      let j = i;
      for (; j < body.length; j++) {
        if (body[j] === '[') depth++;
        else if (body[j] === ']' && --depth === 0) break;
      }
      let token = body.slice(i, j + 1);
      for (const [en, fr] of TABLE_KEYS_EN_FR) token = dir === 'toEnglish' ? token.split(fr).join(en) : token.split(en).join(fr);
      out += token;
      i = j + 1;
      continue;
    }
    // Erreur littérale (#NOM?, #VALEUR!…).
    if (c === '#') {
      const m = /^#[A-Za-zÀ-ÿ/0-9]+[!?]?/.exec(body.slice(i));
      if (m) {
        out += errors.get(m[0].toUpperCase()) ?? m[0];
        i += m[0].length;
        continue;
      }
    }
    // Nombre : la virgule décimale devient un point (et inversement).
    if (/[0-9]/.test(c) && !(i > 0 && /[A-Za-z0-9_.$!]/.test(body[i - 1]))) {
      // En anglais (affichage), seule la virgule est un séparateur d'arguments : le point seul est décimal.
      const m = (dir === 'toEnglish' ? /^\d+(?:,\d+|\.\d+)?(?:[eE][+-]?\d+)?/ : /^\d+(?:\.\d+)?(?:[eE][+-]?\d+)?/).exec(body.slice(i))!;
      const dotted = m[0];
      // « 1,5 » → « 1.5 » en français ; en anglais « 1.5 » → « 1,5 ». Une virgule d'argument (« 1,2 » anglais) n'est pas ici : le français l'écrit « ; ».
      out += dir === 'toEnglish' ? dotted.replace(',', '.') : dotted.replace('.', ',');
      i += m[0].length;
      continue;
    }
    if (isIdentStart(c)) {
      let j = i + 1;
      while (j < body.length && isIdentChar(body[j])) j++;
      const name = body.slice(i, j);
      const upper = name.toUpperCase();
      let k = j;
      while (body[k] === ' ') k++;
      if (body[k] === '(') out += fn.get(upper) ?? name;
      else if (bool.has(upper) && body[k] !== '!' && !(i > 0 && /[!$]/.test(body[i - 1]))) out += bool.get(upper);
      else out += name;
      i = j;
      continue;
    }
    if (c === '{') braces++;
    else if (c === '}') braces = Math.max(0, braces - 1);
    if (dir === 'toEnglish') {
      // « ; » sépare les arguments (français) ; dans {…}, « ; » sépare les lignes et « \ » les colonnes.
      if (c === ';' && braces === 0) out += ',';
      else if (c === '\\' && braces > 0) out += ',';
      else out += c;
    } else if (c === ',') out += braces > 0 ? '\\' : ';';
    else out += c;
    i++;
  }
  return out;
}

/**
 * Vrai si la formule saisie est déjà en anglais (collée d'ailleurs, par exemple) : aucun « ; » hors
 * texte, et au moins un nom de fonction anglais qui n'existe pas en français.
 */
export function looksEnglish(formula: string): boolean {
  const body = formula.replace(/"(?:[^"]|"")*"/g, '""');
  if (body.includes(';')) return false;
  let english = false;
  for (const m of body.matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g)) {
    const name = m[1].toUpperCase();
    if (FR_TO_EN.has(name) && !EN_TO_FR.has(name)) return false;
    if (EN_TO_FR.has(name) && !FR_TO_EN.has(name)) english = true;
  }
  return english;
}

/** Formule saisie en français → forme anglaise stockée. Les textes qui ne commencent pas par « = » sont renvoyés tels quels. */
export function frenchToEnglish(formula: string): string {
  if (!formula.startsWith('=') || looksEnglish(formula)) return formula;
  return `=${translateBody(formula.slice(1), 'toEnglish')}`;
}

/** Formule stockée (anglais) → forme française affichée dans la barre de formule. */
export function englishToFrench(formula: string): string {
  if (!formula.startsWith('=')) return formula;
  return `=${translateBody(formula.slice(1), 'toFrench')}`;
}


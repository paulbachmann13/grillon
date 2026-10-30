/**
 * Banc de comparaison avec Excel : mêmes formules et mêmes formats calculés par le vrai Excel et par Grillon.
 * Les résultats d'Excel sont gardés dans excel-reference.json (relevés avec Excel pour Mac, voir docs/EXCEL-COMPARAISON.md).
 */
import type { IWorkbookData } from '../../src/model/univer-types';

const C = 'D!$C$2:$C$8'; // notes
const A = 'D!$A$2:$A$8'; // noms
const B = 'D!$B$2:$B$8'; // villes
const DT = 'D!$D$2:$D$8'; // dates
const E = 'D!$E$2:$E$8'; // nombres avec négatifs
const G = 'D!$G$2:$G$8'; // mélange (nombre, texte, vide, logique)
const X = 'D!$H$2:$H$8';
const Y = 'D!$I$2:$I$8';
const T = 'D!$A$1:$I$8'; // tout le tableau

/** Formules sans résultat aléatoire ni date du jour. */
export const FORMULAS: string[] = [
  // Maths
  `SUM(${C})`, `SUM(${G})`, `SUM(1,"2",TRUE)`, `SUMIF(${B},"Lyon",${C})`, `SUMIF(${C},">10")`, `SUMIFS(${C},${B},"Paris",${C},">5")`, `SUMPRODUCT(${C},${E})`, `SUMSQ(1,2,3)`,
  `SUMX2MY2(${X},${Y})`, `SUMX2PY2(${X},${Y})`, `SUMXMY2(${X},${Y})`, `PRODUCT(${X})`, `ABS(-3.5)`, `ROUND(2.5,0)`, `ROUND(-2.5,0)`, `ROUND(1.005,2)`, `ROUND(1234.567,-2)`, `ROUNDUP(1.21,1)`, `ROUNDDOWN(-1.29,1)`,
  `MROUND(10,3)`, `MROUND(-10,-3)`, `MROUND(10,-3)`, `CEILING(2.5,1)`, `CEILING(-2.5,-2)`, `CEILING.MATH(-2.5)`, `CEILING.MATH(-2.5,2,1)`, `CEILING.PRECISE(-4.1,2)`, `FLOOR(2.5,1)`, `FLOOR(-2.5,-2)`, `FLOOR(-2.5,2)`, `FLOOR.MATH(-2.5)`, `FLOOR.PRECISE(4.1,2)`, `ISO.CEILING(4.1,2)`,
  `INT(-1.5)`, `TRUNC(-1.59,1)`, `MOD(10,3)`, `MOD(-10,3)`, `MOD(10,-3)`, `MOD(5,0)`, `QUOTIENT(7,2)`, `QUOTIENT(-7,2)`, `POWER(2,10)`, `POWER(0,0)`, `SQRT(16)`, `SQRT(-1)`, `SQRTPI(1)`, `EXP(1)`, `LN(10)`, `LOG(100)`, `LOG(8,2)`, `LOG10(1000)`, `LN(0)`,
  `SIGN(-2)`, `EVEN(3)`, `EVEN(-3)`, `ODD(2)`, `ODD(-2)`, `FACT(5)`, `FACTDOUBLE(7)`, `COMBIN(10,3)`, `COMBINA(4,3)`, `PERMUT(10,3)`, `PERMUTATIONA(3,2)`, `GCD(12,18)`, `LCM(4,6)`, `MULTINOMIAL(2,3,4)`, `SERIESSUM(2,0,1,{1,2,3})`,
  `PI()`, `DEGREES(PI())`, `RADIANS(180)`, `SIN(1)`, `COS(1)`, `TAN(1)`, `ASIN(0.5)`, `ACOS(0.5)`, `ATAN(1)`, `ATAN2(1,2)`, `SINH(1)`, `COSH(1)`, `TANH(1)`, `ASINH(1)`, `ACOSH(2)`, `ATANH(0.5)`, `COT(1)`, `COTH(1)`, `ACOT(1)`, `ACOTH(2)`, `CSC(1)`, `CSCH(1)`, `SEC(1)`, `SECH(1)`,
  `ROMAN(1994)`, `ARABIC("MCMXCIV")`, `BASE(255,16)`, `DECIMAL("FF",16)`, `MDETERM({1,2;3,4})`, `SUM(MINVERSE({1,2;3,4}))`, `SUM(MMULT({1,2;3,4},{1,2;3,4}))`, `SUM(MUNIT(3))`, `SUBTOTAL(9,${C})`, `SUBTOTAL(1,${C})`, `SUBTOTAL(109,${C})`, `AGGREGATE(9,6,${C})`, `AGGREGATE(14,6,${C},2)`,
  // Statistiques
  `AVERAGE(${C})`, `AVERAGE(${G})`, `AVERAGEA(${G})`, `AVERAGEIF(${B},"Lyon",${C})`, `AVERAGEIFS(${C},${B},"Paris",${C},">5")`, `COUNT(${G})`, `COUNTA(${G})`, `COUNTBLANK(${G})`, `COUNTIF(${B},"Lyon")`, `COUNTIF(${A},"?o*")`, `COUNTIF(${C},">=12.5")`, `COUNTIF(${B},"lyon")`, `COUNTIFS(${B},"Lyon",${C},">12.5")`,
  `MAX(${C})`, `MIN(${C})`, `MAXA(${G})`, `MINA(${G})`, `MAXIFS(${C},${B},"Lyon")`, `MINIFS(${C},${B},"Lyon")`, `MEDIAN(${C})`, `MODE(${C})`, `MODE.SNGL(${C})`, `SUM(MODE.MULT(${C}))`, `LARGE(${C},2)`, `SMALL(${C},2)`, `RANK(12.5,${C})`, `RANK.EQ(12.5,${C},1)`, `RANK.AVG(12.5,${C})`,
  `STDEV(${C})`, `STDEV.S(${C})`, `STDEV.P(${C})`, `STDEVP(${C})`, `STDEVA(${G})`, `STDEVPA(${G})`, `VAR(${C})`, `VAR.S(${C})`, `VAR.P(${C})`, `VARP(${C})`, `VARA(${G})`, `VARPA(${G})`, `AVEDEV(${C})`, `DEVSQ(${C})`, `GEOMEAN(${X})`, `HARMEAN(${X})`, `TRIMMEAN(${C},0.3)`, `SKEW(${C})`, `SKEW.P(${C})`, `KURT(${C})`,
  `PERCENTILE(${C},0.3)`, `PERCENTILE.INC(${C},0.3)`, `PERCENTILE.EXC(${C},0.3)`, `PERCENTRANK(${C},12.5)`, `PERCENTRANK.INC(${C},12.5)`, `PERCENTRANK.EXC(${C},12.5)`, `QUARTILE(${C},1)`, `QUARTILE.INC(${C},3)`, `QUARTILE.EXC(${C},1)`,
  `CORREL(${X},${Y})`, `PEARSON(${X},${Y})`, `RSQ(${X},${Y})`, `SLOPE(${Y},${X})`, `INTERCEPT(${Y},${X})`, `STEYX(${Y},${X})`, `COVAR(${X},${Y})`, `COVARIANCE.P(${X},${Y})`, `COVARIANCE.S(${X},${Y})`, `FORECAST(8,${Y},${X})`, `FORECAST.LINEAR(8,${Y},${X})`, `INDEX(LINEST(${Y},${X}),1,1)`, `INDEX(LOGEST(${Y},${X}),1,1)`, `SUM(TREND(${Y},${X}))`, `SUM(GROWTH(${Y},${X}))`, `SUM(FREQUENCY(${C},{5,10,15}))`,
  `NORM.DIST(1,0,1,TRUE)`, `NORM.DIST(1,0,1,FALSE)`, `NORMDIST(1,0,1,TRUE)`, `NORM.INV(0.9,0,1)`, `NORMINV(0.9,0,1)`, `NORM.S.DIST(1,TRUE)`, `NORMSDIST(1)`, `NORM.S.INV(0.9)`, `NORMSINV(0.9)`, `STANDARDIZE(1,0,2)`, `GAUSS(1)`, `PHI(1)`, `CONFIDENCE(0.05,1,50)`, `CONFIDENCE.NORM(0.05,1,50)`, `CONFIDENCE.T(0.05,1,50)`,
  `BINOM.DIST(3,10,0.5,TRUE)`, `BINOMDIST(3,10,0.5,FALSE)`, `BINOM.INV(10,0.5,0.7)`, `CRITBINOM(10,0.5,0.7)`, `BINOM.DIST.RANGE(10,0.5,2,4)`, `NEGBINOM.DIST(3,5,0.5,TRUE)`, `NEGBINOMDIST(3,5,0.5)`, `POISSON.DIST(3,2,TRUE)`, `POISSON(3,2,FALSE)`, `HYPGEOM.DIST(1,4,8,20,TRUE)`, `HYPGEOMDIST(1,4,8,20)`, `EXPON.DIST(1,2,TRUE)`, `EXPONDIST(1,2,FALSE)`, `GAMMA(5)`, `GAMMALN(5)`, `GAMMALN.PRECISE(5)`, `GAMMA.DIST(2,3,2,TRUE)`, `GAMMADIST(2,3,2,FALSE)`, `GAMMA.INV(0.5,3,2)`, `GAMMAINV(0.5,3,2)`, `BETA.DIST(0.5,2,3,TRUE)`, `BETADIST(0.5,2,3)`, `BETA.INV(0.5,2,3)`, `BETAINV(0.5,2,3)`,
  `T.DIST(1,5,TRUE)`, `T.DIST.2T(1,5)`, `T.DIST.RT(1,5)`, `TDIST(1,5,2)`, `T.INV(0.9,5)`, `T.INV.2T(0.1,5)`, `TINV(0.1,5)`, `T.TEST(${X},${Y},2,1)`, `TTEST(${X},${Y},2,3)`, `F.DIST(1,5,5,TRUE)`, `F.DIST.RT(1,5,5)`, `FDIST(1,5,5)`, `F.INV(0.5,5,5)`, `F.INV.RT(0.5,5,5)`, `FINV(0.5,5,5)`, `F.TEST(${X},${Y})`, `FTEST(${X},${Y})`, `CHISQ.DIST(1,5,TRUE)`, `CHISQ.DIST.RT(1,5)`, `CHIDIST(1,5)`, `CHISQ.INV(0.5,5)`, `CHISQ.INV.RT(0.5,5)`, `CHIINV(0.5,5)`, `CHISQ.TEST(${X},${Y})`, `Z.TEST(${C},10)`, `ZTEST(${C},10)`, `WEIBULL.DIST(1,2,3,TRUE)`, `WEIBULL(1,2,3,FALSE)`, `LOGNORM.DIST(1,0,1,TRUE)`, `LOGNORMDIST(1,0,1)`, `LOGNORM.INV(0.5,0,1)`, `LOGINV(0.5,0,1)`, `FISHER(0.5)`, `FISHERINV(0.5)`, `PROB(${X},{0.1,0.2,0.3,0.1,0.1,0.1,0.1},2,4)`,
  // Logique
  `IF(1>2,"a","b")`, `IF(1>2,"a")`, `IFS(1>2,"a",TRUE,"b")`, `IFERROR(1/0,"err")`, `IFNA(NA(),"na")`, `AND(TRUE,1)`, `OR(FALSE,0)`, `NOT(1)`, `XOR(TRUE,TRUE,TRUE)`, `SWITCH(2,1,"a",2,"b","c")`, `LET(x,2,x*3)`, `TRUE()`, `FALSE()`, `AND(${G})`, `OR("a")`,
  // Comparaisons et conversions
  `1="1"`, `"a"="A"`, `"a"<"b"`, `"é">"e"`, `TRUE=1`, `TRUE>1000`, `"1">5`, `1+"2"`, `"2"*"3"`, `TRUE+TRUE`, `1/0`, `0/0`, `"a"+1`, `2^3^2`, `-2^2`, `"a"&1&TRUE`, `5>4>3`, `0.1+0.2=0.3`,
  // Texte
  `CONCAT("a",1,TRUE)`, `CONCATENATE("a","b")`, `TEXTJOIN(";",TRUE,${A})`, `TEXTJOIN(",",FALSE,${G})`, `LEFT("Bonjour",3)`, `RIGHT("Bonjour",3)`, `MID("Bonjour",2,3)`, `LEN("été")`, `LENB("été")`, `LEFTB("Bonjour",2)`, `RIGHTB("Bonjour",2)`, `MIDB("Bonjour",2,2)`, `FIND("o","Bonjour")`, `FIND("o","Bonjour",3)`, `FIND("z","Bonjour")`, `SEARCH("J","Bonjour")`, `SEARCH("?o","Bonjour")`, `FINDB("o","Bonjour")`, `SEARCHB("j","Bonjour")`,
  `SUBSTITUTE("aXbXc","X","-")`, `SUBSTITUTE("aXbXc","X","-",2)`, `REPLACE("Bonjour",2,3,"--")`, `REPLACEB("Bonjour",2,3,"--")`, `UPPER("été")`, `LOWER("ÉTÉ")`, `PROPER("jean-paul d'arc")`, `TRIM("  a   b  ")`, `CLEAN("a"&CHAR(10)&"b")`, `REPT("ab",3)`, `CHAR(65)`, `CODE("A")`, `UNICHAR(233)`, `UNICODE("é")`, `EXACT("a","A")`, `T("a")`, `T(1)`, `N("a")`, `N(TRUE)`, `VALUE("12,5")`, `VALUE("1 000")`, `VALUE("12%")`, `VALUE("a")`, `NUMBERVALUE("1 234,5",","," ")`, `NUMBERVALUE("1.234,5",",",".")`,
  `TEXT(1234.567,"# ##0,00")`, `TEXT(0.25,"0%")`, `TEXT(45292,"jj/mm/aaaa")`, `TEXT(45292,"jjjj d mmmm aaaa")`, `TEXT(1.5,"0,0")`, `TEXT(-5,"0;(0)")`, `TEXT(0.5,"hh:mm")`, `TEXT(1234,"0,00E+00")`, `TEXT("a","@@")`, `FIXED(1234.567,1)`, `FIXED(1234.567,1,TRUE)`, `DOLLAR(1234.567,2)`, `BAHTTEXT(12)`, `TEXTBEFORE("a-b-c","-")`, `TEXTAFTER("a-b-c","-",2)`, `INDEX(TEXTSPLIT("a,b;c",","),1,2)`, `ARRAYTOTEXT({1,2,"a"})`, `VALUETOTEXT("a",1)`, `ASC("a")`, `DBCS("a")`,
  // Date
  `DATE(2024,2,30)`, `DATE(24,1,1)`, `DATE(2024,14,1)`, `YEAR(45292)`, `MONTH(45292)`, `DAY(45292)`, `HOUR(0.75)`, `MINUTE(0.7513)`, `SECOND(0.75134)`, `TIME(25,61,1)`, `WEEKDAY(45292)`, `WEEKDAY(45292,2)`, `WEEKDAY(45292,11)`, `WEEKNUM(45292)`, `WEEKNUM(45292,2)`, `WEEKNUM(45292,21)`, `ISOWEEKNUM(45292)`, `EDATE(45322,1)`, `EOMONTH(45292,1)`, `DAYS(45322,45292)`, `DAYS360(45292,45322)`, `DAYS360(45292,45322,TRUE)`, `DATEDIF(45292,45500,"M")`, `DATEDIF(45292,45500,"D")`, `DATEDIF(45292,45800,"Y")`, `DATEDIF(45292,45500,"MD")`, `DATEDIF(45292,45500,"YM")`, `DATEDIF(45292,45500,"YD")`, `NETWORKDAYS(45292,45322)`, `NETWORKDAYS(45292,45322,45293)`, `NETWORKDAYS.INTL(45292,45322,7)`, `NETWORKDAYS.INTL(45292,45322,"0000011")`, `WORKDAY(45292,10)`, `WORKDAY(45292,-10)`, `WORKDAY.INTL(45292,10,7)`, `YEARFRAC(45292,45500)`, `YEARFRAC(45292,45500,1)`, `YEARFRAC(45292,45500,3)`, `DATEVALUE("15/03/2024")`, `TIMEVALUE("12:30")`, `DATEVALUE("2024-03-15")`, `DATE(1900,1,1)`, `DATE(1900,3,1)`, `DAY(60)`, `MONTH(60)`,
  // Recherche
  `VLOOKUP("Zoé",D!$A$2:$C$8,3,FALSE)`, `VLOOKUP("zoé",D!$A$2:$C$8,3,FALSE)`, `VLOOKUP("Inconnu",D!$A$2:$C$8,3,FALSE)`, `VLOOKUP("?o*",D!$A$2:$C$8,3,FALSE)`, `VLOOKUP(5,D!$H$2:$I$8,2,TRUE)`, `VLOOKUP(5,D!$H$2:$I$8,5,FALSE)`, `HLOOKUP("Note",D!$A$1:$C$8,3,FALSE)`, `LOOKUP(5,${X},${Y})`, `LOOKUP(0,${X},${Y})`, `INDEX(${C},3)`, `INDEX(D!$A$2:$C$8,2,3)`, `INDEX(${C},99)`, `MATCH("Tom",${A},0)`, `MATCH(12.5,${C},0)`, `MATCH(4,${X},1)`, `MATCH(4,${X},-1)`, `MATCH("T*",${A},0)`, `XLOOKUP("Zoé",${A},${C})`, `XLOOKUP("Inconnu",${A},${C},"absent")`, `XLOOKUP(4.5,${X},${Y},,-1)`, `XLOOKUP(4.5,${X},${Y},,1)`, `XLOOKUP("Tom",${A},${C},,0,-1)`, `XMATCH("Tom",${A})`, `XMATCH("Tom",${A},0,-1)`, `CHOOSE(2,"a","b")`, `CHOOSE(5,"a","b")`, `SUM(OFFSET(D!$C$2,1,0,2,1))`, `SUM(INDIRECT("D!C2:C4"))`, `ROW(D!$C$5)`, `COLUMN(D!$C$5)`, `ROWS(${C})`, `COLUMNS(${T})`, `ADDRESS(2,3)`, `ADDRESS(2,3,4)`, `ADDRESS(2,3,1,FALSE,"D")`, `AREAS((D!$A$1,D!$C$1))`, `SUM(TRANSPOSE(${X}))`, `SUM(SORT(${C}))`, `INDEX(SORT(${C}),1)`, `INDEX(SORT(${C},1,-1),1)`, `INDEX(SORT(${A}),1)`, `INDEX(SORTBY(${A},${C},-1),1)`, `SUM(UNIQUE(${C}))`, `INDEX(UNIQUE(${B}),2)`, `ROWS(UNIQUE(${B}))`, `SUM(FILTER(${C},${B}="Lyon"))`, `ROWS(FILTER(${C},${C}>100,"x"))`, `SUM(SEQUENCE(3))`, `SUM(SEQUENCE(2,3,5,2))`, `SUM(TAKE(${C},2))`, `SUM(DROP(${C},2))`, `SUM(HSTACK(${X},${Y}))`, `SUM(VSTACK(${X},${Y}))`, `SUM(CHOOSECOLS(D!$H$2:$I$8,2))`, `SUM(CHOOSEROWS(${C},1,2))`, `SUM(EXPAND({1,2},3,3,0))`, `SUM(TOCOL(D!$H$2:$I$8))`, `SUM(TOROW(D!$H$2:$I$8))`, `SUM(WRAPROWS(${X},2,0))`, `SUM(WRAPCOLS(${X},2,0))`, `FORMULATEXT(D!$J$2)`, `HYPERLINK("https://a.fr","lien")`, `SINGLE(${C})`,
  // Information
  `ISBLANK(D!$G$4)`, `ISBLANK("")`, `ISERR(1/0)`, `ISERR(NA())`, `ISERROR(NA())`, `ISNA(NA())`, `ISEVEN(4)`, `ISODD(4)`, `ISFORMULA(D!$J$2)`, `ISLOGICAL(TRUE)`, `ISNONTEXT(1)`, `ISNUMBER("1")`, `ISNUMBER(1)`, `ISREF(D!$A$1)`, `ISTEXT("a")`, `NA()`, `ERROR.TYPE(1/0)`, `ERROR.TYPE(NA())`, `TYPE(1)`, `TYPE("a")`, `TYPE(TRUE)`, `TYPE(1/0)`, `TYPE({1,2})`, `SHEET()`, `SHEETS()`,
  // Financier
  `PMT(0.05/12,120,10000)`, `PMT(0.05/12,120,10000,0,1)`, `IPMT(0.05/12,1,120,10000)`, `PPMT(0.05/12,1,120,10000)`, `FV(0.05/12,120,-100)`, `PV(0.05/12,120,-100)`, `NPV(0.1,100,200,300)`, `XNPV(0.1,{-100,50,60},{45292,45500,45800})`, `IRR({-100,50,60})`, `XIRR({-100,50,60},{45292,45500,45800})`, `MIRR({-100,50,60},0.1,0.1)`, `RATE(120,-100,10000)`, `NPER(0.05/12,-100,5000)`, `SLN(1000,100,5)`, `SYD(1000,100,5,1)`, `DB(1000,100,5,1)`, `DDB(1000,100,5,1)`, `VDB(1000,100,5,0,1)`, `CUMIPMT(0.05/12,120,10000,1,12,0)`, `CUMPRINC(0.05/12,120,10000,1,12,0)`, `EFFECT(0.05,12)`, `NOMINAL(0.05,12)`, `ISPMT(0.05,1,10,1000)`, `PDURATION(0.05,100,200)`, `RRI(10,100,200)`, `DOLLARDE(1.02,16)`, `DOLLARFR(1.125,16)`, `TBILLEQ(45292,45400,0.05)`, `TBILLPRICE(45292,45400,0.05)`, `TBILLYIELD(45292,45400,98)`, `PRICE(45292,46000,0.05,0.04,100,2)`, `YIELD(45292,46000,0.05,98,100,2)`, `ACCRINT(45292,45400,45350,0.05,1000,2)`, `ACCRINTM(45292,45500,0.05,1000)`, `DISC(45292,45500,98,100)`, `INTRATE(45292,45500,98,100)`, `RECEIVED(45292,45500,98,0.05)`, `PRICEDISC(45292,45500,0.05,100)`, `PRICEMAT(45292,45800,45200,0.05,0.04)`, `YIELDDISC(45292,45500,98,100)`, `YIELDMAT(45292,45800,45200,0.05,98)`, `DURATION(45292,46000,0.05,0.04,2)`, `MDURATION(45292,46000,0.05,0.04,2)`, `COUPDAYBS(45292,46000,2)`, `COUPDAYS(45292,46000,2)`, `COUPDAYSNC(45292,46000,2)`, `COUPNCD(45292,46000,2)`, `COUPNUM(45292,46000,2)`, `COUPPCD(45292,46000,2)`, `AMORLINC(1000,45292,45500,100,1,0.1)`, `AMORDEGRC(1000,45292,45500,100,1,0.1)`, `ODDFPRICE(45292,46000,45200,45400,0.05,0.04,100,2)`, `ODDLPRICE(45700,46000,45600,0.05,0.04,100,2)`,
  // Ingénierie
  `BIN2DEC("101")`, `BIN2HEX("101")`, `BIN2OCT("101")`, `DEC2BIN(5)`, `DEC2HEX(255)`, `DEC2OCT(8)`, `HEX2BIN("F")`, `HEX2DEC("FF")`, `HEX2OCT("F")`, `OCT2BIN("7")`, `OCT2DEC("10")`, `OCT2HEX("10")`, `BITAND(5,3)`, `BITOR(5,3)`, `BITXOR(5,3)`, `BITLSHIFT(1,3)`, `BITRSHIFT(8,2)`, `DELTA(1,1)`, `GESTEP(2,1)`, `ERF(1)`, `ERF.PRECISE(1)`, `ERFC(1)`, `ERFC.PRECISE(1)`, `BESSELI(1,1)`, `BESSELJ(1,1)`, `BESSELK(1,1)`, `BESSELY(1,1)`, `CONVERT(1,"m","ft")`, `CONVERT(100,"C","F")`, `COMPLEX(1,2)`, `IMABS("3+4i")`, `IMAGINARY("3+4i")`, `IMREAL("3+4i")`, `IMARGUMENT("1+i")`, `IMCONJUGATE("3+4i")`, `IMSUM("1+i","2+i")`, `IMSUB("1+i","2+i")`, `IMPRODUCT("1+i","2+i")`, `IMDIV("1+i","2+i")`, `IMPOWER("1+i",2)`, `IMSQRT("3+4i")`, `IMEXP("1+i")`, `IMLN("1+i")`, `IMLOG10("1+i")`, `IMLOG2("1+i")`, `IMSIN("1+i")`, `IMCOS("1+i")`, `IMTAN("1+i")`, `IMSEC("1+i")`, `IMCSC("1+i")`, `IMCOT("1+i")`, `IMSINH("1+i")`, `IMCOSH("1+i")`, `IMSECH("1+i")`, `IMCSCH("1+i")`,
  // Base de données
  `DSUM(${T},"Note",D!$K$1:$K$2)`, `DAVERAGE(${T},"Note",D!$K$1:$K$2)`, `DCOUNT(${T},"Note",D!$K$1:$K$2)`, `DCOUNTA(${T},"Nom",D!$K$1:$K$2)`, `DMAX(${T},"Note",D!$K$1:$K$2)`, `DMIN(${T},"Note",D!$K$1:$K$2)`, `DPRODUCT(${T},"Note",D!$K$1:$K$2)`, `DSTDEV(${T},"Note",D!$K$1:$K$2)`, `DSTDEVP(${T},"Note",D!$K$1:$K$2)`, `DVAR(${T},"Note",D!$K$1:$K$2)`, `DVARP(${T},"Note",D!$K$1:$K$2)`, `DGET(${T},"Nom",D!$K$1:$K$2)`,
  // À la française
  `"x"&1.5`, `"x"&TRUE`, `"1,5"+1`, `"1,5"*2`, `COUNTIF(${C},">=12,5")`, `COUNTIF(${C},"12,5")`, `SUMIF(${C},">"&12.5)`, `AVERAGEIF(${C},">="&AVERAGE(${C}))`, `LEFT(1.5,2)`, `LEN(1.5)`, `TEXT(1234.5,"# ##0,00 €")`, `TEXT(0.256,"0,0%")`, `TEXT(45292,"jj mmm aa")`, `TEXT(45292.5,"jj/mm/aaaa hh:mm")`, `TEXT(D!$D$2,"mmmm")`, `UPPER(TEXT(D!$D$2,"mmm"))`, `TEXT(TRUE,"0")`, `SUM(VALUE("1,5"),VALUE("2,5"))`,
  // Parcoursup et bureau : usages courants
  `SUMPRODUCT((${B}="Lyon")*(${C}>10))`, `SUMPRODUCT((${B}="Lyon")*${C})`, `SUM((${B}="Lyon")*${C})`, `COUNTIF(${C},">"&AVERAGE(${C}))`, `AVERAGEIF(${C},">="&10)`, `RANK(D!$C$2,${C},0)`, `RANK(D!$C$2,${C},0)+COUNTIF(D!$C$2:$C$2,D!$C$2)-1`, `IF(D!$C$2>=10,"Admis","Refusé")`, `TEXT(D!$D$2,"mmmm aaaa")`, `VLOOKUP(MAX(${C}),CHOOSE({1,2},${C},${A}),2,FALSE)`, `INDEX(${A},MATCH(MAX(${C}),${C},0))`, `ROUND(AVERAGE(${C}),2)`, `SUMIF(${DT},">="&DATE(2024,3,1),${C})`, `COUNTIFS(${DT},">="&DATE(2024,3,1),${DT},"<"&DATE(2024,5,1))`, `LEFT(${A},1)&"."`, `MAX(IF(${B}="Lyon",${C}))`,
];

/** Valeur × format : texte affiché par Excel et Grillon. */
export const NUMBER_FORMATS: [value: number | string | boolean, format: string][] = [
  ...[1234.567, -1234.567, 0, 0.5].flatMap((v): [number, string][] =>
    ['General', '0', '0.00', '#,##0', '#,##0.00', '0%', '0.00%', '0.00E+00', '# ?/?', '# ??/??', '#,##0 "€"', '#,##0.00 [$€-40C]', '0.00;[Red]-0.00', '#,##0;(#,##0)', '0;-0;"zéro"', '0.0,', '#,##0,,', '00000', '"Total : "0', '_-* #,##0.00 €_-;-* #,##0.00 €_-;_-* "-"?? €_-;_-@_-', '+0;-0;0', '[>=1000]#,##0;0.0', '[Blue]0', '0_)', '#.##', '?.??'].map((f): [number, string] => [v, f]),
  ),
  ...[45292, 45383.5, 45444.75, 0.5, 1.5, 61, 366].flatMap((v): [number, string][] =>
    ['dd/mm/yyyy', 'd/m/yy', 'dd mmm yyyy', 'dddd d mmmm yyyy', 'ddd', 'mmmm', 'mmm-yy', 'yyyy', 'yy', 'hh:mm', 'h:mm:ss', 'hh:mm AM/PM', '[h]:mm', '[m]', '[s]', 'mm:ss', 'mm:ss.0', 'dd/mm/yyyy hh:mm', 'yyyy-mm-dd', 'm', 'mmmmm', 'd'].map((f): [number, string] => [v, f]),
  ),
  ['Texte', '@'], ['Texte', '"> "@'], ['Texte', '0'], ['Texte', '0;0;0;"<"@">"'], [12, '@'], [true, '0'], [true, 'General'], [1e15, 'General'], [1e11, 'General'], [123456789012, 'General'], [1.23456789012345, 'General'], [0.000012345, 'General'], [1e-10, 'General'], [1 / 3, 'General'], [100000000000000000000, 'General'], [1234567.891, 'General'],
];

/**
 * Écarts connus et acceptés : présentation (Excel pour Mac écrit « janv », Excel Windows et Grillon « janv. »), formes lues plus
 * souplement que dans Excel français (« >=12.5 »), et cas rares non traités (voir docs/EXCEL-COMPARAISON.md).
 */
export const KNOWN_DIFFERENCES = new Set<string>([
  `COUNTIF(${C},">=12.5")`, `COUNTIFS(${B},"Lyon",${C},">12.5")`, `TEXT(45292,"jjjj d mmmm aaaa")`, `TEXT(45292,"jj mmm aa")`, `UPPER(TEXT(D!$D$2,"mmm"))`,
  `DBCS("a")`, `MATCH(4,${X},-1)`, `ADDRESS(2,3,1,FALSE,"D")`, `AREAS((D!$A$1,D!$C$1))`, `LEFT(${A},1)&"."`, 'IMCSCH("1+i")',
]);

/** Classeur de comparaison : données (D), formules (F, colonne B) et formats (N, colonne C). */
export function referenceWorkbook(): IWorkbookData {
  const rows: (string | number | boolean | null)[][] = [
    ['Nom', 'Ville', 'Note', 'Date', 'Nombre', 'Actif', 'Mélange', 'X', 'Y', 'Formule', 'Note'],
    ['Léa', 'Lyon', 15, 45292, 1.5, true, 1, 1, 2.1, null, '>10'],
    ['Tom', 'Paris', 8, 45323, 2.25, false, 'a', 2, 3.9],
    ['Éloi', 'Lyon', 12.5, 45352, -3, true, null, 3, 6.2],
    ['Zoé', 'Nîmes', 19, 45383, 4, true, 2, 4, 7.8],
    ['Ana', 'Paris', 7, 45413, 0, false, true, 5, 10.1],
    ['Tom', 'Lyon', 12.5, 45444, 100, true, '3', 6, 12.2],
    ['Max', 'Paris', 3, 45474, 1000.75, false, 4.5, 7, 13.8],
  ];
  const data: Record<number, Record<number, object>> = {};
  rows.forEach((r, i) => {
    data[i] = {};
    r.forEach((v, j) => {
      if (v === null) return;
      data[i]![j] = typeof v === 'number' ? { v, t: 2 } : typeof v === 'boolean' ? { v: v ? 1 : 0, t: 3 } : { v, t: 1 };
    });
  });
  data[1]![9] = { f: '=SUM(1,2)', v: 3, t: 2 };
  const formulas: Record<number, Record<number, object>> = {};
  FORMULAS.forEach((f, i) => (formulas[i] = { 0: { v: f, t: 1 }, 1: { f: `=${f}` } }));
  const styles: Record<string, object> = {};
  const formats: Record<number, Record<number, object>> = {};
  NUMBER_FORMATS.forEach(([v, pattern], i) => {
    const id = `n${i}`;
    styles[id] = { n: { pattern } };
    formats[i] = {
      0: { v: pattern, t: 1 },
      2: typeof v === 'number' ? { v, t: 2, s: id } : typeof v === 'boolean' ? { v: v ? 1 : 0, t: 3, s: id } : { v, t: 1, s: id },
    };
  });
  return {
    id: 'ref',
    name: 'reference',
    appVersion: '1',
    locale: 'frFR' as IWorkbookData['locale'],
    styles,
    sheetOrder: ['D', 'F', 'N'],
    sheets: {
      D: { id: 'D', name: 'D', cellData: data },
      F: { id: 'F', name: 'F', cellData: formulas, rowCount: FORMULAS.length + 5, columnData: { 1: { w: 900 } } },
      N: { id: 'N', name: 'N', cellData: formats, rowCount: NUMBER_FORMATS.length + 5, columnData: { 2: { w: 900 } } },
    },
  } as never;
}

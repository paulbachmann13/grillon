# Fonctions de formule

Les formules s'écrivent **comme dans Excel en français** : noms de fonctions français (`SOMME`, `SI`,
`RECHERCHEV`…), « ; » entre les arguments, virgule décimale, `VRAI` / `FAUX`, `{1\2;3\4}` pour une constante
matricielle. Les formules en anglais (`=SUM(A1,B1)`) sont aussi acceptées. Le classeur et le fichier .xlsx gardent
la forme anglaise (celle d'Excel dans le fichier) : la traduction se fait à la saisie et à l'affichage.
Environ 400 fonctions ont un nom français (`src/data/formula-locale.ts`) ; une fonction sans équivalent connu
(`MAP`, `SCAN`, `TOCOL`…) s'écrit avec son nom anglais. Les noms identiques dans les deux langues (`ABS`, `MAX`,
`INDEX`…) ne changent pas. Les résultats s'affichent à la française.

Le calcul est assuré par le moteur de formules d'Univer (Apache-2.0), dans un Web Worker.

## Couverture

Environ **520 fonctions Excel** sont prises en charge, dont :

- **Maths et trigo** : SUM, SUMIF(S), SUMPRODUCT, ROUND/ROUNDUP/ROUNDDOWN, MROUND, CEILING/FLOOR (.MATH, .PRECISE), MOD, POWER, SQRT, EXP, LN, LOG, ABS, RAND, RANDBETWEEN, RANDARRAY, AGGREGATE, SUBTOTAL…
- **Statistiques** : AVERAGE(IF/IFS), COUNT(A/BLANK/IF/IFS), MIN/MAX(IFS), MEDIAN, MODE, STDEV(.S/.P), VAR, PERCENTILE, QUARTILE, RANK, LARGE/SMALL, CORREL, FORECAST(.LINEAR), LINEST, TREND, NORM.DIST… et toutes les lois usuelles.
- **Logique** : IF, IFS, IFERROR, IFNA, AND, OR, NOT, XOR, SWITCH, LET, LAMBDA, MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY.
- **Recherche** : XLOOKUP, XMATCH, VLOOKUP, HLOOKUP, LOOKUP, INDEX, MATCH, OFFSET, INDIRECT, CHOOSE, CHOOSECOLS/ROWS, FILTER, SORT, SORTBY, UNIQUE, TRANSPOSE, HSTACK, VSTACK, TAKE, DROP, EXPAND, TOCOL, TOROW, WRAPROWS/COLS, GROUPBY…
- **Texte** : CONCAT, TEXTJOIN, TEXTSPLIT, TEXTBEFORE/AFTER, LEFT/RIGHT/MID, LEN, FIND/SEARCH, SUBSTITUTE, REPLACE, TEXT, VALUE, TRIM, UPPER/LOWER/PROPER, REGEXEXTRACT/REGEXREPLACE/REGEXTEST…
- **Date et heure** : TODAY, NOW, DATE, DATEDIF, EDATE, EOMONTH, NETWORKDAYS(.INTL), WORKDAY(.INTL), WEEKDAY, WEEKNUM, ISOWEEKNUM, YEARFRAC…
- **Financier** : PMT, IPMT, PPMT, FV, PV, NPV, XNPV, IRR, XIRR, RATE, NPER, SLN, DB, DDB… et les fonctions obligataires.
- **Ingénierie, information, base de données (DSUM…)**, matrices dynamiques (résultats « débordants »).
- **Tableaux structurés** : références `Tableau[Colonne]`, `Tableau[[#Data],[Colonne]]`.

### Ajoutées par Grillon

| Fonction | Rôle |
|---|---|
| `ISO.CEILING(nombre; [précision])` | Arrondi à l'excès, quel que soit le signe |
| `TRIMRANGE(plage)` | Retire les lignes et colonnes vides en bordure d'une plage |

## Non prises en charge

| Fonction(s) | Raison |
|---|---|
| CUBEKPIMEMBER, CUBEMEMBER, CUBEMEMBERPROPERTY, CUBERANKEDMEMBER, CUBESET, CUBESETCOUNT, CUBEVALUE | Connexion à un serveur OLAP |
| CALL, REGISTER.ID, RTD, WEBSERVICE | Appels externes (DLL, serveurs) |
| FORECAST.ETS, FORECAST.ETS.CONFINT, FORECAST.ETS.SEASONALITY, FORECAST.ETS.STAT | Lissage exponentiel triple non implémenté (utiliser FORECAST.LINEAR / TREND) |
| GETPIVOTDATA | Les TCD de Grillon sont des tableaux de valeurs |
| PIVOTBY | Non implémentée (utiliser GROUPBY ou un TCD) |
| AMORDEGRC, EUROCONVERT, INFO, ISOMITTED, PHONETIC, UNION | Rares ou propres à Excel Windows |

Une formule utilisant une de ces fonctions affiche `#NAME?` ; elle est conservée telle quelle
à l'enregistrement et fonctionnera de nouveau à l'ouverture dans Excel.

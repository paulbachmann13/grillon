/**
 * Détection des pertes potentielles lors de l'enregistrement d'un fichier Excel (.xlsx).
 * Analyse le classeur pour identifier les fonctionnalités non prises en charge par Grillon.
 */
import JSZip from 'jszip';

export type LossKind =
  | 'macros'
  | 'threadedComments'
  | 'smartArt'
  | 'unsupportedCharts'
  | 'formulaNames'
  | 'advancedFilters'
  | 'extraPivots'
  | 'excelTimelines';

const ORDER: LossKind[] = [
  'macros',
  'threadedComments',
  'smartArt',
  'unsupportedCharts',
  'formulaNames',
  'advancedFilters',
  'extraPivots',
  'excelTimelines',
];

// Coordonnée d'une zone : cellule (A1), plage de cellules (A1:B2), colonnes (A:B) ou lignes (1:2).
const COORD = '(?:\\$?[A-Za-z]+\\$?[1-9]\\d*(?::\\$?[A-Za-z]+\\$?[1-9]\\d*)?|\\$?[A-Za-z]+:\\$?[A-Za-z]+|\\$?[1-9]\\d*:\\$?[1-9]\\d*)';
// Préfixe de feuille facultatif : 'Nom avec espaces'! ou NomSansEspace!
const SHEET = "(?:'(?:[^']|'')+'!|[\\p{L}\\p{N}_.]+!)";
const ZONE = `(?:${SHEET})?${COORD}`;
const SIMPLE_REF_RE = new RegExp(`^\\s*${ZONE}(?:\\s*,\\s*${ZONE})*\\s*$`, 'u');

function unescapeXml(s: string): string {
  return s
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');
}

function getAttr(tagAttrs: string, name: string): string | undefined {
  const m = new RegExp(`(?:^|\\s)${name}\\s*=\\s*["']([^"']*)["']`, 'i').exec(tagAttrs);
  return m?.[1];
}

const MAX_SHEET_TEXT = 30 * 1024 * 1024;
const uncompressedSize = (file: JSZip.JSZipObject): number => (file as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? 0;

function normalizePath(path: string): string {
  return path.replace(/\\/g, '/').replace(/^\//, '');
}

/** Vérifie si le classeur contient des noms définis portant des formules ou constantes. */
function hasFormulaDefinedNames(rawXml: string): boolean {
  const xml = rawXml.replace(/<!--[\s\S]*?-->/g, '');
  const defNameRe = /<(?:\w+:)?definedName\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?definedName>/gi;
  let m: RegExpExecArray | null;
  while ((m = defNameRe.exec(xml)) !== null) {
    const attrs = m[1];
    const name = getAttr(attrs, 'name');
    if (name && /^_xlnm\./i.test(name)) continue;

    const hidden = getAttr(attrs, 'hidden');
    if (hidden === '1' || hidden === 'true') continue;

    const content = unescapeXml(m[2]).trim();
    if (!content || content.includes('#REF!')) continue;

    if (!SIMPLE_REF_RE.test(content)) {
      return true;
    }
  }
  return false;
}

/**
 * Détecte les fonctionnalités présentes dans le fichier .xlsx qui seraient perdues à l'enregistrement.
 * Renvoie un tableau contenant au plus une fois chaque type, dans l'ordre strict de LossKind.
 */
export async function detectLosses(data: ArrayBuffer | Uint8Array): Promise<LossKind[]> {
  let zip: JSZip;
  try {
    if (!data) return [];
    zip = await JSZip.loadAsync(data);
  } catch {
    return [];
  }

  try {
    const detected = new Set<LossKind>();
    const files = Object.values(zip.files).filter((f) => !f.dir);

    const fileEntries: { path: string; file: JSZip.JSZipObject }[] = files.map((file) => ({
      path: normalizePath(file.name),
      file,
    }));

    // 1. Macros VBA
    if (fileEntries.some((e) => e.path.toLowerCase() === 'xl/vbaproject.bin')) {
      detected.add('macros');
    }

    // 2. Commentaires hiérarchisés
    if (fileEntries.some((e) => /^xl\/threadedcomments\/[^/]+\.xml$/i.test(e.path))) {
      detected.add('threadedComments');
    }

    // 3. SmartArt
    if (fileEntries.some((e) => /^xl\/diagrams\/.+/i.test(e.path))) {
      detected.add('smartArt');
    }

    // 4. Graphiques non pris en charge
    const chartEntries = fileEntries.filter((e) => /^xl\/charts\/chart[^/]*\.xml$/i.test(e.path));
    for (const { path, file } of chartEntries) {
      const xml = await file.async('string');
      if (/^xl\/charts\/chartex[^/]*\.xml$/i.test(path)) {
        if (/<(?:\w+:)?series\b[^>]*\blayoutId\s*=\s*["'](?:paretoLine|treemap|sunburst|regionMap)["']/i.test(xml)) {
          detected.add('unsupportedCharts');
          break;
        }
      } else {
        if (/<(?:\w+:)?(?:bar3DChart|line3DChart|pie3DChart|area3DChart|surface3DChart|surfaceChart|stockChart)\b/i.test(xml)) {
          detected.add('unsupportedCharts');
          break;
        }
      }
    }

    // 5. Noms définis avec formules
    const wbEntry = fileEntries.find((e) => e.path.toLowerCase() === 'xl/workbook.xml');
    if (wbEntry) {
      const xml = await wbEntry.file.async('string');
      if (hasFormulaDefinedNames(xml)) {
        detected.add('formulaNames');
      }
    }

    // 6. Filtres avancés
    // Feuilles énormes (des centaines de milliers de lignes) : pas relues en texte, la mémoire doublerait.
    const worksheetEntries = fileEntries.filter((e) => /^xl\/worksheets\/[^/]+\.xml$/i.test(e.path) && uncompressedSize(e.file) <= MAX_SHEET_TEXT);
    for (const { file } of worksheetEntries) {
      const xml = await file.async('string');
      if (/<(?:\w+:)?(?:colorFilter|top10|dynamicFilter|iconFilter)\b/i.test(xml)) {
        detected.add('advancedFilters');
        break;
      }
    }

    // 7. Tableaux croisés dynamiques multiples sur une même feuille
    const relsEntries = fileEntries.filter((e) => /^xl\/worksheets\/_rels\/[^/]+\.xml\.rels$/i.test(e.path));
    for (const { file } of relsEntries) {
      const xml = await file.async('string');
      const matches = xml.match(/<(?:\w+:)?Relationship\b[^>]*\bType\s*=\s*["'][^"']*\/pivotTable["']/gi);
      if (matches && matches.length >= 2) {
        detected.add('extraPivots');
        break;
      }
    }

    // 8. Chronologies Excel
    if (fileEntries.some((e) => /^xl\/(timelines|timelinecaches)\/[^/]+\.xml$/i.test(e.path))) {
      detected.add('excelTimelines');
    }

    return ORDER.filter((kind) => detected.has(kind));
  } catch {
    return [];
  }
}

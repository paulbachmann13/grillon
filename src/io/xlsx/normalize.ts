/**
 * Remise en forme d'un .xlsx avant lecture. Excel, LibreOffice et la plupart des logiciels écrivent le format de la même
 * façon, mais d'autres (générateurs de rapports, .NET, Hancom…) prennent des libertés que la norme permet et qu'ExcelJS
 * ne comprend pas : balises préfixées (<x:worksheet>), feuilles nommées « sheet.xml », lignes et cellules sans numéro,
 * propriétés du document inhabituelles. On les ramène à la forme courante (vérifié sur les fichiers de test d'Apache POI).
 */
import type JSZip from 'jszip';
import { readRelationships, resolveRelTarget } from './ooxml';
import { columnToLetters, fromA1 } from './units';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS_TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types';
const DEFAULT_NAMESPACES = new Set([NS_MAIN, NS_RELS, NS_TYPES]);
const REL_WORKSHEET = /\/worksheet$/;

/** <x:worksheet xmlns:x="…main"> → <worksheet xmlns="…main"> (seulement si la racine est préfixée et sans espace par défaut). */
export function unprefixRoot(xml: string): string {
  const root = /<([A-Za-z_][\w.-]*):([\w.-]+)\b([^>]*)>/.exec(xml.replace(/^﻿/, '').replace(/^<\?[^>]*\?>\s*/, ''));
  if (!root) return xml;
  const prefix = root[1];
  const ns = new RegExp(`\\sxmlns:${prefix}="([^"]*)"`).exec(root[3])?.[1];
  if (!ns || !DEFAULT_NAMESPACES.has(ns) || /\sxmlns="/.test(root[3])) return xml;
  return xml
    .replace(new RegExp(`<(/?)${prefix}:`, 'g'), '<$1')
    .replace(new RegExp(`\\sxmlns:${prefix}="${ns.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&')}"`), ` xmlns="${ns}"`);
}

/** Numérote les lignes et les cellules qui n'ont pas d'attribut r (permis par la norme : elles se suivent). */
export function numberRowsAndCells(xml: string): string {
  if (!/<row(?=[\s>/])(?![^>]*\sr=")/.test(xml) && !/<c(?=[\s>/])(?![^>]*\sr=")/.test(xml)) return xml;
  let row = 0;
  let col = -1;
  return xml.replace(/<(row|c)(?=[\s>/])([^>]*?)(\/?)>/g, (tag, name: string, attrs: string, close: string) => {
    const r = /\sr="([^"]*)"/.exec(attrs)?.[1];
    if (name === 'row') {
      row = r ? Number(r) : row + 1;
      col = -1;
      return r ? tag : `<row r="${row}"${attrs}${close}>`;
    }
    if (r) {
      try {
        col = fromA1(r).col;
      } catch {
        col++;
      }
      return tag;
    }
    col++;
    return `<c r="${columnToLetters(col)}${row}"${attrs}${close}>`;
  });
}

/**
 * Remet le paquet dans la forme courante (avant toute lecture, par Grillon comme par ExcelJS).
 * Renvoie vrai si quelque chose a changé.
 */
export async function normalizePackage(zip: JSZip): Promise<boolean> {
  let changed = false;
  // Balises préfixées : relations, types de contenu, classeur, feuilles, chaînes, styles.
  for (const path of Object.keys(zip.files)) {
    if (zip.files[path].dir || !/(\.rels|\[Content_Types\]\.xml|^xl\/(workbook|styles|sharedStrings)\.xml|^xl\/worksheets\/[^/]+\.xml)$/.test(path)) continue;
    const xml = await zip.file(path)!.async('string');
    let out = unprefixRoot(xml);
    if (path.startsWith('xl/worksheets/') && path.endsWith('.xml')) out = numberRowsAndCells(out);
    if (out !== xml) {
      zip.file(path, out);
      changed = true;
    }
  }
  // Feuilles qui ne s'appellent pas « sheetN.xml » (ExcelJS ne reconnaît que ce nom).
  const relsPath = 'xl/_rels/workbook.xml.rels';
  const rels = await readRelationships(zip, 'xl/workbook.xml');
  const used = new Set(Object.keys(zip.files).filter((p) => /^xl\/worksheets\/sheet\d+\.xml$/.test(p)));
  let relsXml = (await zip.file(relsPath)?.async('string')) ?? '';
  let types = (await zip.file('[Content_Types].xml')?.async('string')) ?? '';
  let n = 1;
  for (const rel of rels.filter((r) => REL_WORKSHEET.test(r.Type))) {
    const path = resolveRelTarget('xl/workbook.xml', rel.Target);
    if (/^xl\/worksheets\/sheet\d+\.xml$/.test(path) || !zip.file(path)) continue;
    while (used.has(`xl/worksheets/sheet${n}.xml`)) n++;
    const next = `xl/worksheets/sheet${n}.xml`;
    used.add(next);
    zip.file(next, await zip.file(path)!.async('string'));
    zip.remove(path);
    const oldRels = `${path.slice(0, path.lastIndexOf('/'))}/_rels/${path.slice(path.lastIndexOf('/') + 1)}.rels`;
    const sheetRels = zip.file(oldRels);
    if (sheetRels) {
      zip.file(`xl/worksheets/_rels/sheet${n}.xml.rels`, await sheetRels.async('string'));
      zip.remove(oldRels);
    }
    relsXml = relsXml.replace(new RegExp(`(<Relationship\\b[^>]*\\bId="${rel.Id}"[^>]*\\bTarget=")[^"]*"`), `$1worksheets/sheet${n}.xml"`).replace(new RegExp(`(<Relationship\\b[^>]*\\bTarget=")[^"]*("[^>]*\\bId="${rel.Id}")`), `$1worksheets/sheet${n}.xml$2`);
    types = types.replace(`PartName="/${path}"`, `PartName="/${next}"`);
    changed = true;
  }
  if (changed) {
    if (relsXml) zip.file(relsPath, relsXml);
    if (types) zip.file('[Content_Types].xml', types);
  }
  // Noms de feuille de plus de 31 caractères : ExcelJS les tronque, et deux noms tronqués peuvent se confondre.
  const wbXml = await zip.file('xl/workbook.xml')?.async('string');
  if (wbXml) {
    const seen = new Set<string>();
    const out = wbXml.replace(/(<sheet\b[^>]*\bname=")([^"]*)(")/g, (m, a: string, name: string, b: string) => {
      let short = name.slice(0, 31);
      for (let i = 2; seen.has(short.toLowerCase()); i++) short = `${name.slice(0, 31 - String(i).length - 1)}~${i}`;
      seen.add(short.toLowerCase());
      return short === name ? m : `${a}${short}${b}`;
    });
    if (out !== wbXml) {
      zip.file('xl/workbook.xml', out);
      changed = true;
    }
  }
  return changed;
}

/**
 * Retire ce que Grillon lit lui-même et qu'ExcelJS lit mal ou pas du tout (tableaux, filtres, propriétés du document) :
 * à appeler après la lecture par Grillon, juste avant ExcelJS.
 */
export async function stripForExcelJs(zip: JSZip, sheetPaths: string[]): Promise<boolean> {
  let changed = false;
  // Propriétés du document (auteur, société…) : inutiles à Grillon ; retirées si elles ne sont pas sous la forme qu'attend ExcelJS.
  const expected: Record<string, RegExp> = { 'docProps/app.xml': /<Properties[\s>]/, 'docProps/core.xml': /<cp:coreProperties[\s>]/ };
  for (const [path, root] of Object.entries(expected)) {
    const xml = await zip.file(path)?.async('string');
    if (xml === undefined) continue;
    const head = xml.slice(0, 600);
    // ExcelJS ne connaît que les éléments préfixés (dc:, cp:…) : un élément sans préfixe le fait échouer.
    const odd = !root.test(head) || (path.endsWith('core.xml') && /<(?!\/|\?|!)(?![\w]+:)[A-Za-z]/.test(xml.replace(/^[\s\S]*?<cp:coreProperties[^>]*>/, '')));
    if (odd) {
      zip.remove(path);
      changed = true;
    }
  }
  // Tableaux : ExcelJS lit toutes les parties xl/tables/ (et échoue sur certains filtres) ; Grillon les lit lui-même.
  for (const path of Object.keys(zip.files).filter((p) => p.startsWith('xl/tables/'))) {
    zip.remove(path);
    changed = true;
  }
  for (const path of sheetPaths) {
    const relsPath = `${path.slice(0, path.lastIndexOf('/'))}/_rels/${path.slice(path.lastIndexOf('/') + 1)}.rels`;
    const rels = await zip.file(relsPath)?.async('string');
    if (rels && /\/table"/.test(rels)) zip.file(relsPath, rels.replace(/<Relationship\b[^>]*\/relationships\/table"[^>]*\/>/g, ''));
    const xml = await zip.file(path)?.async('string');
    if (!xml) continue;
    const out = xml.replace(/<tableParts\b[^>]*\/>|<tableParts\b[\s\S]*?<\/tableParts>/, '').replace(/<autoFilter\b[^>]*\/>|<autoFilter\b[\s\S]*?<\/autoFilter>/, '')
      .replace(/<mergeCells\b[^>]*\/>|<mergeCells\b[\s\S]*?<\/mergeCells>/, '');
    if (out !== xml) {
      zip.file(path, out);
      changed = true;
    }
  }
  return changed;
}

/**
 * Notes (commentaires de cellule) : ExcelJS (cell.note) ⇄ Univer
 * (ressource SHEET_NOTE_PLUGIN = { [sheetId]: { [ligne]: { [colonne]: { note, width, height } } } }).
 */
import type ExcelJS from 'exceljs';
import type JSZip from 'jszip';
import { extractAll, readRelationships, resolveRelTarget } from './ooxml';
import { fromA1 } from './units';

export const NOTE_RESOURCE = 'SHEET_NOTE_PLUGIN';

export interface UniverNote {
  note: string;
  width?: number;
  height?: number;
  id?: string;
  row?: number;
  col?: number;
}

export type SheetNotes = Record<number, Record<number, UniverNote>>;
export type UniverNoteResource = Record<string, SheetNotes>;

function noteText(note: ExcelJS.Comment | string | undefined): string {
  if (!note) return '';
  if (typeof note === 'string') return note;
  const texts = (note as { texts?: { text: string }[] }).texts;
  return texts ? texts.map((t) => t.text).join('') : '';
}

export function importNotes(ws: ExcelJS.Worksheet): SheetNotes {
  const out: SheetNotes = {};
  ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const cells = (row as unknown as { _cells: (ExcelJS.Cell | undefined)[] })._cells ?? [];
    cells.forEach((cell, c) => {
      const text = noteText(cell?.note as ExcelJS.Comment | string | undefined).trim();
      if (!text) return;
      const r = rowNumber - 1;
      (out[r] ??= {})[c] = { note: text, width: 216, height: 92, row: r, col: c };
    });
  });
  return out;
}

const unescape = (s: string) =>
  s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_m, n: string) => String.fromCodePoint(Number(n))).replace(/&#x([0-9a-f]+);/gi, (_m, n: string) => String.fromCodePoint(parseInt(n, 16))).replace(/&amp;/g, '&');

/**
 * Notes d'une feuille lues directement dans xl/commentsN.xml : ExcelJS ne les rattache qu'aux cellules présentes dans
 * le fichier, et perdait celles des cellules vides (fréquent : une note sur un en-tête vide, une case libre…).
 */
export async function readNotes(zip: JSZip, sheetPath: string): Promise<SheetNotes> {
  const out: SheetNotes = {};
  for (const rel of (await readRelationships(zip, sheetPath)).filter((r) => r.Type.endsWith('/comments'))) {
    const xml = await zip.file(resolveRelTarget(sheetPath, rel.Target))?.async('string');
    if (!xml) continue;
    for (const c of extractAll(xml, 'comment')) {
      const ref = /\sref="([^"]+)"/.exec(c.xml.slice(0, c.xml.indexOf('>') + 1))?.[1];
      if (!ref) continue;
      let pos: { row: number; col: number };
      try {
        pos = fromA1(ref.split(':')[0]);
      } catch {
        continue;
      }
      const text = unescape([...c.xml.matchAll(/<(?:\w+:)?t\b[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/g)].map((m) => m[1]).join('')).trim();
      if (!text) continue;
      (out[pos.row] ??= {})[pos.col] = { note: text, width: 216, height: 92, row: pos.row, col: pos.col };
    }
  }
  return out;
}

export function exportNotes(ws: ExcelJS.Worksheet, notes: SheetNotes | undefined): void {
  for (const [r, cols] of Object.entries(notes ?? {})) {
    for (const [c, n] of Object.entries(cols ?? {})) {
      if (!n?.note) continue;
      ws.getRow(Number(r) + 1).getCell(Number(c) + 1).note = n.note;
    }
  }
}

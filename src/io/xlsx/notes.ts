/**
 * Notes (commentaires de cellule) : ExcelJS (cell.note) ⇄ Univer
 * (ressource SHEET_NOTE_PLUGIN = { [sheetId]: { [ligne]: { [colonne]: { note, width, height } } } }).
 */
import type ExcelJS from 'exceljs';

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

export function exportNotes(ws: ExcelJS.Worksheet, notes: SheetNotes | undefined): void {
  for (const [r, cols] of Object.entries(notes ?? {})) {
    for (const [c, n] of Object.entries(cols ?? {})) {
      if (!n?.note) continue;
      ws.getRow(Number(r) + 1).getCell(Number(c) + 1).note = n.note;
    }
  }
}

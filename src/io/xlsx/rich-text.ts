/**
 * Texte enrichi (plusieurs styles dans une même cellule) :
 * ExcelJS (richText: [{ text, font }]) ⇄ document Univer (cell.p).
 *
 * Dans un document Univer, le texte est un « dataStream » où « \r » sépare les paragraphes
 * et qui se termine par « \r\n » ; les styles sont des « textRuns » [st, ed[.
 */
import type ExcelJS from 'exceljs';
import type { ICellData, IStyleData } from '../../model/univer-types';
import { styleFromUniver, styleToUniver } from './styles';

type CellDocument = Exclude<ICellData['p'], null | undefined | void>;

interface TextRun {
  st: number;
  ed: number;
  ts?: IStyleData;
}

/** Propriétés de police seulement (les textRuns n'acceptent pas bordures, fond…). */
const FONT_KEYS = ['ff', 'fs', 'it', 'bl', 'ul', 'st', 'cl', 'va'] as const;

function pickFont(s: IStyleData | null): IStyleData | undefined {
  if (!s) return undefined;
  const out: IStyleData = {};
  for (const k of FONT_KEYS) if (s[k] !== undefined) (out as Record<string, unknown>)[k] = s[k];
  return Object.keys(out).length ? out : undefined;
}

export function richTextToDocument(runs: ExcelJS.RichText[], theme: string[]): CellDocument {
  let dataStream = '';
  const textRuns: TextRun[] = [];
  const paragraphs: { startIndex: number }[] = [];
  for (const run of runs) {
    const text = (run.text ?? '').replace(/\r\n?/g, '\n');
    const st = dataStream.length;
    for (const ch of text) {
      if (ch === '\n') {
        paragraphs.push({ startIndex: dataStream.length });
        dataStream += '\r';
      } else dataStream += ch;
    }
    const ts = pickFont(styleToUniver({ font: run.font }, theme));
    if (ts && dataStream.length > st) textRuns.push({ st, ed: dataStream.length, ts });
  }
  paragraphs.push({ startIndex: dataStream.length });
  dataStream += '\r\n';
  return {
    id: 'd',
    documentStyle: {},
    body: { dataStream, textRuns, paragraphs },
  } as unknown as CellDocument;
}

/** Texte brut d'un document de cellule (sauts de paragraphe → « \n »). */
export function documentToPlainText(doc: CellDocument): string {
  const stream = doc.body?.dataStream ?? '';
  return stream.replace(/\r\n$/, '').replace(/\r/g, '\n');
}

/** Document Univer → richText ExcelJS (null s'il n'y a aucun style à conserver). */
export function documentToRichText(doc: CellDocument): ExcelJS.RichText[] | null {
  const stream = (doc.body?.dataStream ?? '').replace(/\r\n$/, '');
  const runs = ((doc.body as { textRuns?: TextRun[] } | undefined)?.textRuns ?? [])
    .filter((r) => r.ed > r.st)
    .sort((a, b) => a.st - b.st);
  if (!runs.length) return null;
  const out: ExcelJS.RichText[] = [];
  const push = (from: number, to: number, ts?: IStyleData) => {
    const text = stream.slice(from, Math.min(to, stream.length)).replace(/\r/g, '\n');
    if (!text) return;
    const font = ts ? (styleFromUniver(ts)?.font as Partial<ExcelJS.Font> | undefined) : undefined;
    out.push(font ? { text, font } : { text });
  };
  let pos = 0;
  for (const r of runs) {
    if (r.st > pos) push(pos, r.st);
    push(Math.max(r.st, pos), r.ed, r.ts);
    pos = Math.max(pos, r.ed);
  }
  if (pos < stream.length) push(pos, stream.length);
  return out.length ? out : null;
}

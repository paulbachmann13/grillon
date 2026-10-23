/**
 * Polices de remplacement : Calibri (police par défaut d'Excel) n'est pas installée sur un Mac sans Office.
 * Sans elle, la grille (un canevas) retombe sur Times. On déclare donc « Calibri » : la vraie si elle est
 * installée (`local`), sinon Carlito, libre (OFL) et de mêmes dimensions — comme le fait LibreOffice.
 */
import latin400 from '@fontsource/carlito/files/carlito-latin-400-normal.woff2';
import latin400i from '@fontsource/carlito/files/carlito-latin-400-italic.woff2';
import latin700 from '@fontsource/carlito/files/carlito-latin-700-normal.woff2';
import latin700i from '@fontsource/carlito/files/carlito-latin-700-italic.woff2';
import ext400 from '@fontsource/carlito/files/carlito-latin-ext-400-normal.woff2';
import ext400i from '@fontsource/carlito/files/carlito-latin-ext-400-italic.woff2';
import ext700 from '@fontsource/carlito/files/carlito-latin-ext-700-normal.woff2';
import ext700i from '@fontsource/carlito/files/carlito-latin-ext-700-italic.woff2';

const LATIN = 'U+0000-00FF,U+0131,U+0152-0153,U+02BB-02BC,U+02C6,U+02DA,U+02DC,U+0304,U+0308,U+0329,U+2000-206F,U+20AC,U+2122,U+2191,U+2193,U+2212,U+2215,U+FEFF,U+FFFD';
const LATIN_EXT = 'U+0100-02BA,U+02BD-02C5,U+02C7-02CC,U+02CE-02D7,U+02DD-02FF,U+0304,U+0308,U+0329,U+1D00-1DBF,U+1E00-1E9F,U+1EF2-1EFF,U+2020,U+20A0-20AB,U+20AD-20C0,U+2113,U+2C60-2C7F,U+A720-A7FF';

const FILES = [
  { weight: '400', style: 'normal', latin: latin400, ext: ext400 },
  { weight: '400', style: 'italic', latin: latin400i, ext: ext400i },
  { weight: '700', style: 'normal', latin: latin700, ext: ext700 },
  { weight: '700', style: 'italic', latin: latin700i, ext: ext700i },
];

/** Déclare Calibri / Calibri Light / Carlito et attend leur chargement (la grille les dessine dès l'ouverture). */
export async function installFallbackFonts(): Promise<void> {
  if (typeof FontFace === 'undefined' || !document.fonts) return;
  const faces: FontFace[] = [];
  for (const family of ['Calibri', 'Calibri Light', 'Carlito']) {
    for (const f of FILES) {
      for (const [url, range] of [[f.latin, LATIN], [f.ext, LATIN_EXT]] as const) {
        const local = family === 'Carlito' ? '' : `local("${family}${f.weight === '700' ? ' Bold' : ''}${f.style === 'italic' ? ' Italic' : ''}"), `;
        faces.push(new FontFace(family, `${local}url(${url}) format("woff2")`, { weight: f.weight, style: f.style, unicodeRange: range }));
      }
    }
  }
  for (const face of faces) document.fonts.add(face);
  // Les styles courants ; les autres se chargent à la demande.
  await Promise.race([
    Promise.allSettled(['11pt Calibri', 'bold 11pt Calibri', 'italic 11pt Calibri'].map((f) => document.fonts.load(f, 'Aé'))),
    new Promise((r) => setTimeout(r, 2000)),
  ]);
}

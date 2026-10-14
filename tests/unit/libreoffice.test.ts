/**
 * Vérifie que LibreOffice (headless) ouvre sans erreur les fichiers exportés par Grillon
 * et y lit les mêmes valeurs. Ignoré si LibreOffice n'est pas installé (il l'est en CI).
 */
import { execFileSync, spawnSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { exportXlsx } from '../../src/io/xlsx/export';
import { importXlsx } from '../../src/io/xlsx/import';
import { buildFixture, FIXTURES } from '../fixtures/build';
import { normalize } from './helpers';

function findSoffice(): string | null {
  const candidates = [process.env.SOFFICE, 'soffice', 'libreoffice', 'C:/Program Files/LibreOffice/program/soffice.exe', '/Applications/LibreOffice.app/Contents/MacOS/soffice'];
  for (const c of candidates) {
    if (!c) continue;
    if (c.includes('/') && !existsSync(c)) continue;
    const r = spawnSync(c, ['--version'], { encoding: 'utf8' });
    if (r.status === 0) return c;
  }
  return null;
}

const soffice = findSoffice();

describe.skipIf(!soffice)('LibreOffice ouvre les exports Grillon', () => {
  for (const name of FIXTURES.filter((n) => n !== 'large')) {
    it(`${name}.xlsx : conversion LibreOffice → xlsx sans perte de valeurs`, async () => {
      const dir = mkdtempSync(join(tmpdir(), 'grillon-lo-'));
      const original = await importXlsx(await buildFixture(name));
      const ours = join(dir, `${name}.xlsx`);
      writeFileSync(ours, await exportXlsx(original));
      const outDir = join(dir, 'out');
      // LibreOffice échoue (code ≠ 0 ou pas de fichier produit) si le classeur est corrompu.
      execFileSync(soffice!, ['--headless', '--norestore', '--convert-to', 'xlsx', '--outdir', outDir, ours], { timeout: 120_000 });
      const converted = join(outDir, `${name}.xlsx`);
      expect(existsSync(converted)).toBe(true);
      const reread = await importXlsx(readFileSync(converted));
      const a = normalize(original);
      const b = normalize(reread);
      expect(b.map((s) => s.name)).toEqual(a.map((s) => s.name));
      // LibreOffice réécrit les constantes logiques en fonctions : FALSE → FALSE(), et une cellule VRAI en =TRUE().
      const sameFormula = (f?: string) => f?.replace(/\b(TRUE|FALSE)\(\)/g, '$1');
      for (let i = 0; i < a.length; i++) {
        for (const [ref, cell] of Object.entries(a[i].cells) as [string, { v?: unknown; t?: number; f?: string }][]) {
          const other = b[i].cells[ref] as { v?: unknown; t?: number; f?: string } | undefined;
          const where = `${a[i].name}!${ref}`;
          if (cell.t === 3 && !cell.f && other?.f) {
            expect(sameFormula(other.f), where).toBe(cell.v ? '=TRUE' : '=FALSE');
            continue;
          }
          if (cell.v !== undefined && typeof cell.v !== 'number') expect(other?.v, where).toEqual(cell.v);
          if (typeof cell.v === 'number') expect(other?.v as number, where).toBeCloseTo(cell.v, 6);
          if (cell.f) expect(sameFormula(other?.f), where).toBe(cell.f);
        }
        expect(b[i].merges).toEqual(a[i].merges);
      }
    });
  }
});

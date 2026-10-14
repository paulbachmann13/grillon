/** Classeur de 100 000 lignes pour mesurer les performances (non versionné). */
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { buildLarge } from '../tests/fixtures/build.ts';

const out = join(import.meta.dirname, '..', 'tests', 'fixtures', 'out');
mkdirSync(out, { recursive: true });
const wb = await buildLarge(100_000);
writeFileSync(join(out, 'large-100k.xlsx'), new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer));
console.log('large-100k.xlsx écrit');

/**
 * Écrit les classeurs de test dans tests/fixtures/out/ (utilisés par Playwright et le contrôle LibreOffice).
 * Usage : npm run fixtures
 */
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { buildFixture, FIXTURES } from '../tests/fixtures/build.ts';

const out = join(import.meta.dirname, '..', 'tests', 'fixtures', 'out');
mkdirSync(out, { recursive: true });
for (const name of FIXTURES) {
  const buf = await buildFixture(name);
  writeFileSync(join(out, `${name}.xlsx`), buf);
  console.log(`${name}.xlsx (${Math.round(buf.byteLength / 1024)} Ko)`);
}

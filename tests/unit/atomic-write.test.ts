import { describe, expect, it } from 'vitest';
import { tempPathFor, writeAtomic } from '../../src/io/tauri-files';

/** Faux disque : un dictionnaire chemin → contenu, avec des pannes à la demande. */
function fakeFs(files: Record<string, string>, fail: { write?: (p: string) => boolean; rename?: boolean } = {}) {
  const dec = new TextDecoder();
  return {
    files,
    async writeFile(p: string, d: Uint8Array) {
      if (fail.write?.(p)) throw new Error('disque plein');
      files[p] = dec.decode(d);
    },
    async rename(from: string, to: string) {
      if (fail.rename) throw new Error('refusé');
      files[to] = files[from]!;
      delete files[from];
    },
    async remove(p: string) {
      delete files[p];
    },
  };
}
const bytes = (s: string) => new TextEncoder().encode(s);

describe('enregistrement sûr (application Mac)', () => {
  it('le fichier temporaire est à côté de l’original', () => {
    expect(tempPathFor('/Users/x/Notes élèves.xlsx')).toBe('/Users/x/Notes élèves.xlsx.grillon-tmp');
  });

  it('écrit une copie puis la renomme : l’original est remplacé, aucun reste', async () => {
    const fs = fakeFs({ '/d/a.xlsx': 'ancien' });
    await writeAtomic('/d/a.xlsx', bytes('nouveau'), fs);
    expect(fs.files).toEqual({ '/d/a.xlsx': 'nouveau' });
  });

  it('disque plein : l’erreur remonte et l’original reste intact', async () => {
    const fs = fakeFs({ '/d/a.xlsx': 'ancien' }, { write: (p) => p.endsWith('.grillon-tmp') });
    await expect(writeAtomic('/d/a.xlsx', bytes('nouveau'), fs)).rejects.toThrow('disque plein');
    expect(fs.files).toEqual({ '/d/a.xlsx': 'ancien' });
  });

  it('renommage refusé : l’original reste intact et la copie est retirée', async () => {
    const fs = fakeFs({ '/d/a.xlsx': 'ancien' }, { rename: true });
    await expect(writeAtomic('/d/a.xlsx', bytes('nouveau'), fs)).rejects.toThrow('refusé');
    expect(fs.files).toEqual({ '/d/a.xlsx': 'ancien' });
  });
});

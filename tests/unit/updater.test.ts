import { describe, expect, it, vi } from 'vitest';
import { isTranslocatedPath, runUpdateCheck, type AvailableUpdate, type Deps } from '../../src/app/updater';

function deps(overrides: Partial<Deps> = {}, update: AvailableUpdate | null = null) {
  const log: string[] = [];
  const d: Deps = {
    check: async () => update,
    currentVersion: async () => '1.11.0',
    ask: vi.fn(async (m: string) => (log.push(`ask:${m}`), true)),
    tell: vi.fn(async (m: string) => void log.push(`tell:${m}`)),
    status: vi.fn((m: string) => void log.push(`status:${m}`)),
    relaunch: vi.fn(async () => void log.push('relaunch')),
    beforeRestart: vi.fn(async () => void log.push('save')),
    isTranslocated: async () => false,
    ...overrides,
  };
  return { d, log };
}

const update = (install = async (p: (r: number, t: number | undefined) => void) => (p(50, 100), p(100, 100))): AvailableUpdate => ({ version: '1.12.0', install });

describe('mise à jour automatique', () => {
  it('à jour : silencieux au démarrage, message à la demande', async () => {
    const quiet = deps();
    expect(await runUpdateCheck(quiet.d, true)).toBe('none');
    expect(quiet.d.tell).not.toHaveBeenCalled();
    const asked = deps();
    expect(await runUpdateCheck(asked.d, false)).toBe('none');
    expect(asked.log).toEqual(['tell:Vous utilisez la dernière version de Grillon (1.11.0).']);
  });

  it('échec de la vérification (hors connexion) : silencieux au démarrage, message à la demande', async () => {
    const boom = () => Promise.reject(new Error('réseau'));
    const quiet = deps({ check: boom });
    expect(await runUpdateCheck(quiet.d, true)).toBe('error');
    expect(quiet.d.tell).not.toHaveBeenCalled();
    const asked = deps({ check: boom });
    expect(await runUpdateCheck(asked.d, false)).toBe('error');
    expect(asked.log[0]).toContain('Impossible de vérifier les mises à jour : réseau');
  });

  it('nouvelle version acceptée : téléchargement, sauvegarde, puis redémarrage', async () => {
    const { d, log } = deps({}, update());
    expect(await runUpdateCheck(d, true)).toBe('installed');
    expect(log[0]).toContain('La version 1.12.0 de Grillon est disponible (vous avez la 1.11.0)');
    expect(log).toContain('status:Téléchargement de la mise à jour : 50 %');
    expect(log.slice(-3)).toEqual(['status:Mise à jour installée : redémarrage…', 'save', 'relaunch']);
  });

  it('nouvelle version refusée : rien n’est installé', async () => {
    const install = vi.fn(async () => undefined);
    const { d } = deps({ ask: async () => false }, update(install));
    expect(await runUpdateCheck(d, false)).toBe('declined');
    expect(install).not.toHaveBeenCalled();
    expect(d.relaunch).not.toHaveBeenCalled();
  });

  it('échec de l’installation : message d’erreur, pas de redémarrage', async () => {
    const { d, log } = deps({}, update(async () => Promise.reject(new Error('signature invalide'))));
    expect(await runUpdateCheck(d, false)).toBe('error');
    expect(log.some((l) => l.includes('La mise à jour a échoué : signature invalide'))).toBe(true);
    expect(d.relaunch).not.toHaveBeenCalled();
  });

  it('une sauvegarde qui échoue n’empêche pas le redémarrage', async () => {
    const { d } = deps({ beforeRestart: () => Promise.reject(new Error('disque plein')) }, update());
    expect(await runUpdateCheck(d, true)).toBe('installed');
    expect(d.relaunch).toHaveBeenCalled();
  });

  it('lancée depuis le .dmg ou Téléchargements (lecture seule) : explique quoi faire, sans tenter l’installation', async () => {
    const install = vi.fn(async () => undefined);
    const { d, log } = deps({ isTranslocated: async () => true }, update(install));
    expect(await runUpdateCheck(d, true)).toBe('error');
    expect(install).not.toHaveBeenCalled();
    expect(d.ask).not.toHaveBeenCalled();
    expect(log[0]).toContain('La version 1.12.0 de Grillon est disponible');
    expect(log[0]).toContain('dossier Applications');
  });

  it('reconnaît le chemin d’une application « translocalisée » par macOS', () => {
    expect(isTranslocatedPath('/private/var/folders/nk/x/T/AppTranslocation/9E56DF82-6/d/Grillon.app/Contents/Resources')).toBe(true);
    expect(isTranslocatedPath('/Applications/Grillon.app/Contents/Resources')).toBe(false);
  });
});

import { beforeEach, describe, expect, it, vi } from 'vitest';
import { isTranslocatedPath, resetUpdateLock, runUpdateCheck, type AvailableUpdate, type Deps } from '../../src/app/updater';

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
  beforeEach(resetUpdateLock);

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

  it('une seule mise à jour à la fois : un deuxième clic pendant le téléchargement n’ouvre pas une deuxième fenêtre', async () => {
    let finish!: () => void;
    const slow = update(() => new Promise<void>((r) => (finish = r)));
    const first = deps({}, slow);
    const second = deps({}, slow);
    const running = runUpdateCheck(first.d, false);
    await new Promise((r) => setTimeout(r, 0));
    expect(await runUpdateCheck(second.d, false)).toBe('busy');
    expect(second.d.ask).not.toHaveBeenCalled();
    expect(second.log).toEqual(['tell:Une mise à jour est déjà en cours : Grillon redémarrera tout seul quand elle sera installée.']);
    // La vérification discrète du démarrage, elle, ne dit rien.
    const quiet = deps({}, slow);
    expect(await runUpdateCheck(quiet.d, true)).toBe('busy');
    expect(quiet.d.tell).not.toHaveBeenCalled();
    finish();
    expect(await running).toBe('installed');
    expect(first.d.ask).toHaveBeenCalledTimes(1);
  });

  it('après une installation réussie, plus aucune proposition (Grillon redémarre)', async () => {
    const { d } = deps({}, update());
    expect(await runUpdateCheck(d, true)).toBe('installed');
    const again = deps({}, update());
    expect(await runUpdateCheck(again.d, true)).toBe('busy');
    expect(again.d.ask).not.toHaveBeenCalled();
  });

  it('refus ou échec : on peut chercher de nouveau', async () => {
    expect(await runUpdateCheck(deps({ ask: async () => false }, update()).d, false)).toBe('declined');
    expect(await runUpdateCheck(deps({}, update(async () => Promise.reject(new Error('réseau')))).d, false)).toBe('error');
    expect(await runUpdateCheck(deps({}, update()).d, false)).toBe('installed');
  });
});

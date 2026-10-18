/**
 * Fichiers récents de l'écran d'accueil. Logique pure : la liste est enregistrée par `storage/recent.ts`.
 */
export interface RecentFile {
  /** Identifiant stable : le chemin (application de bureau) ou un nom unique (navigateur, avec un handle). */
  id: string;
  name: string;
  /** Chemin complet sur le disque (application de bureau). */
  path?: string;
  /** Dernière ouverture ou enregistrement (millisecondes). */
  at: number;
}

export const MAX_RECENT = 12;

/** Place le fichier en tête (sans doublon) et garde les plus récents. */
export function addRecent(list: readonly RecentFile[], file: RecentFile, max = MAX_RECENT): RecentFile[] {
  return [file, ...list.filter((f) => f.id !== file.id)].slice(0, max);
}

export function removeRecent(list: readonly RecentFile[], id: string): RecentFile[] {
  return list.filter((f) => f.id !== id);
}

/** Dossier parent d'un chemin, raccourci (« ~/Documents/Maths ») pour la liste. */
export function folderOf(path: string | undefined, home?: string): string {
  if (!path) return '';
  const parts = path.split('/').filter(Boolean);
  parts.pop();
  let folder = `/${parts.join('/')}`;
  if (home && (folder === home || folder.startsWith(`${home}/`))) folder = `~${folder.slice(home.length)}`;
  // On garde les trois derniers dossiers : le début du chemin n'aide pas à reconnaître un fichier.
  const kept = folder.split('/').filter(Boolean);
  if (kept.length > 3) return `…/${kept.slice(-3).join('/')}`;
  return folder === '~' ? '~' : folder;
}

const MONTHS = ['janv.', 'févr.', 'mars', 'avr.', 'mai', 'juin', 'juil.', 'août', 'sept.', 'oct.', 'nov.', 'déc.'];

/** « Aujourd'hui, 14:05 », « Hier », « Il y a 3 jours », puis la date. */
export function formatWhen(at: number, now = Date.now()): string {
  const day = (ms: number) => {
    const d = new Date(ms);
    return Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
  };
  const days = Math.round((day(now) - day(at)) / 86_400_000);
  const d = new Date(at);
  const hhmm = `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
  if (days <= 0) return `Aujourd’hui, ${hhmm}`;
  if (days === 1) return 'Hier';
  if (days < 7) return `Il y a ${days} jours`;
  return `${d.getDate()} ${MONTHS[d.getMonth()]}${d.getFullYear() === new Date(now).getFullYear() ? '' : ` ${d.getFullYear()}`}`;
}

/** Valide ce qui vient du stockage (un tableau abîmé ne doit pas casser l'écran d'accueil). */
export function parseRecent(json: string | null): RecentFile[] {
  if (!json) return [];
  try {
    const data = JSON.parse(json) as unknown;
    if (!Array.isArray(data)) return [];
    return data
      .filter((f): f is RecentFile => !!f && typeof f.id === 'string' && typeof f.name === 'string' && typeof f.at === 'number')
      .map((f) => ({ id: f.id, name: f.name, at: f.at, ...(typeof f.path === 'string' ? { path: f.path } : {}) }))
      .slice(0, MAX_RECENT);
  } catch {
    return [];
  }
}

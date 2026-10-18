/**
 * Sauvegarde automatique locale dans IndexedDB (aucune donnée ne quitte l'appareil).
 * On garde le dernier état du classeur + le nom du fichier + son handle (Chrome) pour Ctrl+S.
 */
import type { IWorkbookData } from '../model/univer-types';
import { pathOfHandle, tauriHandleFromPath } from '../io/tauri-files';

const DB_NAME = 'grillon';
const STORE = 'autosave';
const KEY = 'current';

/** Les fenêtres supplémentaires (application de bureau) ne touchent pas à la sauvegarde de la fenêtre principale. */
let enabled = true;
export const disableAutosave = (): void => void (enabled = false);

export interface AutosaveEntry {
  data: IWorkbookData;
  fileName: string;
  savedAt: number;
  /** Modifié depuis le dernier enregistrement explicite dans un fichier. */
  dirty: boolean;
  handle?: FileSystemFileHandle;
  /** Application de bureau : chemin du fichier (un handle Tauri ne peut pas être stocké tel quel). */
  path?: string;
}

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function tx<T>(mode: IDBTransactionMode, fn: (store: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  const db = await openDb();
  try {
    return await new Promise<T>((resolve, reject) => {
      const t = db.transaction(STORE, mode);
      const req = fn(t.objectStore(STORE));
      t.oncomplete = () => resolve(req.result);
      t.onerror = () => reject(t.error);
      t.onabort = () => reject(t.error);
    });
  } finally {
    db.close();
  }
}

async function put(key: string, entry: AutosaveEntry): Promise<void> {
  try {
    await tx('readwrite', (s) => s.put(entry, key));
  } catch {
    // Certains navigateurs refusent de cloner un FileSystemFileHandle : on réessaie sans.
    const { handle, ...rest } = entry;
    await tx('readwrite', (s) => s.put({ ...rest, path: pathOfHandle(handle) }, key));
  }
}

async function get(key: string): Promise<AutosaveEntry | null> {
  try {
    const entry = ((await tx('readonly', (s) => s.get(key))) as AutosaveEntry | undefined) ?? null;
    if (entry && !entry.handle && entry.path) entry.handle = tauriHandleFromPath(entry.path);
    return entry;
  } catch {
    return null;
  }
}

export async function saveAutosave(entry: AutosaveEntry): Promise<void> {
  if (!enabled) return;
  await put(KEY, entry);
}

export async function loadAutosave(): Promise<AutosaveEntry | null> {
  if (!enabled) return null;
  return get(KEY);
}

/**
 * Brouillons : les travaux non enregistrés des dernières sessions (cinq au plus, le plus récent d'abord). Au démarrage, Grillon
 * s'ouvre sur l'écran d'accueil et propose de les reprendre (« Reprendre ») au lieu de rouvrir le dernier d'office.
 */
export const MAX_DRAFTS = 5;
const DRAFTS_KEY = 'drafts';
/** Ancien brouillon unique (versions 1.13.0 et 1.13.1) : repris dans la liste. */
const OLD_DRAFT_KEY = 'draft';

export type DraftEntry = AutosaveEntry & { id: string };

async function readDrafts(): Promise<DraftEntry[]> {
  const list = ((await tx('readonly', (s) => s.get(DRAFTS_KEY)).catch(() => undefined)) as DraftEntry[] | undefined) ?? [];
  const old = await get(OLD_DRAFT_KEY);
  const all = old ? [...list, { ...old, id: `old-${old.savedAt}` }] : list;
  if (old) {
    await tx('readwrite', (s) => s.put(all, DRAFTS_KEY)).catch(() => undefined);
    await tx('readwrite', (s) => s.delete(OLD_DRAFT_KEY)).catch(() => undefined);
  }
  return all.map((d) => (!d.handle && d.path ? { ...d, handle: tauriHandleFromPath(d.path) } : d)).sort((a, b) => b.savedAt - a.savedAt);
}

async function writeDrafts(list: DraftEntry[]): Promise<void> {
  // Un handle de navigateur ne se clone pas toujours : pour les brouillons on garde seulement le chemin (application de bureau).
  const plain = list.map(({ handle, ...rest }) => ({ ...rest, ...(pathOfHandle(handle) ? { path: pathOfHandle(handle) } : {}) }));
  try {
    await tx('readwrite', (s) => s.put(list, DRAFTS_KEY));
  } catch {
    await tx('readwrite', (s) => s.put(plain, DRAFTS_KEY));
  }
}

/** Met de côté la sauvegarde automatique si elle contient des modifications non enregistrées. */
export async function stashDraft(): Promise<void> {
  if (!enabled) return;
  const drafts = await readDrafts();
  const entry = await get(KEY);
  if (entry?.dirty && entry.data?.sheetOrder?.length) {
    drafts.unshift({ ...entry, id: `${entry.savedAt}-${Math.random().toString(36).slice(2, 6)}` });
  }
  await writeDrafts(drafts.slice(0, MAX_DRAFTS));
}

export async function loadDrafts(): Promise<DraftEntry[]> {
  if (!enabled) return [];
  return readDrafts();
}

export async function clearDraft(id: string): Promise<void> {
  if (!enabled) return;
  await writeDrafts((await readDrafts()).filter((d) => d.id !== id));
}

export async function clearAutosave(): Promise<void> {
  if (!enabled) return;
  try {
    await tx('readwrite', (s) => s.delete(KEY));
  } catch {
    /* rien à effacer */
  }
}

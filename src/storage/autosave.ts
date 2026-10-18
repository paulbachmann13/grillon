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
 * Brouillon : le travail non enregistré de la dernière session. Au démarrage, Grillon s'ouvre sur l'écran d'accueil
 * et propose de le reprendre (« Reprendre ») au lieu de le rouvrir d'office. Un seul brouillon est gardé.
 */
const DRAFT_KEY = 'draft';

/** Met de côté la sauvegarde automatique si elle contient des modifications non enregistrées. */
export async function stashDraft(): Promise<AutosaveEntry | null> {
  if (!enabled) return null;
  const entry = await get(KEY);
  if (!entry?.dirty || !entry.data?.sheetOrder?.length) return loadDraft();
  await put(DRAFT_KEY, entry);
  return entry;
}

export async function loadDraft(): Promise<AutosaveEntry | null> {
  if (!enabled) return null;
  return get(DRAFT_KEY);
}

export async function clearDraft(): Promise<void> {
  if (!enabled) return;
  try {
    await tx('readwrite', (s) => s.delete(DRAFT_KEY));
  } catch {
    /* rien à effacer */
  }
}

export async function clearAutosave(): Promise<void> {
  if (!enabled) return;
  try {
    await tx('readwrite', (s) => s.delete(KEY));
  } catch {
    /* rien à effacer */
  }
}

/**
 * Sauvegarde automatique locale dans IndexedDB (aucune donnée ne quitte l'appareil).
 * On garde le dernier état du classeur + le nom du fichier + son handle (Chrome) pour Ctrl+S.
 */
import type { IWorkbookData } from '../model/univer-types';

const DB_NAME = 'grillon';
const STORE = 'autosave';
const KEY = 'current';

export interface AutosaveEntry {
  data: IWorkbookData;
  fileName: string;
  savedAt: number;
  /** Modifié depuis le dernier enregistrement explicite dans un fichier. */
  dirty: boolean;
  handle?: FileSystemFileHandle;
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

export async function saveAutosave(entry: AutosaveEntry): Promise<void> {
  try {
    await tx('readwrite', (s) => s.put(entry, KEY));
  } catch {
    // Certains navigateurs refusent de cloner un FileSystemFileHandle : on réessaie sans.
    const { handle: _ignored, ...rest } = entry;
    await tx('readwrite', (s) => s.put(rest, KEY));
  }
}

export async function loadAutosave(): Promise<AutosaveEntry | null> {
  try {
    return ((await tx('readonly', (s) => s.get(KEY))) as AutosaveEntry | undefined) ?? null;
  } catch {
    return null;
  }
}

export async function clearAutosave(): Promise<void> {
  try {
    await tx('readwrite', (s) => s.delete(KEY));
  } catch {
    /* rien à effacer */
  }
}

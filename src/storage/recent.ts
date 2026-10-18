/**
 * Fichiers récents : la liste (nom, chemin, date) est dans localStorage ; dans le navigateur, le « handle »
 * du fichier (File System Access) est gardé dans IndexedDB pour pouvoir le rouvrir. Application de bureau :
 * le chemin suffit.
 */
import { addRecent, parseRecent, removeRecent, type RecentFile } from '../data/recent';
import { pathOfHandle, tauriHandleFromPath } from '../io/tauri-files';

const KEY = 'grillon-recent';
const DB = 'grillon-recent';
const STORE = 'handles';

export function loadRecent(): RecentFile[] {
  try {
    return parseRecent(localStorage.getItem(KEY));
  } catch {
    return [];
  }
}

function saveList(list: RecentFile[]): void {
  try {
    localStorage.setItem(KEY, JSON.stringify(list));
  } catch {
    /* stockage indisponible : la liste vaut pour la session */
  }
}

function handles(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function handleStore<T>(mode: IDBTransactionMode, fn: (s: IDBObjectStore) => IDBRequest<T>): Promise<T | undefined> {
  try {
    const db = await handles();
    try {
      return await new Promise<T>((resolve, reject) => {
        const t = db.transaction(STORE, mode);
        const req = fn(t.objectStore(STORE));
        t.oncomplete = () => resolve(req.result);
        t.onerror = t.onabort = () => reject(t.error);
      });
    } finally {
      db.close();
    }
  } catch {
    return undefined;
  }
}

/** Note un fichier ouvert ou enregistré. Sans chemin ni handle (Safari, glisser-déposer), il n'y a rien à rouvrir : on n'enregistre rien. */
export async function rememberFile(handle: FileSystemFileHandle | undefined, name: string): Promise<void> {
  const path = pathOfHandle(handle);
  if (!path && !handle) return;
  const id = path ?? `h:${name}`;
  if (!path && handle) await handleStore('readwrite', (s) => s.put(handle, id));
  saveList(addRecent(loadRecent(), { id, name, at: Date.now(), ...(path ? { path } : {}) }));
}

export async function forgetRecent(id: string): Promise<void> {
  saveList(removeRecent(loadRecent(), id));
  await handleStore('readwrite', (s) => s.delete(id));
}

/** Le handle d'un fichier récent, prêt à être lu (le navigateur redemande l'autorisation), ou null s'il n'est plus accessible. */
export async function recentHandle(file: RecentFile): Promise<FileSystemFileHandle | null> {
  if (file.path) return tauriHandleFromPath(file.path);
  const handle = (await handleStore('readonly', (s) => s.get(file.id))) as (FileSystemFileHandle & { requestPermission?(o: { mode: string }): Promise<string> }) | undefined;
  if (!handle) return null;
  try {
    if (handle.requestPermission && (await handle.requestPermission({ mode: 'readwrite' })) !== 'granted') return null;
    return handle;
  } catch {
    return null;
  }
}

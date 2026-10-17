/**
 * Fichiers dans l'application de bureau (Tauri) : boîtes de dialogue natives et réécriture du
 * fichier d'origine. Les modules @tauri-apps ne sont chargés que dans l'application de bureau.
 *
 * Le reste de l'appli manipule un « handle » à la manière de l'API File System Access
 * (getFile / createWritable) : on en fabrique un qui lit et écrit un chemin du disque.
 */

export const isTauri = (): boolean => typeof window !== 'undefined' && '__TAURI_INTERNALS__' in window;

export interface TauriFileHandle {
  kind: 'file';
  name: string;
  /** Chemin sur le disque (sert aussi à retrouver le fichier après un redémarrage). */
  path: string;
  getFile(): Promise<File>;
  createWritable(): Promise<{ write(data: Uint8Array | string): Promise<void>; close(): Promise<void> }>;
}

export const baseName = (path: string): string => path.split(/[\/]/).pop() || path;

export function tauriHandleFromPath(path: string): FileSystemFileHandle {
  const name = baseName(path);
  const handle: TauriFileHandle = {
    kind: 'file',
    name,
    path,
    async getFile() {
      const { readFile } = await import('@tauri-apps/plugin-fs');
      return new File([(await readFile(path)) as BlobPart], name);
    },
    async createWritable() {
      const chunks: Uint8Array[] = [];
      return {
        async write(data) {
          chunks.push(typeof data === 'string' ? new TextEncoder().encode(data) : data);
        },
        async close() {
          const { writeFile } = await import('@tauri-apps/plugin-fs');
          const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
          let offset = 0;
          for (const c of chunks) {
            out.set(c, offset);
            offset += c.length;
          }
          await writeFile(path, out);
        },
      };
    },
  };
  return handle as unknown as FileSystemFileHandle;
}

export const pathOfHandle = (handle: FileSystemFileHandle | undefined): string | undefined => (handle as unknown as Partial<TauriFileHandle> | undefined)?.path;

export interface DialogFilter {
  name: string;
  extensions: string[];
}

/** Boîte « Ouvrir » native. Renvoie null si l'utilisateur annule. */
export async function tauriPickOpen(filters: DialogFilter[]): Promise<FileSystemFileHandle | null> {
  const { open } = await import('@tauri-apps/plugin-dialog');
  const path = await open({ multiple: false, directory: false, filters });
  return typeof path === 'string' ? tauriHandleFromPath(path) : null;
}

/** Boîte « Enregistrer sous » native. Renvoie null si l'utilisateur annule. */
export async function tauriPickSave(defaultPath: string, filters: DialogFilter[]): Promise<FileSystemFileHandle | null> {
  const { save } = await import('@tauri-apps/plugin-dialog');
  const path = await save({ defaultPath, filters });
  return path ? tauriHandleFromPath(path) : null;
}

/** Étiquette de la fenêtre créée par « Nouveau » / « Ouvrir » (absente pour la fenêtre principale). */
export const secondaryWindowLabel = (): string | null => (typeof location === 'undefined' ? null : new URLSearchParams(location.search).get('w'));

/** Fichiers demandés par le Finder avant le démarrage (fenêtre principale seulement). */
export async function takeStartupOpens(): Promise<string[]> {
  const { invoke } = await import('@tauri-apps/api/core');
  return invoke<string[]>('take_pending_opens');
}

/** Fichier que doit ouvrir une fenêtre créée par openInNewWindow (null : classeur vide). */
export async function takeWindowOpen(label: string): Promise<string | null> {
  const { invoke } = await import('@tauri-apps/api/core');
  return (await invoke<string | null>('take_window_open', { label })) ?? null;
}

/** Ouvre une nouvelle fenêtre Grillon (comme Excel : un classeur par fenêtre), avec un fichier éventuel. */
export async function openInNewWindow(path?: string): Promise<void> {
  const { invoke } = await import('@tauri-apps/api/core');
  await invoke('new_window', { path: path ?? null });
}

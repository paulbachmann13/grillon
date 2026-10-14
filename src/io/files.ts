/**
 * Ouvrir / enregistrer des fichiers.
 *
 * - Chrome/Edge (bureau) : File System Access API → on peut réécrire le même fichier (Ctrl+S).
 * - Safari, iOS, Firefox : sélecteur de fichier classique + téléchargement.
 */

export interface OpenedFile {
  name: string;
  bytes: Uint8Array;
  handle?: FileSystemFileHandle;
}

type PickerType = { description: string; accept: Record<string, string[]> };

const XLSX_TYPE: PickerType = {
  description: 'Classeur Excel',
  accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] },
};
const CSV_TYPE: PickerType = { description: 'Fichier CSV', accept: { 'text/csv': ['.csv'] } };

interface FsWindow {
  showOpenFilePicker?: (opts: { types?: PickerType[]; multiple?: boolean; excludeAcceptAllOption?: boolean }) => Promise<FileSystemFileHandle[]>;
  showSaveFilePicker?: (opts: { suggestedName?: string; types?: PickerType[] }) => Promise<FileSystemFileHandle>;
}

const fsWindow = (): FsWindow => window as unknown as FsWindow;

export function hasFileSystemAccess(): boolean {
  return typeof fsWindow().showOpenFilePicker === 'function' && typeof fsWindow().showSaveFilePicker === 'function';
}

function isAbort(e: unknown): boolean {
  return e instanceof DOMException && e.name === 'AbortError';
}

/** Sélecteur via <input type=file> (Safari/iOS). Résout null si l'utilisateur annule. */
function pickWithInput(accept: string): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = accept;
    input.style.display = 'none';
    input.dataset.testid = 'file-input';
    let done = false;
    const finish = (f: File | null) => {
      if (done) return;
      done = true;
      input.remove();
      resolve(f);
    };
    input.addEventListener('change', () => finish(input.files?.[0] ?? null));
    input.addEventListener('cancel', () => finish(null));
    document.body.appendChild(input);
    input.click();
  });
}

export async function openFile(kind: 'xlsx' | 'csv' | 'any' = 'any'): Promise<OpenedFile | null> {
  const types = kind === 'xlsx' ? [XLSX_TYPE] : kind === 'csv' ? [CSV_TYPE] : [XLSX_TYPE, CSV_TYPE];
  if (hasFileSystemAccess()) {
    try {
      const [handle] = await fsWindow().showOpenFilePicker!({ types, multiple: false });
      const file = await handle.getFile();
      return { name: file.name, bytes: new Uint8Array(await file.arrayBuffer()), handle };
    } catch (e) {
      if (isAbort(e)) return null;
      throw e;
    }
  }
  const accept = types.flatMap((t) => [...Object.keys(t.accept), ...Object.values(t.accept).flat()]).join(',');
  const file = await pickWithInput(accept);
  if (!file) return null;
  return { name: file.name, bytes: new Uint8Array(await file.arrayBuffer()) };
}

/** Écrit dans un fichier déjà ouvert (sans dialogue). Renvoie false si la permission est refusée. */
export async function writeToHandle(handle: FileSystemFileHandle, bytes: Uint8Array | string): Promise<boolean> {
  const h = handle as FileSystemFileHandle & {
    queryPermission?: (d: { mode: 'readwrite' }) => Promise<PermissionState>;
    requestPermission?: (d: { mode: 'readwrite' }) => Promise<PermissionState>;
  };
  if (h.queryPermission && (await h.queryPermission({ mode: 'readwrite' })) !== 'granted') {
    if (!h.requestPermission || (await h.requestPermission({ mode: 'readwrite' })) !== 'granted') return false;
  }
  const writable = await handle.createWritable();
  await writable.write(bytes as FileSystemWriteChunkType);
  await writable.close();
  return true;
}

export function download(bytes: Uint8Array | string, name: string, mime: string): void {
  const blob = new Blob([bytes as BlobPart], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

export const MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
export const MIME_CSV = 'text/csv;charset=utf-8';

/**
 * « Enregistrer sous » : dialogue natif si disponible, sinon téléchargement.
 * Renvoie le handle choisi (pour les Ctrl+S suivants), null si annulé, undefined si téléchargé.
 */
export async function saveAs(
  bytes: Uint8Array | string,
  suggestedName: string,
  kind: 'xlsx' | 'csv',
): Promise<FileSystemFileHandle | null | undefined> {
  if (hasFileSystemAccess()) {
    try {
      const handle = await fsWindow().showSaveFilePicker!({ suggestedName, types: [kind === 'xlsx' ? XLSX_TYPE : CSV_TYPE] });
      await writeToHandle(handle, bytes);
      return handle;
    } catch (e) {
      if (isAbort(e)) return null;
      throw e;
    }
  }
  download(bytes, suggestedName, kind === 'xlsx' ? MIME_XLSX : MIME_CSV);
  return undefined;
}

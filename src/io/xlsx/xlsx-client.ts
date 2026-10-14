/**
 * Accès au convertisseur .xlsx : dans un Web Worker si possible, sinon (worker
 * indisponible ou en échec) directement dans la page.
 */
import type { IWorkbookData } from '../../model/univer-types';
import type { ImportOptions } from './import';

let worker: Worker | null = null;
let failed = false;
let nextId = 0;
const pending = new Map<number, { resolve: (v: unknown) => void; reject: (e: Error) => void }>();

function getWorker(): Worker | null {
  if (failed || typeof Worker === 'undefined') return null;
  if (!worker) {
    try {
      worker = new Worker(new URL('./xlsx-worker.ts', import.meta.url), { type: 'module' });
      worker.onmessage = (e: MessageEvent<{ id: number; ok: boolean; result?: unknown; error?: string }>) => {
        const p = pending.get(e.data.id);
        if (!p) return;
        pending.delete(e.data.id);
        if (e.data.ok) p.resolve(e.data.result);
        else p.reject(new Error(e.data.error));
      };
      worker.onerror = () => {
        failed = true;
        for (const p of pending.values()) p.reject(new Error('worker'));
        pending.clear();
      };
    } catch {
      failed = true;
      return null;
    }
  }
  return worker;
}

function call<T>(message: Record<string, unknown>, transfer: Transferable[] = []): Promise<T> {
  const w = getWorker();
  if (!w) return Promise.reject(new Error('worker'));
  const id = ++nextId;
  return new Promise<T>((resolve, reject) => {
    pending.set(id, { resolve: resolve as (v: unknown) => void, reject });
    w.postMessage({ ...message, id }, transfer);
  });
}

export async function importXlsxAsync(bytes: Uint8Array, options: ImportOptions = {}): Promise<IWorkbookData> {
  try {
    // Copie : le tampon est transféré au worker.
    const copy = bytes.slice();
    return await call<IWorkbookData>({ kind: 'import', bytes: copy, options }, [copy.buffer]);
  } catch (e) {
    if ((e as Error).message !== 'worker') throw e;
    const { importXlsx } = await import('./import');
    return importXlsx(bytes, options);
  }
}

export async function exportXlsxAsync(data: IWorkbookData): Promise<Uint8Array> {
  try {
    return await call<Uint8Array>({ kind: 'export', data });
  } catch (e) {
    if ((e as Error).message !== 'worker') throw e;
    const { exportXlsx } = await import('./export');
    return exportXlsx(data);
  }
}

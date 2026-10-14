/**
 * Web Worker de conversion .xlsx : l'ouverture et l'enregistrement de gros classeurs
 * (100 000 lignes) ne figent pas l'interface.
 */
import type { IWorkbookData } from '../../model/univer-types';
import { exportXlsx } from './export';
import { importXlsx, type ImportOptions } from './import';

type Request =
  | { id: number; kind: 'import'; bytes: Uint8Array; options: ImportOptions }
  | { id: number; kind: 'export'; data: IWorkbookData };

self.onmessage = async (e: MessageEvent<Request>) => {
  const req = e.data;
  try {
    if (req.kind === 'import') {
      const data = await importXlsx(req.bytes, req.options);
      (self as unknown as Worker).postMessage({ id: req.id, ok: true, result: data });
    } else {
      const bytes = await exportXlsx(req.data);
      (self as unknown as Worker).postMessage({ id: req.id, ok: true, result: bytes }, [bytes.buffer]);
    }
  } catch (err) {
    (self as unknown as Worker).postMessage({ id: req.id, ok: false, error: (err as Error).message ?? String(err) });
  }
};

/**
 * Grille qui s'agrandit toute seule, jusqu'aux limites d'Excel (1 048 576 lignes, 16 384 colonnes).
 *
 * Univer dessine une grille de taille fixe et certaines opérations coûtent proportionnellement
 * au nombre de lignes (changer une hauteur prend ~1,5 s sur 1 048 576 lignes). On garde donc
 * une grille raisonnable et on ajoute des lignes / colonnes quand l'utilisateur approche du bord,
 * en faisant défiler ou en déplaçant la sélection. L'ajout passe par une mutation : il ne
 * marque pas le classeur comme modifié et ne remplit pas l'historique d'annulation.
 */
import { ICommandService, type FUniver } from '@univerjs/presets';

export const EXCEL_MAX_ROWS = 1_048_576;
export const EXCEL_MAX_COLUMNS = 16_384;
const ROW_STEP = 1000;
const COLUMN_STEP = 26;
/** Distance au bord (en lignes / colonnes) à partir de laquelle on agrandit. */
const ROW_MARGIN = 100;
const COLUMN_MARGIN = 8;

/** Nouvelle taille quand la dernière ligne (ou colonne) utile `last` approche de `count`. */
export function grownSize(count: number, last: number, margin: number, step: number, max: number): number {
  if (count >= max || last < count - margin) return count;
  // Assez de place pour `last` + la marge, arrondi au pas suivant.
  const wanted = Math.ceil((last + margin + 1) / step) * step;
  return Math.min(max, Math.max(count + step, wanted));
}

export function installGridGrowth(api: FUniver): void {
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const commands = injector.get<ICommandService>(ICommandService);

  const grow = (lastRow: number, lastColumn: number) => {
    const wb = api.getActiveWorkbook();
    const ws = wb?.getActiveSheet();
    if (!wb || !ws) return;
    const rows = ws.getMaxRows();
    const columns = ws.getMaxColumns();
    const newRows = grownSize(rows, lastRow, ROW_MARGIN, ROW_STEP, EXCEL_MAX_ROWS);
    const newColumns = grownSize(columns, lastColumn, COLUMN_MARGIN, COLUMN_STEP, EXCEL_MAX_COLUMNS);
    const target = { unitId: wb.getId(), subUnitId: ws.getSheetId() };
    if (newRows !== rows) commands.syncExecuteCommand('sheet.mutation.set-worksheet-row-count', { ...target, rowCount: newRows });
    if (newColumns !== columns) commands.syncExecuteCommand('sheet.mutation.set-worksheet-column-count', { ...target, columnCount: newColumns });
  };

  const fromViewport = () => {
    const ws = api.getActiveWorkbook()?.getActiveSheet() as unknown as { getVisibleRange?(): { endRow: number; endColumn: number } | null } | undefined;
    const range = ws?.getVisibleRange?.();
    if (range) grow(range.endRow, range.endColumn);
  };

  api.addEvent(api.Event.Scroll, () => setTimeout(fromViewport, 0));
  // Toute sélection (souris, clavier, Ctrl+flèche, zone Nom, programme) passe par cette opération.
  commands.onCommandExecuted((info) => {
    if (info.id !== 'sheet.operation.set-selections') return;
    const selections = (info.params as { selections?: { range: { endRow: number; endColumn: number } }[] } | undefined)?.selections ?? [];
    let row = 0;
    let column = 0;
    for (const { range } of selections) {
      row = Math.max(row, range.endRow);
      column = Math.max(column, range.endColumn);
    }
    // Après la fin de l'opération en cours (on ne lance pas une mutation au milieu d'une autre).
    setTimeout(() => grow(row, column), 0);
  });
  // Fenêtre plus grande que la grille (grand écran, zoom arrière) : on remplit tout de suite.
  api.addEvent(api.Event.SheetZoomChanged, fromViewport);
  api.addEvent(api.Event.ActiveSheetChanged, () => setTimeout(fromViewport, 0));
  setTimeout(fromViewport, 0);
}

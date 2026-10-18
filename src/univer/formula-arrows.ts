/**
 * Barre de formule : les flèches gauche / droite déplacent le curseur dans le texte (comme dans Excel
 * en mode « Modifier »), au lieu de déplacer la plage de cellules de la formule.
 * Univer traite sinon la barre de formule comme une saisie de références : après un clic dans la barre,
 * une flèche change `A1:A3` en `Y1000`.
 */
import { FOCUSING_FX_BAR_EDITOR, IContextService, Direction } from '@univerjs/core';
import type { FUniver } from '@univerjs/presets';
import { MoveCursorOperation, MoveSelectionOperation } from '@univerjs/docs-ui';
import { IShortcutService, KeyCode, MetaKeys } from '@univerjs/ui';

const installed = new WeakSet<object>();

export function installFormulaBarArrows(api: FUniver): void {
  if (installed.has(api)) return;
  installed.add(api);
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const shortcuts = injector.get<IShortcutService>(IShortcutService);
  const context = injector.get<IContextService>(IContextService);
  const inFormulaBar = () => !!context.getContextValue(FOCUSING_FX_BAR_EDITOR);
  const keys: [number, Direction][] = [
    [KeyCode.ARROW_LEFT, Direction.LEFT],
    [KeyCode.ARROW_RIGHT, Direction.RIGHT],
  ];
  for (const [keyCode, direction] of keys) {
    // Flèche seule : le curseur ; Maj + flèche : étend la sélection de texte.
    shortcuts.registerShortcut({ id: MoveCursorOperation.id, binding: keyCode, priority: 1000, preconditions: inFormulaBar, staticParameters: { direction } });
    shortcuts.registerShortcut({ id: MoveSelectionOperation.id, binding: keyCode | MetaKeys.SHIFT, priority: 1000, preconditions: inFormulaBar, staticParameters: { direction } });
  }
}

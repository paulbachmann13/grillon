/**
 * Formules → Options de calcul (Automatique / Manuel) et « Calculer maintenant » (F9), comme Excel.
 * En mode manuel, le service de déclenchement d'Univer est mis en pause : les modifications
 * s'accumulent et sont calculées d'un coup avec F9.
 */
import { FormulaCalculationTriggerService } from '@univerjs/preset-sheets-core';
import type { FUniver } from '@univerjs/presets';
import { t } from '../i18n';

interface Trigger {
  _started: boolean;
  _waitingCommandQueue: unknown[];
  _scheduleFlush(): void;
}

let manual = false;
let badge: HTMLElement | null = null;

function trigger(api: FUniver): Trigger {
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  return injector.get<Trigger>(FormulaCalculationTriggerService);
}

export const isManualCalculation = () => manual;

function showBadge(): void {
  if (!manual) {
    badge?.remove();
    return;
  }
  if (!badge) {
    badge = document.createElement('div');
    badge.className = 'g-calc-badge';
    badge.dataset.testid = 'calc-manual';
    badge.setAttribute('role', 'status');
  }
  badge.textContent = t('calc.manualBadge');
  if (!badge.isConnected) document.body.append(badge);
}

export function setCalculationMode(api: FUniver, mode: 'auto' | 'manual'): void {
  manual = mode === 'manual';
  const s = trigger(api);
  s._started = !manual;
  // Retour en automatique : Excel recalcule tout de suite ce qui attendait.
  if (!manual) s._scheduleFlush();
  showBadge();
}

/** F9 : calcule les formules en attente (ou tout le classeur s'il n'y a rien en attente). */
export function calculateNow(api: FUniver): void {
  const s = trigger(api);
  s._started = true;
  if (s._waitingCommandQueue.length) s._scheduleFlush();
  else api.getFormula().executeCalculation();
  // En mode manuel, on se remet en pause une fois le calcul lancé (le minuteur part de _scheduleFlush).
  if (manual) setTimeout(() => (s._started = false), 50);
}

export function installCalculationShortcut(api: FUniver): void {
  window.addEventListener(
    'keydown',
    (e) => {
      if (e.key !== 'F9' || e.ctrlKey || e.metaKey || e.altKey) return;
      const active = document.activeElement as HTMLElement | null;
      if (active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA' || active.tagName === 'SELECT')) return;
      e.preventDefault();
      calculateNow(api);
    },
    { capture: true },
  );
}

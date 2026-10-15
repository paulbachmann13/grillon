/**
 * Boîte de dialogue modale simple et accessible (Échap ferme, Entrée valide).
 * Les outils Grillon (Supprimer les doublons, Sous-total…) l'utilisent pour leurs options.
 */
import { t } from '../i18n';

export interface DialogOptions {
  title: string;
  body: HTMLElement;
  okLabel?: string;
  cancelLabel?: string;
  /** Renvoie false (ou un message d'erreur) pour garder la boîte ouverte. */
  onOk: () => boolean | string | void | Promise<boolean | string | void>;
  testId?: string;
  /** Boîte plus large (aperçus, tableaux). */
  wide?: boolean;
}

export function openDialog(opts: DialogOptions): Promise<boolean> {
  return new Promise((resolve) => {
    const backdrop = document.createElement('div');
    backdrop.className = 'g-dialog-backdrop';
    const dialog = document.createElement('div');
    dialog.className = opts.wide ? 'g-dialog g-dialog-wide' : 'g-dialog';
    dialog.setAttribute('role', 'dialog');
    dialog.setAttribute('aria-modal', 'true');
    if (opts.testId) dialog.dataset.testid = opts.testId;
    const titleId = `g-dialog-title-${Math.random().toString(36).slice(2, 8)}`;
    dialog.setAttribute('aria-labelledby', titleId);

    const title = document.createElement('h2');
    title.id = titleId;
    title.textContent = opts.title;
    const error = document.createElement('p');
    error.className = 'g-dialog-error';
    error.setAttribute('role', 'alert');
    const actions = document.createElement('div');
    actions.className = 'g-dialog-actions';
    const ok = button(opts.okLabel ?? t('dialog.ok'), 'primary');
    const cancel = button(opts.cancelLabel ?? t('dialog.cancel'));
    actions.append(cancel, ok);
    dialog.append(title, opts.body, error, actions);
    backdrop.append(dialog);
    document.body.append(backdrop);

    const previous = document.activeElement as HTMLElement | null;
    const close = (result: boolean) => {
      backdrop.remove();
      document.removeEventListener('keydown', onKey, true);
      previous?.focus?.();
      resolve(result);
    };
    const submit = async () => {
      ok.disabled = true;
      try {
        const r = await opts.onOk();
        if (r === false || typeof r === 'string') {
          error.textContent = typeof r === 'string' ? r : '';
          return;
        }
        close(true);
      } catch (e) {
        error.textContent = (e as Error).message;
      } finally {
        ok.disabled = false;
      }
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        close(false);
      } else if (e.key === 'Enter' && !(e.target instanceof HTMLTextAreaElement) && !(e.target instanceof HTMLButtonElement)) {
        e.preventDefault();
        e.stopPropagation();
        void submit();
      }
      // Les touches tapées dans la boîte ne doivent pas atteindre la grille.
      if (dialog.contains(e.target as Node)) e.stopPropagation();
    };
    document.addEventListener('keydown', onKey, true);
    ok.addEventListener('click', () => void submit());
    cancel.addEventListener('click', () => close(false));
    backdrop.addEventListener('pointerdown', (e) => {
      if (e.target === backdrop) close(false);
    });
    (dialog.querySelector('input, select, textarea, button') as HTMLElement | null)?.focus();
  });
}

function button(label: string, kind?: 'primary'): HTMLButtonElement {
  const b = document.createElement('button');
  b.type = 'button';
  b.textContent = label;
  if (kind) b.className = kind;
  return b;
}

// --- Petits constructeurs de formulaire ---

export function field(label: string, control: HTMLElement): HTMLLabelElement {
  const l = document.createElement('label');
  l.className = 'g-field';
  const span = document.createElement('span');
  span.textContent = label;
  l.append(span, control);
  return l;
}

export function checkbox(label: string, checked = false): { el: HTMLLabelElement; input: HTMLInputElement } {
  const el = document.createElement('label');
  el.className = 'g-check';
  const input = document.createElement('input');
  input.type = 'checkbox';
  input.checked = checked;
  el.append(input, document.createTextNode(label));
  return { el, input };
}

export function radio(name: string, label: string, value: string, checked = false): { el: HTMLLabelElement; input: HTMLInputElement } {
  const el = document.createElement('label');
  el.className = 'g-check';
  const input = document.createElement('input');
  input.type = 'radio';
  input.name = name;
  input.value = value;
  input.checked = checked;
  el.append(input, document.createTextNode(label));
  return { el, input };
}

export function select(options: [value: string, label: string][], value?: string): HTMLSelectElement {
  const s = document.createElement('select');
  for (const [v, l] of options) {
    const o = document.createElement('option');
    o.value = v;
    o.textContent = l;
    s.append(o);
  }
  if (value !== undefined) s.value = value;
  return s;
}

export function textInput(value = '', placeholder = ''): HTMLInputElement {
  const i = document.createElement('input');
  i.type = 'text';
  i.value = value;
  i.placeholder = placeholder;
  return i;
}

export function group(title: string, ...children: HTMLElement[]): HTMLFieldSetElement {
  const fs = document.createElement('fieldset');
  const legend = document.createElement('legend');
  legend.textContent = title;
  fs.append(legend, ...children);
  return fs;
}

export function box(...children: (HTMLElement | string)[]): HTMLDivElement {
  const d = document.createElement('div');
  d.className = 'g-dialog-body';
  for (const c of children) d.append(typeof c === 'string' ? Object.assign(document.createElement('p'), { textContent: c }) : c);
  return d;
}

/** Message d'information (une seule action « OK »). */
export async function alertDialog(title: string, message: string): Promise<void> {
  await openDialog({ title, body: box(message), onOk: () => true, cancelLabel: t('dialog.close') });
}

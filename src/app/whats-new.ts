/**
 * Nouveautés et historique des versions : le CHANGELOG, affiché dans une fenêtre. Montré une fois après chaque mise à jour,
 * et à la demande (Fichier → Nouveautés, ou le numéro de version de l'accueil).
 */
import changelogText from '../../CHANGELOG.md?raw';
import { compareVersions, notesSince, parseChangelog, parseInline, parseNotes, type NoteBlock, type ReleaseNotes } from '../data/changelog';
import { t } from '../i18n';
import { loadDrafts } from '../storage/autosave';
import { loadRecent } from '../storage/recent';
import { openDialog } from '../ui/dialog';
import { appVersion } from './version';

const STORAGE_KEY = 'grillon-last-version';
const all = (): ReleaseNotes[] => parseChangelog(changelogText);

function inline(parent: HTMLElement, text: string): void {
  for (const part of parseInline(text)) {
    if (part.bold || part.code) {
      const el = document.createElement(part.bold ? 'strong' : 'code');
      el.textContent = part.text;
      parent.append(el);
    } else parent.append(document.createTextNode(part.text));
  }
}

/** Blocs de notes → éléments (jamais de HTML brut : tout passe par textContent). */
export function renderBlocks(blocks: NoteBlock[]): HTMLElement {
  const root = document.createElement('div');
  root.className = 'g-notes';
  let list: HTMLUListElement | null = null;
  const lists: HTMLUListElement[] = [];
  for (const block of blocks) {
    if (block.kind !== 'item') list = null;
    if (block.kind === 'heading') {
      const h = document.createElement('h4');
      h.textContent = block.text;
      root.append(h);
    } else if (block.kind === 'paragraph') {
      const p = document.createElement('p');
      inline(p, block.text);
      root.append(p);
    } else {
      if (!list) {
        list = document.createElement('ul');
        root.append(list);
        lists.length = 0;
        lists.push(list);
      }
      // Niveaux : une sous-puce va dans la liste du niveau au-dessus.
      while (lists.length > block.depth + 1) lists.pop();
      while (lists.length < block.depth + 1) {
        const parentList = lists[lists.length - 1]!;
        const host = (parentList.lastElementChild as HTMLElement | null) ?? parentList.appendChild(document.createElement('li'));
        const sub = document.createElement('ul');
        host.append(sub);
        lists.push(sub);
      }
      const li = document.createElement('li');
      inline(li, block.text);
      lists[lists.length - 1]!.append(li);
    }
  }
  return root;
}

function releaseSection(r: ReleaseNotes): HTMLElement {
  const section = document.createElement('section');
  section.className = 'g-release';
  section.dataset.version = r.version;
  const h = document.createElement('h3');
  h.textContent = `Grillon ${r.version}`;
  const meta = [r.title, r.date ? t('whatsNew.date', { date: new Date(`${r.date}T12:00:00`).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' }) }) : undefined].filter(Boolean).join(' · ');
  if (meta) {
    const small = document.createElement('span');
    small.className = 'g-release-date';
    small.textContent = meta;
    h.append(small);
  }
  section.append(h, r.blocks.length ? renderBlocks(r.blocks) : Object.assign(document.createElement('p'), { textContent: t('whatsNew.empty') }));
  return section;
}

/** Notes d'une version à venir (celles de latest.json) : sous la question « Installer ? » de la mise à jour. */
export function notesElement(markdown: string): HTMLElement {
  const wrap = document.createElement('div');
  wrap.className = 'g-update-notes';
  const label = document.createElement('p');
  label.className = 'g-update-notes-label';
  label.textContent = t('whatsNew.updateNotes');
  wrap.append(label, renderBlocks(parseNotes(markdown)));
  return wrap;
}

function listElement(releases: ReleaseNotes[], hint?: string): HTMLElement {
  const body = document.createElement('div');
  body.className = 'g-dialog-body g-notes-body';
  body.dataset.testid = 'whats-new';
  if (hint) {
    const p = document.createElement('p');
    p.className = 'g-start-sub';
    p.textContent = hint;
    body.append(p);
  }
  for (const r of releases) body.append(releaseSection(r));
  return body;
}

/** Historique complet des versions. */
export async function showHistory(): Promise<void> {
  await openDialog({
    title: t('whatsNew.historyTitle'),
    body: listElement(all(), t('whatsNew.historyHint')),
    okLabel: t('whatsNew.close'),
    cancelLabel: t('whatsNew.close'),
    onOk: () => undefined,
    wide: true,
    testId: 'dialog-history',
  });
}

function lastSeen(): string | null {
  try {
    return localStorage.getItem(STORAGE_KEY);
  } catch {
    return null;
  }
}

function remember(version: string): void {
  try {
    localStorage.setItem(STORAGE_KEY, version);
  } catch {
    /* stockage indisponible : on réessaiera au prochain démarrage */
  }
}

/** Grillon a-t-elle déjà servi sur cet ordinateur ? (fichiers récents ou travaux gardés) */
async function alreadyUsed(): Promise<boolean> {
  try {
    return loadRecent().length > 0 || (await loadDrafts()).length > 0;
  } catch {
    return false;
  }
}

/**
 * Au démarrage : si Grillon vient d'être mise à jour, raconte ce qui a changé (une seule fois par version).
 * Rien de mémorisé : une installation neuve ne raconte rien (pas d'historique à un nouvel utilisateur) ; une copie qui a déjà
 * servi, avant que Grillon ne garde sa version, raconte la version en cours.
 */
export async function showWhatsNewIfUpdated(): Promise<void> {
  const current = appVersion();
  if (!current) return;
  const from = lastSeen();
  remember(current);
  if (from && compareVersions(from, current) >= 0) return;
  if (!from && !(await alreadyUsed())) return;
  const releases = notesSince(all(), from, current);
  if (!releases.length) return;
  const wantsAll = !(await openDialog({
    title: releases.length > 1 && from ? t('whatsNew.titleMany', { from }) : t('whatsNew.title', { version: current }),
    body: listElement(releases),
    okLabel: t('whatsNew.ok'),
    cancelLabel: t('whatsNew.all'),
    onOk: () => undefined,
    wide: true,
    testId: 'dialog-whats-new',
  }));
  if (wantsAll) await showHistory();
}

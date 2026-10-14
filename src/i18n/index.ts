import { fr, type Messages } from './fr';

const catalogs: Record<string, Messages> = { fr };
let current: Messages = fr;

export function setLanguage(lang: string): void {
  current = catalogs[lang] ?? fr;
}

/**
 * t('file.open') → « Ouvrir… » ; t('status.loading', { name }) remplace {name}.
 */
export function t(key: string, params: Record<string, string | number> = {}): string {
  const value = key.split('.').reduce<unknown>((o, k) => (o as Record<string, unknown> | undefined)?.[k], current);
  if (typeof value !== 'string') return key;
  return value.replace(/\{(\w+)\}/g, (_, p: string) => String(params[p] ?? `{${p}}`));
}

/**
 * Boîtes « Thèmes », « Couleurs » et « Polices » (comme Mise en page d'Excel).
 * Renvoie le nouveau thème du classeur, ou null si l'utilisateur annule.
 */
import { COLOR_SCHEMES, FONT_SCHEMES, THEME_SLOTS, THEMES, sameColors, sameFonts, type Theme, type ThemeColors, type ThemeFonts } from '../data/theme';
import { t } from '../i18n';
import { box, field, openDialog, select } from '../ui/dialog';

export type ThemeMode = 'theme' | 'colors' | 'fonts';

const FONT_CHOICES = ['Calibri', 'Calibri Light', 'Aptos', 'Aptos Display', 'Arial', 'Candara', 'Cambria', 'Consolas', 'Courier New', 'Georgia', 'Helvetica', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'];

function swatches(colors: ThemeColors): HTMLElement {
  const row = document.createElement('span');
  row.className = 'g-theme-swatches';
  // Texte 1/2 puis les six accentuations, comme l'aperçu d'Excel.
  for (const i of [1, 0, 3, 2, 4, 5, 6, 7, 8, 9]) {
    const s = document.createElement('span');
    s.style.background = `#${colors.colors[i]}`;
    row.append(s);
  }
  return row;
}

/** Liste d'éléments exclusifs (un seul choisi). */
function pickList<T>(items: { value: T; label: string; extra?: HTMLElement; style?: string }[], isSelected: (v: T) => boolean, onPick: (v: T) => void, testId: string): HTMLElement {
  const list = document.createElement('div');
  list.className = 'g-theme-list';
  list.setAttribute('role', 'listbox');
  list.dataset.testid = testId;
  const refresh = () => items.forEach((it, i) => list.children[i].classList.toggle('selected', isSelected(it.value)));
  items.forEach((it) => {
    const b = document.createElement('button');
    b.type = 'button';
    b.setAttribute('role', 'option');
    b.dataset.testid = 'theme-item';
    if (it.extra) b.append(it.extra);
    const name = document.createElement('span');
    name.textContent = it.label;
    if (it.style) name.style.fontFamily = it.style;
    b.append(name);
    b.addEventListener('click', () => {
      onPick(it.value);
      refresh();
    });
    list.append(b);
  });
  refresh();
  return list;
}

export async function themeDialog(mode: ThemeMode, current: Theme): Promise<Theme | null> {
  let colors: ThemeColors = current.colors;
  let fonts: ThemeFonts = current.fonts;
  const body = box();
  const customColors = document.createElement('div');
  customColors.className = 'g-theme-custom';
  const customFonts = document.createElement('div');
  customFonts.className = 'g-row';

  const renderCustomColors = () => {
    customColors.replaceChildren();
    THEME_SLOTS.forEach((slot, i) => {
      const input = document.createElement('input');
      input.type = 'color';
      input.value = `#${colors.colors[i]}`;
      input.dataset.testid = `theme-color-${slot}`;
      input.addEventListener('input', () => {
        const next = [...colors.colors];
        next[i] = input.value.slice(1).toUpperCase();
        colors = { name: t('theme.custom'), colors: next };
        refreshLists();
      });
      customColors.append(field(t(`theme.slot.${slot}`), input));
    });
  };

  const major = select(FONT_CHOICES.map((f) => [f, f] as [string, string]), FONT_CHOICES.includes(fonts.major) ? fonts.major : undefined);
  const minor = select(FONT_CHOICES.map((f) => [f, f] as [string, string]), FONT_CHOICES.includes(fonts.minor) ? fonts.minor : undefined);
  if (!FONT_CHOICES.includes(fonts.major)) major.prepend(new Option(fonts.major, fonts.major, true, true));
  if (!FONT_CHOICES.includes(fonts.minor)) minor.prepend(new Option(fonts.minor, fonts.minor, true, true));
  const syncFonts = () => {
    major.value = fonts.major;
    minor.value = fonts.minor;
  };
  const onFontSelect = () => {
    fonts = { name: t('theme.custom'), major: major.value, minor: minor.value };
    refreshLists();
  };
  major.addEventListener('change', onFontSelect);
  minor.addEventListener('change', onFontSelect);
  major.dataset.testid = 'theme-font-major';
  minor.dataset.testid = 'theme-font-minor';
  customFonts.append(field(t('theme.headingFont'), major), field(t('theme.bodyFont'), minor));

  let lists: HTMLElement[] = [];
  const refreshLists = () => lists.forEach((l) => l.dispatchEvent(new Event('g-refresh')));

  const themeList = pickList(
    THEMES.map((th) => ({ value: th, label: th.name, extra: swatches(th.colors), style: th.fonts.minor })),
    (th) => sameColors(th.colors, colors) && sameFonts(th.fonts, fonts),
    (th) => {
      colors = th.colors;
      fonts = th.fonts;
      renderCustomColors();
      syncFonts();
    },
    'theme-list',
  );
  const colorList = pickList(
    COLOR_SCHEMES.map((c) => ({ value: c, label: c.name, extra: swatches(c) })),
    (c) => sameColors(c, colors),
    (c) => {
      colors = c;
      renderCustomColors();
    },
    'color-list',
  );
  const fontList = pickList(
    FONT_SCHEMES.map((f) => ({ value: f, label: `${f.name} (${f.major} / ${f.minor})`, style: f.minor })),
    (f) => sameFonts(f, fonts),
    (f) => {
      fonts = f;
      syncFonts();
    },
    'font-list',
  );
  // pickList ne sait pas se rafraîchir seul quand on édite une couleur ou une police : on le refait ici.
  const rebuild = (list: HTMLElement, pred: (i: number) => boolean) => [...list.children].forEach((c, i) => c.classList.toggle('selected', pred(i)));
  lists = [themeList, colorList, fontList];
  themeList.addEventListener('g-refresh', () => rebuild(themeList, (i) => sameColors(THEMES[i].colors, colors) && sameFonts(THEMES[i].fonts, fonts)));
  colorList.addEventListener('g-refresh', () => rebuild(colorList, (i) => sameColors(COLOR_SCHEMES[i], colors)));
  fontList.addEventListener('g-refresh', () => rebuild(fontList, (i) => sameFonts(FONT_SCHEMES[i], fonts)));

  renderCustomColors();
  if (mode === 'theme') body.append(themeList);
  if (mode === 'colors') body.append(colorList, Object.assign(document.createElement('p'), { textContent: t('theme.customize') }), customColors);
  if (mode === 'fonts') body.append(fontList, Object.assign(document.createElement('p'), { textContent: t('theme.customize') }), customFonts);
  body.append(Object.assign(document.createElement('p'), { className: 'g-eval-status', textContent: t('theme.hint') }));

  const ok = await openDialog({ title: t(`theme.title.${mode}`), testId: 'dialog-theme', wide: true, body, okLabel: t('theme.apply'), onOk: () => true });
  if (!ok) return null;
  const name = mode === 'theme' ? (THEMES.find((th) => sameColors(th.colors, colors) && sameFonts(th.fonts, fonts))?.name ?? t('theme.custom')) : current.name;
  return { name, colors, fonts };
}

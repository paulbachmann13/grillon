import JSZip from 'jszip';
import { exportXlsx } from '../../src/io/xlsx/export';
import { importXlsx } from '../../src/io/xlsx/import';
import { describe, expect, it } from 'vitest';
import { applyTint } from '../../src/io/xlsx/colors';
import {
  applyThemeToWorkbook,
  colorLookup,
  COLOR_SCHEMES,
  DEFAULT_THEME,
  parseTheme,
  patchThemeXml,
  remapColor,
  remapFont,
  remapStyle,
  sameTheme,
  themeOf,
  THEMES,
} from '../../src/data/theme';

const ocean = THEMES.find((t) => t.name === 'Océan')!;

describe('thèmes', () => {
  it('chaque jeu de couleurs a 12 couleurs valides', () => {
    for (const s of COLOR_SCHEMES) {
      expect(s.colors).toHaveLength(12);
      for (const c of s.colors) expect(c).toMatch(/^[0-9A-F]{6}$/);
    }
  });

  it('recolore une couleur d’accentuation et ses teintes', () => {
    const lookup = colorLookup(DEFAULT_THEME.colors);
    expect(remapColor('#4472C4', lookup, ocean.colors)).toBe('#0F6E8C');
    // « Bleu, Accentuation 1, plus clair 60 % » devient le même éclaircissement de la nouvelle couleur.
    const light = `#${applyTint('4472C4', 0.6)}`;
    expect(remapColor(light, lookup, ocean.colors)).toBe(`#${applyTint('0F6E8C', 0.6)}`);
    const dark = `#${applyTint('ED7D31', -0.25)}`;
    expect(remapColor(dark, lookup, ocean.colors)).toBe(`#${applyTint('2BA3B8', -0.25)}`);
  });

  it('laisse intactes les couleurs qui ne sont pas du thème', () => {
    const lookup = colorLookup(DEFAULT_THEME.colors);
    expect(remapColor('#123456', lookup, ocean.colors)).toBe('#123456');
    expect(remapColor('rgb(1,2,3)', lookup, ocean.colors)).toBe('rgb(1,2,3)');
  });

  it('ne touche pas au blanc ni au noir, et ne confond pas les gris avec un accent', () => {
    const lookup = colorLookup(DEFAULT_THEME.colors);
    expect(remapColor('#FFFFFF', lookup, ocean.colors)).toBe('#FFFFFF');
    expect(remapColor('#000000', lookup, ocean.colors)).toBe('#000000');
    expect(remapColor('#808080', lookup, ocean.colors)).toBe('#808080');
  });

  it('recolore fond, texte et bordures d’un style, sans modifier l’original', () => {
    const style = { bg: { rgb: '#4472C4' }, cl: { rgb: '#FFFFFF' }, bd: { t: { s: 1, cl: { rgb: '#ED7D31' } } }, ff: 'Calibri', bl: 1 };
    const out = remapStyle(style, DEFAULT_THEME, ocean);
    expect(out).toMatchObject({ bg: { rgb: '#0F6E8C' }, cl: { rgb: '#FFFFFF' }, bd: { t: { cl: { rgb: '#2BA3B8' } } }, ff: 'Arial', bl: 1 });
    expect(style.bg.rgb).toBe('#4472C4');
    expect(remapStyle({ bl: 1 }, DEFAULT_THEME, ocean)).toEqual({ bl: 1 });
  });

  it('renvoie le même objet quand il n’y a rien à changer', () => {
    const style = { bg: { rgb: '#123456' }, ff: 'Menlo' };
    expect(remapStyle(style, DEFAULT_THEME, ocean)).toBe(style);
  });

  it('remplace les polices du corps et des titres', () => {
    expect(remapFont('Calibri', DEFAULT_THEME.fonts, ocean.fonts)).toBe('Arial');
    expect(remapFont('Calibri Light', DEFAULT_THEME.fonts, ocean.fonts)).toBe('Arial');
    expect(remapFont('Menlo', DEFAULT_THEME.fonts, ocean.fonts)).toBe('Menlo');
  });

  it('applique un thème à un classeur et le mémorise', () => {
    const wb = {
      styles: { a: { bg: { rgb: '#4472C4' }, ff: 'Calibri' }, b: null },
      defaultStyle: { ff: 'Calibri', fs: 11 },
      sheets: { s1: { defaultStyle: { ff: 'Calibri', fs: 11 } } },
    };
    const out = applyThemeToWorkbook(wb, ocean);
    expect(out.styles.a).toMatchObject({ bg: { rgb: '#0F6E8C' }, ff: 'Arial' });
    expect(out.defaultStyle).toMatchObject({ ff: 'Arial', fs: 11 });
    expect(out.sheets.s1.defaultStyle).toMatchObject({ ff: 'Arial' });
    expect(sameTheme(themeOf(out), ocean)).toBe(true);
    expect(wb.styles.a.bg.rgb).toBe('#4472C4');

    // Revenir au thème d’origine rétablit les couleurs d’origine et efface la mention.
    const back = applyThemeToWorkbook(out, DEFAULT_THEME);
    expect(back.styles.a).toMatchObject({ bg: { rgb: '#4472C4' }, ff: 'Calibri' });
    expect((back.custom as { grillon: object }).grillon).toEqual({});
    expect(themeOf(back)).toBe(DEFAULT_THEME);
  });

  it('conserve les autres données de custom.grillon', () => {
    const out = applyThemeToWorkbook({ custom: { grillon: { autre: 1 }, x: 2 } }, ocean);
    expect(out.custom).toMatchObject({ x: 2, grillon: { autre: 1, theme: { name: 'Océan' } } });
  });
});

const THEME_XML = `<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements>
<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
<a:fmtScheme name="Office"></a:fmtScheme></a:themeElements></a:theme>`;

describe('theme1.xml', () => {
  it('lit le thème d’un fichier', () => {
    const colors = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
    const theme = parseTheme(THEME_XML, colors)!;
    expect(theme.colors.name).toBe('Office 2013 – 2022');
    expect(theme.fonts).toMatchObject({ major: 'Calibri Light', minor: 'Calibri' });
    expect(sameTheme(theme, DEFAULT_THEME)).toBe(true);
    expect(parseTheme(undefined, colors)).toBeUndefined();
  });

  it('remplace couleurs et polices en gardant le reste du thème', () => {
    const xml = patchThemeXml(THEME_XML, ocean);
    expect(xml).toContain('<a:accent1><a:srgbClr val="0F6E8C"/></a:accent1>');
    expect(xml).toContain('<a:clrScheme name="Océan">');
    expect(xml).toContain('<a:majorFont><a:latin typeface="Arial"/>');
    expect(xml).toContain('<a:minorFont><a:latin typeface="Arial"/>');
    expect(xml).toContain('typeface="游ゴシック Light"');
    expect(xml).toContain('<a:fmtScheme');
    // Relire le fichier modifié redonne le même thème.
    expect(parseTheme(xml, ocean.colors.colors)!.fonts).toMatchObject({ major: 'Arial', minor: 'Arial' });
  });
});

describe('thème dans le .xlsx', () => {
  const book = (custom?: object) =>
    ({
      id: 'wb',
      name: 'T',
      appVersion: '1',
      locale: 'frFR',
      styles: { a: { bg: { rgb: '#0F6E8C' }, ff: 'Arial' } },
      sheetOrder: ['s'],
      sheets: { s: { id: 's', name: 'Feuil1', cellData: { 0: { 0: { v: 'x', s: 'a' } } }, rowCount: 10, columnCount: 5 } },
      ...(custom ? { custom } : {}),
    }) as never;

  it('écrit le thème choisi dans theme1.xml et le relit', async () => {
    const bytes = await exportXlsx(book({ grillon: { theme: ocean } }));
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file('xl/theme/theme1.xml')!.async('string');
    expect(xml).toContain('<a:accent1><a:srgbClr val="0F6E8C"/></a:accent1>');
    expect(xml).toMatch(/<a:minorFont>\s*<a:latin typeface="Arial"/);
    const back = await importXlsx(bytes);
    const theme = themeOf(back as never);
    expect(theme.colors.colors).toEqual(ocean.colors.colors);
    expect(theme.fonts).toMatchObject({ major: 'Arial', minor: 'Arial' });
  });

  it('un classeur sans thème particulier garde le thème Office et n’ajoute rien', async () => {
    const bytes = await exportXlsx(book());
    const back = await importXlsx(bytes);
    expect(back.custom).toBeUndefined();
    expect(themeOf(back as never)).toBe(DEFAULT_THEME);
  });
});

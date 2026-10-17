import { describe, expect, it } from 'vitest';
import { excelStyleName, galleryStyleFor, TABLE_STYLES } from '../../src/data/table-styles';

describe('styles de tableau', () => {
  it('21 styles : Clair, Moyen et Foncé en sept couleurs', () => {
    expect(TABLE_STYLES).toHaveLength(21);
    expect(TABLE_STYLES.find((s) => s.name === 'TableStyleMedium2')?.style.headerRowStyle?.bg?.rgb).toBe('#4472C4');
  });
  it('les variantes 8 à 28 d’Excel prennent la couleur correspondante', () => {
    expect(galleryStyleFor('TableStyleMedium9')).toBe('TableStyleMedium2');
    expect(galleryStyleFor('TableStyleLight16')).toBe('TableStyleLight2');
    expect(galleryStyleFor('PivotStyleLight16')).toBe('TableStyleMedium2');
  });
  it('garde le nom Excel d’origine tant que le style n’a pas changé', () => {
    expect(excelStyleName('TableStyleMedium2', 'TableStyleMedium9')).toBe('TableStyleMedium9');
    expect(excelStyleName('TableStyleDark4', 'TableStyleMedium9')).toBe('TableStyleDark4');
    expect(excelStyleName(undefined, undefined)).toBe('TableStyleMedium2');
  });
});

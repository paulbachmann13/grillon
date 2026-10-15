import { describe, expect, it } from 'vitest';
import { EXCEL_MAX_COLUMNS, EXCEL_MAX_ROWS, grownSize } from '../../src/univer/grid-growth';

describe('agrandissement de la grille', () => {
  it('n’agrandit pas loin du bord', () => {
    expect(grownSize(1000, 500, 100, 1000, EXCEL_MAX_ROWS)).toBe(1000);
  });

  it('ajoute un pas près du bord, ou assez pour atteindre la cellule visée', () => {
    expect(grownSize(1000, 950, 100, 1000, EXCEL_MAX_ROWS)).toBe(2000);
    expect(grownSize(1000, 999, 100, 1000, EXCEL_MAX_ROWS)).toBe(2000);
    expect(grownSize(26, 25, 8, 26, EXCEL_MAX_COLUMNS)).toBe(52);
    expect(grownSize(1000, 4500, 100, 1000, EXCEL_MAX_ROWS)).toBe(5000);
  });

  it('s’arrête aux limites d’Excel', () => {
    expect(grownSize(1_048_000, 1_047_990, 100, 1000, EXCEL_MAX_ROWS)).toBe(EXCEL_MAX_ROWS);
    expect(grownSize(EXCEL_MAX_ROWS, EXCEL_MAX_ROWS - 1, 100, 1000, EXCEL_MAX_ROWS)).toBe(EXCEL_MAX_ROWS);
    expect(grownSize(16_380, 16_383, 8, 26, EXCEL_MAX_COLUMNS)).toBe(EXCEL_MAX_COLUMNS);
  });
});

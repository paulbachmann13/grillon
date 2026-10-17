import { describe, expect, it } from 'vitest';
import { cellsToRects, isRangeUnlocked, lockedState, setUnlocked, type Rect } from '../../src/data/cell-locks';

const rect = (startRow: number, startColumn: number, endRow: number, endColumn: number): Rect => ({ startRow, endRow, startColumn, endColumn });

describe('cellules déverrouillées', () => {
  it('déverrouille puis reverrouille une partie', () => {
    let rects = setUnlocked([], rect(0, 0, 9, 3), true);
    expect(isRangeUnlocked(rects, rect(2, 1, 4, 2))).toBe(true);
    rects = setUnlocked(rects, rect(5, 0, 5, 3), false);
    expect(isRangeUnlocked(rects, rect(5, 1, 5, 1))).toBe(false);
    expect(isRangeUnlocked(rects, rect(4, 0, 4, 3))).toBe(true);
    expect(isRangeUnlocked(rects, rect(6, 0, 9, 3))).toBe(true);
    expect(isRangeUnlocked(rects, rect(4, 0, 6, 0))).toBe(false);
  });

  it('indique un état mixte', () => {
    const rects = [rect(0, 0, 0, 0)];
    expect(lockedState(rects, rect(0, 0, 0, 0))).toBe(false);
    expect(lockedState(rects, rect(0, 0, 1, 0))).toBe(null);
    expect(lockedState(rects, rect(3, 3, 3, 3))).toBe(true);
  });

  it('regroupe des cellules en rectangles', () => {
    const rects = cellsToRects([[1, 1], [1, 2], [2, 1], [2, 2], [3, 1], [5, 4]]);
    expect(rects).toHaveLength(3);
    expect(isRangeUnlocked(rects, rect(1, 1, 2, 2))).toBe(true);
    expect(isRangeUnlocked(rects, rect(3, 1, 3, 1))).toBe(true);
    expect(isRangeUnlocked(rects, rect(3, 2, 3, 2))).toBe(false);
  });
});

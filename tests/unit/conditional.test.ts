import { describe, expect, it } from 'vitest';
import { CF_RESOURCE } from '../../src/io/xlsx/conditional-formatting';
import { exportXlsx } from '../../src/io/xlsx/export';
import { importXlsx } from '../../src/io/xlsx/import';
import type { IWorkbookData } from '../../src/model/univer-types';
import { buildFixture } from '../fixtures/build';
import { conditionalFormats } from './helpers';

const A = (sr: number, er: number, c: number) => ({ startRow: sr, endRow: er, startColumn: c, endColumn: c });

describe('mises en forme conditionnelles — import', () => {
  it('lit toutes les familles de règles dans l’ordre de priorité', async () => {
    const wb = await importXlsx(await buildFixture('conditional'));
    const rules = conditionalFormats(wb).MFC as { ranges: unknown[]; rule: any }[];
    expect(rules.map((r) => r.rule.type + ':' + (r.rule.subType ?? ''))).toEqual([
      'highlightCell:number',
      'highlightCell:number',
      'highlightCell:rank',
      'highlightCell:average',
      'highlightCell:text',
      'highlightCell:duplicateValues',
      'highlightCell:text',
      'highlightCell:timePeriod',
      'dataBar:',
      'colorScale:',
      'iconSet:',
      'highlightCell:formula',
    ]);
    expect(rules[0]).toMatchObject({ ranges: [A(1, 10, 0)], rule: { operator: 'greaterThan', value: 10, style: { bg: { rgb: '#FFC7CE' }, cl: { rgb: '#9C0006' } } } });
    expect(rules[1].rule).toMatchObject({ operator: 'between', value: [2, 6] });
    expect(rules[2].rule).toMatchObject({ isBottom: false, isPercent: false, value: 3, style: { bl: 1 } });
    expect(rules[3].rule).toMatchObject({ operator: 'lessThan' });
    expect(rules[4].rule).toMatchObject({ operator: 'containsText', value: 'li' });
    expect(rules[6].rule).toMatchObject({ operator: 'beginsWith', value: 'G' });
    expect(rules[7].rule).toMatchObject({ operator: 'today' });
    expect(rules[8].rule).toMatchObject({ config: { min: { type: 'min' }, max: { type: 'max' }, positiveColor: '#638EC6' } });
    expect(rules[9].rule.config).toEqual([
      { index: 0, color: '#F8696B', value: { type: 'min' } },
      { index: 1, color: '#FFEB84', value: { type: 'percentile', value: 50 } },
      { index: 2, color: '#63BE7B', value: { type: 'max' } },
    ]);
    expect(rules[10].rule.config).toEqual([
      { operator: 'greaterThanOrEqual', value: { type: 'percent', value: 67 }, iconType: '3Arrows', iconId: '0' },
      { operator: 'greaterThanOrEqual', value: { type: 'percent', value: 33 }, iconType: '3Arrows', iconId: '1' },
      { operator: 'greaterThanOrEqual', value: { type: 'percent', value: 0 }, iconType: '3Arrows', iconId: '2' },
    ]);
    expect(rules[11]).toMatchObject({ ranges: [A(12, 19, 0), A(12, 19, 2)], rule: { value: '=MOD(ROW(),2)=0' } });
  });

  it('round-trip : identique après export puis réimport', async () => {
    const first = await importXlsx(await buildFixture('conditional'));
    const second = await importXlsx(await exportXlsx(first));
    expect(conditionalFormats(second)).toEqual(conditionalFormats(first));
  });
});

describe('mises en forme conditionnelles créées dans Grillon', () => {
  it('les règles sans équivalent direct dans ExcelJS reviennent à l’identique', async () => {
    const sheetId = 's1';
    const rng = [{ ...A(0, 9, 0), sheetId }];
    const style = { bg: { rgb: 'rgb(255,235,156)' } };
    const rules = [
      { cfId: 'a', ranges: rng, stopIfTrue: false, rule: { type: 'highlightCell', subType: 'duplicateValues', style } },
      { cfId: 'b', ranges: rng, stopIfTrue: true, rule: { type: 'highlightCell', subType: 'uniqueValues', style } },
      { cfId: 'c', ranges: rng, stopIfTrue: false, rule: { type: 'highlightCell', subType: 'text', operator: 'endsWith', value: 'é"z', style } },
      { cfId: 'd', ranges: rng, stopIfTrue: false, rule: { type: 'highlightCell', subType: 'text', operator: 'notContainsText', value: 'x', style } },
      { cfId: 'e', ranges: rng, stopIfTrue: false, rule: { type: 'highlightCell', subType: 'text', operator: 'equal', value: 'Oui', style } },
      { cfId: 'f', ranges: rng, stopIfTrue: false, rule: { type: 'highlightCell', subType: 'text', operator: 'containsBlanks', style } },
      { cfId: 'g', ranges: rng, stopIfTrue: false, rule: { type: 'highlightCell', subType: 'rank', isBottom: true, isPercent: true, value: 20, style } },
      {
        cfId: 'h',
        ranges: rng,
        stopIfTrue: false,
        rule: {
          type: 'iconSet',
          isShowValue: false,
          config: [
            { operator: 'greaterThanOrEqual', value: { type: 'num', value: 80 }, iconType: '4Rating', iconId: '0' },
            { operator: 'greaterThanOrEqual', value: { type: 'num', value: 50 }, iconType: '4Rating', iconId: '1' },
            { operator: 'greaterThanOrEqual', value: { type: 'num', value: 20 }, iconType: '4Rating', iconId: '2' },
            { operator: 'greaterThanOrEqual', value: { type: 'percent', value: 0 }, iconType: '4Rating', iconId: '3' },
          ],
        },
      },
    ];
    const wb: IWorkbookData = {
      id: 'w',
      name: 'w',
      appVersion: '1',
      locale: 'frFR' as IWorkbookData['locale'],
      styles: {},
      sheetOrder: [sheetId],
      sheets: { [sheetId]: { id: sheetId, name: 'F', cellData: { 0: { 0: { v: 1 } } } } },
      resources: [{ name: CF_RESOURCE, data: JSON.stringify({ [sheetId]: rules }) }],
    };
    const back = await importXlsx(await exportXlsx(wb));
    const got = conditionalFormats(back).F as { stopIfTrue: boolean; rule: any }[];
    expect(got.map((r) => r.rule)).toEqual(
      rules.map((r) => ({ ...r.rule, ...('style' in r.rule ? { style: { bg: { rgb: '#FFEB9C' } } } : {}) })),
    );
    // « Interrompre si vrai » n'est pas géré par ExcelJS (limite connue).
    expect(got[1].stopIfTrue).toBe(false);
  });
});

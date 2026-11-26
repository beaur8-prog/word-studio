import { expect, test } from '@playwright/test';
import type { HwpDocument } from '@rhwp/core';
import { sourceSplitCellBefore } from '../src/hwp/paragraph-layout';

function proof(literal: string, storedPosition = 42.4) {
  const line = (verticalPosition: number, textStart = 0) =>
    ({ verticalPosition, textStart, height: 16, textHeight: 16, spacing: 6.4 });
  const source = {
    getCellParagraphCount: () => 2,
    getCellParaPropertiesAt: () => JSON.stringify({
      spacingBefore: 40, spacingAfter: 0, headType: 'None', lineSpacingType: 'Percent', lineSpacing: 140,
    }),
    getCellCharPropertiesAt: () => JSON.stringify({ fontSize: 1200 }),
    getTextInCellByPath: (_section: number, _paragraph: number, path: string) =>
      JSON.parse(path)[0].cellParaIndex === 1 ? literal : 'first second',
    getParagraphLineMetrics: (_section: number, _paragraph: number, path: string) =>
      JSON.stringify({ stored: true, controlCount: 0, lines: JSON.parse(path)[0].cellParaIndex === 1
        ? [line(storedPosition)] : [line(20), line(0, 6)] }),
  } as unknown as HwpDocument;
  const painted = new Map([['0:0:0:2:0', [
    { y: 121.88, h: 16, page: 0 }, { y: 101.88, h: 16, page: 1 },
  ]]]);
  return sourceSplitCellBefore(source, [0, 0, 0, 2], painted, [
    { pageIndex: 0, y: 100, h: 39.76 }, { pageIndex: 1, y: 100, h: 62.16 },
  ], 1.88, 1.88);
}

test('an authored terminal line of ordinary spaces keeps its verified paragraph leading', () => {
  const expected = { before: [20, 20], fillsFragments: true };
  expect(proof('')).toEqual(expected);
  expect(proof('   ')).toEqual(expected);
});

test('unpainted terminal text still needs a stored matching line and ordinary spaces', () => {
  expect(proof('   ', 43)).toBeUndefined();
  expect(proof(' \t')).toBeUndefined();
  expect(proof(' \n')).toBeUndefined();
  expect(proof('\u00a0')).toBeUndefined();
  expect(proof(' text')).toBeUndefined();
});

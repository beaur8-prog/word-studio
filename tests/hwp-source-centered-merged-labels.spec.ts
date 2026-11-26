import type { HwpDocument } from '@rhwp/core';
import { expect, test } from '@playwright/test';
import { sourceCenteredCellLabels } from '../src/hwp/source-centered-cell-labels';
import type { HwpPaintedLine } from '../src/hwp/paragraph-layout';
import { hwpWrapFingerprint } from '../src/hwp-text-advances';

const labels = ['금 액', '항 목'];

function readLabels(painted: HwpPaintedLine[], merged: boolean, cellLabels = labels) {
  const source = {
    getCellParagraphCount: () => cellLabels.length,
    getTextInCellByPath: (_section: number, _paragraph: number, cell: string) =>
      cellLabels[JSON.parse(cell)[0].cellParaIndex],
    getParagraphLineMetrics: () => JSON.stringify({ stored: true, controlCount: 0, lines: [{}] }),
  } as unknown as HwpDocument;
  const lines = new Map(painted.map((line, index) => ['0:3:0:2:' + index, [line]]));
  return sourceCenteredCellLabels(source, 0, 3, 0, 2, { pageIndex: 4, y: 100, h: 65, w: 67.3 }, lines, merged);
}

test('both centered labels may straddle a witnessed first fragment inside a merge', () => {
  const painted = [{ page: 4, y: 121.2, h: 13.3 }, { page: 4, y: 141.2, h: 13.3 }];
  expect(readLabels(painted, true)).toBe('67.3 ' + hwpWrapFingerprint(JSON.stringify(labels)));
  expect(readLabels(painted, false)).toBeUndefined();
});

test('both labels must occupy the same witnessed fragment and surround its center', () => {
  expect(readLabels([{ page: 4, y: 121.2, h: 13.3 }, { page: 5, y: 141.2, h: 13.3 }], true)).toBeUndefined();
  expect(readLabels([{ page: 4, y: 130, h: 13.3 }, { page: 4, y: 148, h: 13.3 }], true)).toBeUndefined();
});

test('a single numbered label needs a physically centered first row fragment', () => {
  const label = ['07 안내 항목'];
  const painted = [{ page: 4, y: 125.85, h: 13.3 }];
  expect(readLabels(painted, false, label)).toBe('67.3 ' + hwpWrapFingerprint(JSON.stringify(label)));
  expect(readLabels(painted, true, label)).toBeUndefined();
  expect(readLabels([{ ...painted[0], y: 130 }], false, label)).toBeUndefined();
});

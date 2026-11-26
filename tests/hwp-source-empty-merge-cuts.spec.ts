import { expect, test } from '@playwright/test';
import type { HwpDocument } from '@rhwp/core';
import { sourceAlignedHeaderEnds, sourceEmptyMergeCuts } from '../src/hwp/source-table-continuation';

function fixture() {
  const cell = (index: number, row: number, col: number, rowSpan: number) =>
    ({ index, row, col, rowSpan, colSpan: 1, textDirection: 0, diagonalSlash: 0, diagonalBackSlash: 0 });
  const owner = cell(0, 2, 0, 4);
  const rows = [[cell(1, 2, 1, 1)], [cell(2, 3, 1, 1)], [cell(3, 4, 1, 1)], [cell(4, 5, 1, 1)]];
  const box = (cellIdx: number, row: number, pageIndex: number, y: number, h: number, w = 240) =>
    ({ cellIdx, row, pageIndex, y, h, w });
  const boxes = [
    box(0, 2, 1, 100, 400, 80), box(0, 2, 2, 40, 460, 80), box(0, 2, 3, 40, 220, 80),
    box(1, 2, 1, 100, 400), box(1, 2, 2, 40, 150), box(2, 3, 2, 190, 90),
    box(3, 4, 2, 280, 220), box(3, 4, 3, 40, 150), box(4, 5, 3, 190, 70),
  ];
  const glyph = (page: number, y: number, charStart: number) => ({ page, y, charStart, h: 16 });
  const painted = new Map([
    ['0:0:0:1:0', [glyph(1, 470, 0), glyph(2, 43, 6)]],
    ['0:0:0:3:0', [glyph(2, 470, 0), glyph(3, 43, 6)]],
  ]);
  let ownerText = '';
  const doc = {
    getCellParagraphCount: () => 1,
    getTextInCellByPath: (_section: number, _paragraph: number, path: string) =>
      JSON.parse(path)[0].cellIndex ? 'first second' : ownerText,
    getParagraphLineMetrics: () => JSON.stringify({ stored: true, controlCount: 0,
      lines: [{ textStart: 0 }, { textStart: 6 }] }),
  } as unknown as HwpDocument;
  const body = () => ({ columns: 1, top: 40, bottom: 500 });
  return { boxes, painted, setOwnerText: (text: string) => { ownerText = text; },
    prove: () => sourceEmptyMergeCuts(doc, [0, 0, 0], owner, rows, boxes, body, painted) };
}

test('paired source row cuts are recognized beside an empty merge on three sheets', () => {
  expect(fixture().prove()).toEqual([2, 4]);
});

test('extended native row cuts require the complete header at the source body edge', () => {
  const cells = [0, 1, 2].map(col => ({ index: col, row: 0, col, rowSpan: 1, colSpan: 1, isHeader: true }));
  const fragment = (pageIndex: number, y: number) => cells.map(cell => ({
    cellIdx: cell.index, row: 0, pageIndex, y, h: 29.8,
  }));
  const body = () => ({ columns: 1, top: 113.3, bottom: 1039.4 });
  const aligned = fragment(1, 113.4);
  const deliberatelyOffset = fragment(2, 115.3);
  const incomplete = fragment(3, 113.4).slice(0, 2);
  const ends = sourceAlignedHeaderEnds(cells, [...aligned, ...deliberatelyOffset, ...incomplete], body);
  expect([...ends.keys()]).toEqual([1]);
  expect(ends.get(1)).toBeCloseTo(143.2, 6);
});

test('a single witnessed cut, different cell width, or nonempty owner keeps the original split rule', () => {
  const missing = fixture();
  missing.painted.delete('0:0:0:3:0');
  expect(missing.prove()).toEqual([]);
  const different = fixture();
  different.boxes[4].w++;
  expect(different.prove()).toEqual([]);
  const occupied = fixture();
  occupied.setOwnerText('authored text');
  expect(occupied.prove()).toEqual([]);
});

test('a long ordinary row is attested while its blank neighbour continues', () => {
  const cell = (index: number, row: number, rowSpan = 1) =>
    ({ index, row, col: index === 0 ? 0 : 1, rowSpan, colSpan: 1, textDirection: 0, diagonalSlash: 0, diagonalBackSlash: 0 });
  const owner = cell(0, 2, 3);
  const rows = [[cell(1, 2)], [cell(2, 3)], [cell(3, 4)]];
  const fragments = [[1, 100, 400], [2, 40, 460], [3, 40, 220], [4, 40, 160]];
  const boxes = [
    ...fragments.map(([pageIndex, y, h]) => ({ pageIndex, y, h, cellIdx: 0, row: 2, w: 80 })),
    ...fragments.slice(0, 3).map(([pageIndex, y, h], index) =>
      ({ pageIndex, y, h: index === 2 ? 150 : h, cellIdx: 1, row: 2, w: 240 })),
  ];
  const painted = new Map([
    ['0:0:0:1:0', [{ page: 1, y: 150, h: 16, charStart: 0 }]],
    ['0:0:0:1:1', [{ page: 2, y: 41, h: 16, charStart: 2 }]],
    ['0:0:0:1:2', [{ page: 3, y: 42, h: 16, charStart: 0 }]],
  ]);
  const literals = ['first', '  second', 'third'];
  const doc = {
    getCellParagraphCount: (_section: number, _paragraph: number, _control: number, index: number) =>
      index === 1 ? 3 : index === 2 ? 2 : 1,
    getTextInCellByPath: (_section: number, _paragraph: number, path: string) => {
      const { cellIndex, cellParaIndex } = JSON.parse(path)[0];
      return cellIndex === 1 ? literals[cellParaIndex] : cellIndex === 2 ? ['fifth', '  sixth'][cellParaIndex] : '';
    },
    getParagraphLineMetrics: () =>
      JSON.stringify({ stored: true, controlCount: 0, lines: [{ textStart: 0 }] }),
  } as unknown as HwpDocument;
  const body = () => ({ columns: 1, top: 40, bottom: 500 });
  const prove = () => sourceEmptyMergeCuts(doc, [0, 0, 0], owner, rows, boxes, body, painted);

  expect(prove()).toEqual([2]);
  literals[1] = 'A second';
  expect(prove()).toEqual([]);
  literals[1] = '  second';
  painted.get('0:0:0:1:2')![0].y = 200;
  expect(prove()).toEqual([]);
  painted.get('0:0:0:1:2')![0].y = 42;
  boxes.push({ cellIdx: 2, row: 3, pageIndex: 3, y: 190, h: 70, w: 240 },
    { cellIdx: 2, row: 3, pageIndex: 4, y: 40, h: 50, w: 240 });
  painted.set('0:0:0:2:0', [{ page: 3, y: 200, h: 16, charStart: 0 }]);
  painted.set('0:0:0:2:1', [{ page: 4, y: 44, h: 16, charStart: 2 }]);
  expect(prove()).toEqual([2, 3]);
  painted.get('0:0:0:2:1')![0].y = 180;
  expect(prove()).toEqual([2]);
});

test('an untouched following row keeps its bookmark during an observed merge cut', async ({ page }) => {
  await page.goto('/');
  const layout = await page.evaluate(async () => {
    const { canSplitInsideMergedRows } = await import('/src/table-fragments.ts');
    const table = document.createElement('table');
    table.innerHTML = '<tbody><tr data-row-split="allow" data-hwp-row-continued="true">' +
      '<td rowspan="2" data-cell-split="allow"><p></p></td><td><p>Source text</p></td></tr>' +
      '<tr data-row-split="allow"><td><p>Next<span data-docx-bookmark="Reference"></span> row</p></td></tr></tbody>';
    const group = [...table.rows];
    const bookmark = table.querySelector('[data-docx-bookmark]')!;
    const accepted = canSplitInsideMergedRows(group);
    const countBefore = table.querySelectorAll('[data-docx-bookmark]').length;
    group[0].cells[1].querySelector('p')!.append(bookmark);
    const splitCell = canSplitInsideMergedRows(group);
    group[0].cells[0].querySelector('p')!.append(bookmark);
    const owner = canSplitInsideMergedRows(group);
    return { accepted, countBefore, splitCell, owner };
  });
  expect(layout).toEqual({ accepted: true, countBefore: 1, splitCell: false, owner: false });
});

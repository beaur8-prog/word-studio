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

test('the compact native row requires the same split-cell text and authored merge', async ({page}) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { hwpCenteredCellLabelsAttribute, hwpNativeCompactRowAttribute,
      markHwpNativeCompactRow, parseHwpCenteredCellLabels } = await import('/src/hwp-centered-cell-labels.ts');
    const { hwpWrapFingerprint, normalizeHwpWrapText } = await import('/src/hwp-text-advances.ts');
    const body = document.createElement('section');
    body.style.cssText = 'width:300px;height:500px;position:absolute;left:10px;top:20px';
    body.innerHTML = '<p>Previous material</p><table data-table-anchor style="border-collapse:collapse;width:300px;table-layout:fixed">' +
      '<colgroup><col style="width:30px"><col style="width:60px"><col style="width:210px"></colgroup>' +
      '<tbody><tr><td rowspan="2">Merge</td><td>Column</td><td>Content</td></tr>' +
      '<tr data-hwp-row-continued="true" style="height:96px"><td style="vertical-align:middle"><p>금 액</p><p>항 목</p></td>' +
      '<td><p>Complete first lines</p><p>Original prefix</p></td></tr></tbody></table>';
    document.body.append(body);
    const table = body.querySelector('table')!, row = table.rows[1], [label, peer] = [...row.cells];
    try {
      const values = [label.getBoundingClientRect().width, hwpWrapFingerprint(JSON.stringify(['금 액', '항 목'])),
        'first', 65.1, peer.getBoundingClientRect().width,
        hwpWrapFingerprint(normalizeHwpWrapText(peer.textContent!))];
      const proof = values.join(' ');
      label.setAttribute(hwpCenteredCellLabelsAttribute, proof);
      const mark = () => {
        row.removeAttribute(hwpNativeCompactRowAttribute);
        markHwpNativeCompactRow(table, body);
        return row.getAttribute(hwpNativeCompactRowAttribute);
      };
      const unchanged = mark();
      peer.firstElementChild!.append('+');
      const revisedPeer = mark();
      peer.firstElementChild!.lastChild!.remove();
      label.firstElementChild!.append('+');
      const revisedLabel = mark();
      label.firstElementChild!.lastChild!.remove();
      table.rows[0].cells[0].rowSpan = 1;
      const changedMerge = mark();
      return {unchanged, revisedPeer, revisedLabel, changedMerge,
        oldProof: !!parseHwpCenteredCellLabels(values.slice(0,3).join(' ')),
        invalidSize: !!parseHwpCenteredCellLabels(values.with(3, 0).join(' '))};
    } finally { body.remove(); }
  });
  expect(result).toEqual({unchanged:'65.1', revisedPeer:null, revisedLabel:null, changedMerge:null,
    oldProof:true, invalidSize:false});
});

test('a shorter authored cut moves the extra fallback glyph and its style to the continuation', async ({page}) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { hwpCenteredCellLabelsAttribute, restoreHwpCompactPeerSplit } = await import('/src/hwp-centered-cell-labels.ts');
    const { hwpWrapFingerprint, normalizeHwpWrapText } = await import('/src/hwp-text-advances.ts');
    const table = document.createElement('table');
    table.style.cssText = 'width:300px;table-layout:fixed';
    table.innerHTML = '<tbody><tr><td><p>Label</p></td><td><p>First</p><p>Main<em>x</em>tail</p><p>End</p></td></tr>' +
      '<tr><td><p>Unused</p></td><td><p>First</p><p>Main<em>x</em></p></td></tr>' +
      '<tr><td><p>Unused</p></td><td><p>tail</p><p>End</p></td></tr></tbody>';
    document.body.append(table);
    const [label, whole] = [...table.rows[0].cells], head = table.rows[1].cells[1], tail = table.rows[2].cells[1];
    try {
      const fingerprint = (text: string) => hwpWrapFingerprint(normalizeHwpWrapText(text));
      label.setAttribute(hwpCenteredCellLabelsAttribute, [
        150, hwpWrapFingerprint('Label'), 'first', 60, Math.round(head.getBoundingClientRect().width * 100) / 100,
        fingerprint('FirstMain'), hwpWrapFingerprint(JSON.stringify(['First', 'Mainxtail', 'End'])), 1, 4,
      ].join(' '));
      const originalTail = tail.firstElementChild!.innerHTML;
      whole.children[0].append(' edit');
      restoreHwpCompactPeerSplit(label, whole, head, tail);
      const editedRejected = tail.firstElementChild!.innerHTML === originalTail;
      whole.children[0].lastChild!.remove();
      restoreHwpCompactPeerSplit(label, whole, head, tail);
      return { head: head.lastElementChild!.textContent, tail: tail.firstElementChild!.innerHTML, editedRejected,
        text: head.textContent! + tail.textContent! };
    } finally { table.remove(); }
  });
  expect(result).toEqual({head:'Main',tail:'<em>x</em>tail',editedRejected:true,text:'FirstMainxtailEnd'});
});

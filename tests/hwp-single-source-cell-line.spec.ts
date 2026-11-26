import { expect, test } from '@playwright/test';
import { hwpWrapFingerprint, normalizeHwpWrapText, parseHwpSourceWrap } from '../src/hwp-text-advances';

const text = 'ABCD EFGH IJKLMN';

test('a stored scaled cell line requires one line, a real width, and an authored scale', () => {
  const label = '결산 보고';
  const proof = {
    fingerprint: hwpWrapFingerprint(label), ends: [normalizeHwpWrapText(label).length],
    cellWidth: { outer: 100, text: 98 }, scaledCellLine: { scale: 97, size: 16, rowHeight: 30 },
  };
  const parse = (value: unknown) => parseHwpSourceWrap(JSON.stringify(value));
  expect(parse(proof)).toEqual(proof);
  expect(parse({ ...proof, ends: [2, 4] })).toBeUndefined();
  expect(parse({ ...proof, cellWidth: undefined })).toBeUndefined();
  expect(parse({ ...proof, scaledCellLine: { ...proof.scaledCellLine, scale: 100 } })).toBeUndefined();
  expect(parse({ ...proof, scaledCellLine: { ...proof.scaledCellLine, rowHeight: 14 } })).toBeUndefined();
  expect(parse({ ...proof, singleCellLine: true })).toBeUndefined();
});

test('recover a source line only in a fixed cell with room for the painted glyphs', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).waitFor();
  const rows = await page.evaluate(async ({ text, fingerprint, end }) => {
    const { layoutHwpSingleCellLines } = await import('/src/hwp-source-cell-lines.ts');
    const root = document.createElement('div');
    root.className = 'tiptap';
    document.body.append(root);
    try {
      return ['native', 'stored', 'change-text', 'change-width', 'overflows', 'autofit'].map(mode => {
        const width = mode === 'overflows' ? 116 : 125;
        root.innerHTML = `<table style="width:${width}px; table-layout:fixed; border-collapse:collapse; margin:0">
          <colgroup><col style="width:${width}px"></colgroup>
          <tbody><tr><td style="border:0;padding:0 7px 0 5px">
            <p style="margin:0;line-height:16px">heading</p>
            <p style="margin:0;line-height:16px;font-family:'Courier New',monospace;font-size:12px">${text}</p>
          </td></tr></tbody>
        </table>`;
        const table = root.querySelector('table')!;
        const caption = table.querySelector('td > p:last-child') as HTMLParagraphElement;
        if (mode === 'stored') table.dataset.wordTableLayout = 'fixed';
        if (mode === 'autofit') table.dataset.wordTableLayout = 'autofit';
        if (mode === 'change-text') caption.append('!');
        caption.dataset.hwpSourceWrap = JSON.stringify({ fingerprint, ends: [end], singleCellLine: true,
          cellWidth: { outer: width + (mode === 'change-width' ? 3 : 0), text: width - 12 } });
        const before = caption.getBoundingClientRect();
        layoutHwpSingleCellLines(table);
        const after = caption.getBoundingClientRect();
        return { mode, literal: caption.textContent, width: after.width, beforeWidth: before.width,
          applied: caption.style.whiteSpace === 'pre', before: before.height, after: after.height };
      });
    } finally { root.remove(); }
  }, { text, fingerprint: hwpWrapFingerprint(text), end: normalizeHwpWrapText(text).length });
  expect(rows.every(row => row.width === row.beforeWidth && row.literal === (row.mode === 'change-text' ? text + '!' : text))).toBe(true);
  expect(rows.filter(row => row.applied).map(row => row.mode)).toEqual(['native', 'stored']);
  expect(rows.filter(row => row.applied).every(row => row.after < row.before)).toBe(true);
  expect(rows.filter(row => !row.applied).every(row => row.after === row.before)).toBe(true);
});

test('an editable source caption keeps its click target and releases the line on edit', async ({ page }) => {
  await page.goto('/');
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  await editor.waitFor();
  const literal = 'ABCD EFGH IJKLMN XY';
  await page.evaluate(text => document.fonts.load('12px "Studio Text"', text).then(() => {}), literal);
  await editor.evaluate((root, { literal, fingerprint, end }) => {
    const view = (root as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
    const proof = JSON.stringify({ fingerprint, ends: [end], singleCellLine: true,
      cellWidth: { outer: 139, text: 132 } }).replace(/"/g, '&quot;');
    view.chain().setMeta("addToHistory", false).setContent(`<table data-word-table-layout="fixed" style="width:139px;table-layout:fixed">
      <tbody><tr><td colwidth="139" style="padding:0 3.5px;">
        <p style="margin:0;line-height:18px">heading</p>
        <p data-hwp-source-wrap="${proof}" style="margin:0;text-align:center;line-height:18px;font-family:Studio Text;font-size:12px">${literal}</p>
      </td></tr></tbody></table>`).run();
  }, { literal, fingerprint: hwpWrapFingerprint(literal), end: normalizeHwpWrapText(literal).length });
  const caption = editor.locator('p[data-hwp-source-wrap]');
  await expect(caption).toHaveAttribute('data-hwp-editor-single-cell-line', 'true');
  await expect(caption).toHaveText(literal);
  await caption.scrollIntoViewIfNeeded();
  const head = await caption.evaluate(node => {
    const tiptap = node.closest('.tiptap') as HTMLElement & { editor: import('@tiptap/core').Editor };
    const range = document.createRange();
    const last = node.lastChild as Text;
    range.setStart(last, last.length - 1);
    range.setEnd(last, last.length);
    const box = range.getBoundingClientRect();
    const edge = node.parentElement!.getBoundingClientRect().right;
    return { pen: tiptap.editor.view.posAtDOM(last, last.length),
      x: box.left + box.width * .7, y: (box.top + box.bottom) / 2, edge };
  });
  expect(head.edge - head.x).toBeLessThan(5);
  await page.mouse.move(head.x, head.y);
  await page.mouse.click(head.x, head.y);
  const selection = () => editor.evaluate(root =>
    (root as HTMLElement & { editor: import('@tiptap/core').Editor }).editor.state.selection.from);
  if (await selection() === head.pen - 1) await page.keyboard.press('ArrowRight');
  await expect.poll(selection).toBe(head.pen);
  await page.keyboard.insertText(' added');
  await expect(caption).toHaveText(`${literal} added`);
  await expect(caption).not.toHaveAttribute('data-hwp-editor-single-cell-line', 'true');
  await page.keyboard.press('ControlOrMeta+z');
  await expect(caption).toHaveText(literal);
  await expect(caption).toHaveAttribute('data-hwp-editor-single-cell-line', 'true');
});

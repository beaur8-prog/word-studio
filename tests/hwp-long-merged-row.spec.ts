import { test, expect, type Locator } from '@playwright/test';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import JSZip from 'jszip';

const { createLongMergeFixture } = createRequire(import.meta.url)('./hwp-merge-row-break-fixture.cjs');

const fixtures = [
  { count: 46, rowHeight: 1022, sourceEnds: [13, 32, 46], studioEnds: [13, 33, 46] },
  { count: 66, rowHeight: 1450, sourceEnds: [13, 32, 51, 66], studioEnds: [13, 33, 53, 66] },
  { count: 84, rowHeight: 1850, sourceEnds: [13, 32, 51, 70, 84], studioEnds: [13, 33, 53, 73, 84] },
];
const positionsFor = ({ count, studioEnds }: typeof fixtures[number]) => Array.from({ length: count + 1 }, (_, index) => ({
  marker: index + 1,
  page: index === count ? studioEnds.length : studioEnds.findIndex(end => index < end) + 1,
  fits: true,
}));
let directory: string;
const files = new Map<number, Record<string, string>>();
test.beforeAll(async () => {
  directory = await fs.mkdtemp(path.join(os.tmpdir(), 'studio-long-merged-row-'));
  for (const fixture of fixtures) files.set(fixture.count, await createLongMergeFixture(directory, fixture));
});
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

function positions(root: Locator, printed = false, count = 46) {
  return root.evaluate((root, { printed, count }) => {
    const sheet = root.closest<HTMLElement>('.document-stack');
    const origin = root.getBoundingClientRect().top;
    const scale = sheet ? sheet.getBoundingClientRect().width / parseFloat(getComputedStyle(sheet).width) : 1;
    const height = Number(sheet?.dataset.pageHeight), stride = height + 24;
    return [...root.querySelectorAll('p')].flatMap(node => {
      const text = node.textContent ?? '';
      const marker = /^TARGET RIGHT (\d+)!?$/.exec(text);
      if (!marker && text !== 'AFTER TABLE') return [];
      const range = document.createRange();
      range.selectNodeContents(node);
      const ink = range.getBoundingClientRect(), cell = node.closest('td')?.getBoundingClientRect();
      const page = printed ? Number(node.closest<HTMLElement>('.print-page')!.dataset.page)
        : Math.floor(((ink.top - origin) / scale + .5) / stride) + 1;
      const body = printed ? node.closest('.print-content')!.getBoundingClientRect() : {
        top: origin + ((page - 1) * stride + Number(sheet!.dataset.margin)) * scale,
        bottom: origin + ((page - 1) * stride + height - Number(sheet!.dataset.bottomMargin)) * scale,
      };
      return [{ marker: marker ? +marker[1] : count + 1, page, fits: ink.height > 1 &&
        ink.top >= Math.max(cell?.top ?? body.top, body.top) - 1 &&
        ink.bottom <= Math.min(cell?.bottom ?? body.bottom, body.bottom) + 1 }];
    }).sort((a, b) => a.marker - b.marker);
  }, { printed, count });
}

for (const fixture of fixtures) for (const format of ['hwp', 'hwpx']) test(`${format}: ${fixture.sourceEnds.length}-page merge preserves live lines, authored row height and two saves`, async ({ page }, info) => {
  const { count, rowHeight, sourceEnds } = fixture;
  const expected = positionsFor(fixture);
  test.setTimeout(90000);
  await page.goto('/');
  const input = page.locator('input[type=file]').first();
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  const stack = page.locator('.document-stack');
  await input.setInputFiles(files.get(count)![format]);
  await page.getByRole('button', { name: 'Word 화면에서 편집', exact: true }).click();
  await page.getByRole('button', { name: '편집본 만들기', exact: true }).click();
  const markedRow = editor.locator('tr[data-hwp-row-reserve]');
  const edited = editor.locator('td > p').filter({ hasText: /^TARGET RIGHT 34!?$/ });
  for (let stage = 0; stage < 3; stage++) {
    await expect(page.locator('.busy-indicator')).toBeHidden();
    await page.evaluate(() => document.fonts.ready);
    await expect(stack).toHaveAttribute('data-pages', String(sourceEnds.length));
    await expect(markedRow).toHaveCount(1);
    await expect.poll(() => markedRow.evaluate(row => row.style.height)).toBe(`${rowHeight}px`);
    await expect(editor.locator('table')).toHaveCount(1);
    await expect(editor.locator('tr:not(.ProseMirror-widget)')).toHaveCount(5);
    await expect(editor.locator('td[rowspan="2"]')).toHaveText(['OWNER 1OWNER 2']);
    await expect.poll(() => positions(editor, false, count)).toEqual(expected);
    if (stage === 0) {
      await edited.click({ position: { x: 10, y: 7 } });
      expect(await edited.evaluate(node => node.contains(document.getSelection()?.anchorNode ?? null))).toBe(true);
      await edited.evaluate(node => document.getSelection()!.collapse(node, node.childNodes.length));
      await page.keyboard.insertText('!');
      await expect(edited).toHaveText('TARGET RIGHT 34!');
      await expect.poll(() => positions(editor, false, count)).toEqual(expected);
      await page.keyboard.press('ControlOrMeta+z');
      await expect(edited).toHaveText('TARGET RIGHT 34');
      await page.keyboard.press('ControlOrMeta+Shift+z');
      await expect(edited).toHaveText('TARGET RIGHT 34!');
    } else {
      await expect(edited).toHaveText('TARGET RIGHT 34!');
    }
    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const preview = page.locator('.print-layout.preview');
    await expect(preview.locator('.print-page')).toHaveCount(sourceEnds.length);
    await expect.poll(() => positions(preview, true, count)).toEqual(expected);
    await expect(preview.getByText('AFTER TABLE', { exact: true })).toHaveCount(1);
    await expect(preview.getByText('OWNER 1', { exact: true })).toHaveCount(1);
    await expect(preview.getByText('OWNER 2', { exact: true })).toHaveCount(1);
    await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
    if (stage < 2) {
      await page.getByRole('button', { name: '파일', exact: true }).click();
      const download = page.waitForEvent('download');
      await page.getByRole('button', { name: 'Word 문서로 내보내기 DOCX', exact: true }).click();
      const output = info.outputPath(`textbox-save-long-merge-${count}-${format}-${stage + 1}.docx`);
      await (await download).saveAs(output);
      const archive = await JSZip.loadAsync(await fs.readFile(output));
      const xml = await archive.file('word/document.xml')!.async('string');
      const row = xml.match(/<w:tr\b(?=[^>]*wsr:reserve=)[^>]*>[\s\S]*?<\/w:tr>/)?.[0];
      const heightTag = row?.match(/<w:trHeight\b[^>]*\/>/)?.[0];
      expect(heightTag).toContain('w:hRule="atLeast"');
      expect(heightTag).toContain(`w:val="${(rowHeight - 2) * 15}"`);
      await input.setInputFiles(output);
      await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(path.basename(output, '.docx'));
    }
  }
});

test('a manually resized terminal row retires the source minimum proof', async ({ page }) => {
  await page.goto('/');
  const input = page.locator('input[type=file]').first();
  await input.setInputFiles(files.get(46)!.hwp);
  await page.getByRole('button', { name: 'Word 화면에서 편집', exact: true }).click();
  await page.getByRole('button', { name: '편집본 만들기', exact: true }).click();
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  const row = editor.locator('tr:not(.ProseMirror-widget)').nth(2);
  const cell = row.locator('td').first();
  await expect(row).toHaveAttribute('data-hwp-row-reserve', /1022 905\.9/);
  await expect.poll(() => positions(editor)).toEqual(positionsFor(fixtures[0]));
  await cell.evaluate(node => node.scrollIntoView({ block: 'end' }));
  const box = (await cell.boundingBox())!;
  const edge = { x: box.x + box.width / 2, y: box.y + box.height - 1 };
  await page.mouse.move(edge.x, edge.y);
  await expect(editor).toHaveClass(/row-resize-cursor/);
  await page.mouse.down();
  await page.mouse.move(edge.x, edge.y + 45, { steps: 6 });
  await page.mouse.up();
  await expect(row).not.toHaveAttribute('data-hwp-row-reserve');
  await expect.poll(() => row.evaluate(node => parseFloat(node.style.height))).toBeGreaterThan(1022);
  await expect.poll(async () => (await positions(editor)).filter(item => !item.fits)).toEqual([]);
  await editor.focus();
  await page.keyboard.press('ControlOrMeta+z');
  await expect(row).toHaveAttribute('data-hwp-row-reserve', /1022 905\.9/);
  await expect.poll(() => positions(editor)).toEqual(positionsFor(fixtures[0]));
});

test('the fitted reserve applies only to a continuing row with the same line geometry', async ({ page }) => {
  await page.goto('/');
  const bytes = [...await fs.readFile(files.get(46)!.hwp)];
  const proof = await page.evaluate(async bytes => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { editorExtensions } = await import('/src/extensions.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { fittedHwpRowReserve } = await import('/src/hwp-row-reserve.ts');
    const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
    const session = new HwpReaderSession();
    let model;
    try {
      await session.open(new File([Uint8Array.from(bytes)], 'row.hwp'));
      model = importEditableHwp(await session.editable(), 'row.hwp');
    } finally { session.close(); }
    const layout = await createPrintLayout(generateHTML(model.content, editorExtensions), { ...model.page, height: 400 });
    try {
      const row = layout.element.querySelector('tr[data-hwp-row-reserve]') as HTMLTableRowElement;
      const paragraph = row.querySelector('p')!;
      const text = paragraph.querySelector('span')!;
      const fit = fittedHwpRowReserve(row), lineHeight = paragraph.style.lineHeight;
      paragraph.style.lineHeight = '40px';
      const changedStyle = fittedHwpRowReserve(row);
      paragraph.style.lineHeight = lineHeight;
      text.textContent = 'This edited line wraps into several lines. '.repeat(15);
      const wrapped = fittedHwpRowReserve(row);
      return { pages: layout.pages, authored: row.style.height, fit, changedStyle, wrapped };
    } finally { layout.dispose(); }
  }, bytes);
  expect(proof).toEqual({ pages: 1, authored: '1022px', fit: 905.9, changedStyle: null, wrapped: null });
});

test('an intact HWP inline table preserves its painted borders while using the stored flow height', async ({ page }) => {
  await page.goto('/');
  const checks = await page.evaluate(async () => {
    const { reserveHwpInlineStrokes, sealHwpRowReserves } = await import('/src/hwp-row-reserve.ts');
    const root = document.createElement('div');
    root.style.cssText = 'position:absolute;top:20px;left:20px;width:400px';
    root.innerHTML = '<p style="height:22px;margin:0">A preceding line</p>' +
      '<p style="height:6px;margin:0"></p><p style="height:.14px;margin:0">&nbsp;</p>' +
      '<table style="border-collapse:collapse;width:300px;margin:0 0 18px">' +
      '<tbody><tr data-row-split="avoid" style="height:70px"><td colwidth="300" style="border:.4mm solid;vertical-align:middle">' +
      '<p>Caption</p><table><tbody><tr><td><p>Inner caption</p></td></tr></tbody></table><p>Tail</p>' +
      '</td></tr></tbody></table>';
    document.body.append(root);
    try {
      const table = root.querySelector('table')!, row = table.rows[0], cell = row.cells[0];
      const nativeHeight = cell.getBoundingClientRect().height;
      row.dataset.hwpRowReserve = `3 70 ${Math.round(nativeHeight * 100) / 100}`;
      sealHwpRowReserves(root);
      const stored = row.dataset.hwpRowReserve;
      // The authored .4mm stroke survives Word's eighth-point encoding.
      cell.style.borderWidth = '1.5px';
      const tableBox = table.getBoundingClientRect(), cellBox = cell.getBoundingClientRect();
      reserveHwpInlineStrokes(table);
      const compatible = [parseFloat(table.style.marginTop), parseFloat(table.style.marginBottom)];
      const sameHeight = cell.getBoundingClientRect().height === cellBox.height;
      const reset = () => { table.style.marginTop = '0px'; table.style.marginBottom = '18px'; };
      reset();
      const blank = table.previousElementSibling as HTMLElement;
      blank.style.backgroundColor = 'red';
      reserveHwpInlineStrokes(table);
      const withPaint = table.style.marginTop;
      blank.style.removeProperty('background-color');
      cell.style.borderWidth = '3px';
      reserveHwpInlineStrokes(table);
      return { stored, compatible, sameHeight, expected: [tableBox.top - cellBox.top, 18 - (tableBox.bottom - cellBox.bottom)],
        withPaint, modifiedBorder: table.style.marginTop };
    } finally { root.remove(); }
  });
  expect(checks.stored).toMatch(/^3 70 \d+(\.\d+)? [a-f0-9]{16}$/);
  expect(checks.expected[0]).toBeLessThan(0);
  expect(checks.compatible).toEqual(checks.expected);
  expect(checks.sameHeight).toBe(true);
  expect(checks.withPaint).toBe('0px');
  expect(checks.modifiedBorder).toBe('0px');
});

test('a measured continued row recognizes the same text runs after DOCX quantization', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { fittedHwpRowReserve, sealHwpRowReserves } = await import('/src/hwp-row-reserve.ts');
    const root = document.createElement('div');
    root.innerHTML = '<table data-hwp-centered-borders="continuation" data-hwp-centered-borders-active="true"><tbody>' +
      '<tr style="height:869.08px" data-row-split="allow" data-hwp-row-reserve="4 869.08 823.3 488">' +
      '<td colwidth="110" style="padding-bottom:7.5333px"><p style="letter-spacing:-.7px"></p></td>' +
      '<td colwidth="220" style="padding-bottom:7.5333px">' +
      '<p style="font-family:Arial;font-size:11pt;text-indent:-44.7px;line-height:19.1px">' +
      '<span style="font-family:Arial;font-size:11pt;letter-spacing:-0.7px"><span style="--word-character-scale:95">An original line</span></span>' +
      '</p></td><td colwidth="110"><p>Third column</p></td></tr></tbody></table>';
    sealHwpRowReserves(root);
    const row = root.querySelector('tr')!;
    const paragraph = row.cells[1].querySelector('p')!;
    const read = () => fittedHwpRowReserve(row, 493.64);
    const original = read();
    row.style.height = '869.1px';
    row.cells[0].querySelector('p')!.style.removeProperty('letter-spacing');
    paragraph.style.textIndent = '-44.7333px';
    paragraph.innerHTML = '<span style="font-family:Arial;font-size:11pt;letter-spacing:-0.5pt;--word-character-scale:95">An original line</span>';
    const saved = read();
    const withDifferentRemainingSpace = fittedHwpRowReserve(row, 470);
    const text = paragraph.querySelector('span')!;
    text.style.setProperty('--word-character-scale', '96');
    const editedWidth = read();
    text.style.setProperty('--word-character-scale', '95');
    paragraph.style.textIndent = '-44.5px';
    const editedIndent = read();
    paragraph.style.textIndent = '-44.7333px';
    text.textContent += ' change';
    const editedText = read();
    return { original, saved, withDifferentRemainingSpace, editedWidth, editedIndent, editedText };
  });
  expect(result).toEqual({
    original: 823.3, saved: 823.3,
    withDifferentRemainingSpace: null, editedWidth: null, editedIndent: null, editedText: null,
  });
});

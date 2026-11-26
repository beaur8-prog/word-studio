import { test, expect, type Locator } from '@playwright/test';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';

const { createInsideRowFixture } = createRequire(import.meta.url)('./hwp-merge-row-break-fixture.cjs');
const expected = [1, 1, 1, 1, 2, 2].map((page, i) => ({ index: i + 1, page, fits: true }));
let directory: string, fixtures: Record<string, string>;

test.beforeAll(async () => {
  directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-hwp-merged-continuation-'));
  fixtures = await createInsideRowFixture(directory);
});
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

function positions(root: Locator, printed = false) {
  return root.evaluate((root, printed) => {
    const sheet = root.closest<HTMLElement>('.document-stack');
    const origin = root.getBoundingClientRect().top;
    const scale = sheet ? sheet.getBoundingClientRect().width / parseFloat(getComputedStyle(sheet).width) : 1;
    const height = Number(sheet?.dataset.pageHeight), stride = height + 24;
    const values = [];
    for (const p of root.querySelectorAll('td > p')) {
      const matched = /^TARGET RIGHT (\d+)$/.exec(p.textContent ?? '');
      if (!matched) continue;
      const range = document.createRange();
      range.selectNodeContents(p);
      const ink = range.getBoundingClientRect(), cell = p.closest('td')!.getBoundingClientRect();
      const page = printed ? Number(p.closest<HTMLElement>('.print-page')!.dataset.page) :
        Math.floor(((ink.top - origin) / scale + .5) / stride) + 1;
      const body = printed ? p.closest('.print-content')!.getBoundingClientRect() : {
        top: origin + ((page - 1) * stride + Number(sheet!.dataset.margin)) * scale,
        bottom: origin + ((page - 1) * stride + height - Number(sheet!.dataset.bottomMargin)) * scale,
      };
      values.push({ index: Number(matched[1]), page,
        fits: ink.height > 1 && ink.top >= Math.max(cell.top, body.top) - 1 && ink.bottom <= Math.min(cell.bottom, body.bottom) + 1 });
    }
    return values.sort((a, b) => a.index - b.index);
  }, printed);
}

for (const format of ['hwp', 'hwpx']) test(`${format}: a row crossed by a vertical merge remains editable across the page cut after two saves`, async ({ page }, info) => {
  test.setTimeout(90000);
  await page.goto('/');
  const file = page.locator('input[type=file]').first();
  await file.setInputFiles(fixtures[format]);
  await page.getByRole('button', { name: 'Word 화면에서 편집', exact: true }).click();
  await page.getByRole('button', { name: '편집본 만들기', exact: true }).click();
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  const paragraph = editor.locator('td > p').filter({ hasText: /^TARGET RIGHT 5[!]?$/ });

  for (let state = 0; state < 3; state++) {
    await expect(page.locator('.busy-indicator')).toBeHidden();
    await page.evaluate(() => document.fonts.ready);
    await expect.poll(() => positions(editor)).toEqual(expected);
    await expect(editor.locator('table')).toHaveCount(1);
    await expect(editor.locator('tr:not(.ProseMirror-widget)')).toHaveCount(8);
    await expect(editor.locator('td[rowspan="5"]')).toHaveText('MERGED OWNER');
    await expect(editor.locator('tr[data-hwp-row-continued="true"]')).toHaveCount(1);
    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const preview = page.locator('.print-layout.preview');
    await expect(preview.locator('.print-page')).toHaveCount(2);
    await expect.poll(() => positions(preview, true)).toEqual(expected);
    await expect(preview.getByText('MERGED OWNER', { exact: true })).toHaveCount(1);
    await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();

    if (state === 0) {
      for (const zoom of [75, 125, 150, 100]) {
        const slider = page.getByRole('slider', { name: '확대 축소' });
        await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
        await expect.poll(() => positions(editor)).toEqual(expected);
      }
      await paragraph.click({ position: { x: 32, y: 8 } });
      const caretInContinuation = await paragraph.evaluate(p => p.contains(document.getSelection()?.anchorNode ?? null));
      expect(caretInContinuation).toBe(true);
      const remaining = () => paragraph.evaluate(p => {
        const caret = document.getSelection();
        if (!caret?.isCollapsed || !caret.anchorNode || !p.contains(caret.anchorNode)) return -1;
        const suffix = document.createRange();
        suffix.setStart(caret.anchorNode, caret.anchorOffset);
        suffix.setEnd(p, p.childNodes.length);
        return suffix.toString().length;
      });
      const distance = await remaining();
      expect(distance).toBeGreaterThanOrEqual(0);
      for (let move = 0; move < distance; move++) await page.keyboard.press('ArrowRight');
      await expect.poll(remaining).toBe(0);
      await page.keyboard.insertText('!');
      await expect(paragraph).toHaveText('TARGET RIGHT 5!');
      await page.keyboard.press('ControlOrMeta+z');
      await expect(paragraph).toHaveText('TARGET RIGHT 5');
      await page.keyboard.press('ControlOrMeta+Shift+z');
      await expect(paragraph).toHaveText('TARGET RIGHT 5!');
      await page.keyboard.press('ControlOrMeta+z');
      await expect.poll(() => positions(editor)).toEqual(expected);
    }
    if (state < 2) {
      await page.getByRole('button', { name: '파일', exact: true }).click();
      const download = page.waitForEvent('download');
      await page.getByRole('button', { name: 'Word 문서로 내보내기 DOCX', exact: true }).click();
      const saved = await download;
      expect(saved.suggestedFilename()).toMatch(/\.docx$/);
      const output = info.outputPath(`save-${state + 1}.docx`);
      await saved.saveAs(output);
      await file.setInputFiles(output);
      await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(`save-${state + 1}`);
    }
  }
});

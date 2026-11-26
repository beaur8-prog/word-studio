import { expect, test, type Locator, type Page } from '@playwright/test';
import type { Editor } from '@tiptap/core';

const targetLabel = 'TARGET PARAGRAPH';
const figureLabel = 'SYNTHETIC FLOAT';

async function syntheticDocument(page: Page, followingRows: number) {
  return page.evaluate(async followingRows => {
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const emu = (pixels: number) => Math.round(pixels * 9525);
    const paragraph = (text: string, height = 24) =>
      '<p data-word-line-rule="exact" style="font-family:Arial;font-size:12pt;line-height:' +
      height + 'px;margin:0">' + text + '</p>';
    const table = (label: string, rows: number, height: number) =>
      '<table style="width:510px;margin:0"><colgroup><col style="width:510px"></colgroup><tbody>' +
      Array.from({ length: rows }, (_, index) =>
        '<tr data-row-split="avoid" style="height:' + height + 'px"><td colwidth="510" ' +
        'style="padding:2px;border:1px solid #333">' +
        paragraph(label + ' ROW ' + index, 20) + '</td></tr>'
      ).join('') + '</tbody></table>';
    const raster = document.createElement('canvas');
    raster.width = raster.height = 4;
    const paint = raster.getContext('2d')!;
    paint.fillStyle = '#275997';
    paint.fillRect(0, 0, 4, 4);
    const source = raster.toDataURL();
    const picture = (label: string, y: number, height: number) => {
      const geometry = {
        horizontalPosition: { relative: 'page', offset: emu(88.30) },
        verticalPosition: { relative: 'page', offset: emu(y) },
        wrap: { type: 1, side: 'bothSides' }, layoutInCell: true, allowOverlap: true,
      };
      return '<img data-inline-image="true" src="' + source + '" width="617" height="' + height +
        '" alt="' + label + '" data-image-floating="' + JSON.stringify(geometry).replaceAll('"', '&quot;') + '">';
    };
    const html = paragraph(picture('PRECEDING FLOAT', 113.4, 80), 1) +
      paragraph('LEGITIMATE PRECEDING WRAP') + table('LEAD', 6, 51) +
      paragraph('TARGET PARAGRAPH intact when the artwork is on the next page.') +
      table('FOLLOW', followingRows, 66) + paragraph(picture('SYNTHETIC FLOAT', 524.5, 147.2), 1) +
      paragraph('TAIL PARAGRAPH');
    return {
      format: 'word-studio', version: 1, id: 'synthetic-table-float',
      title: 'Synthetic table wrap ' + followingRows, updatedAt: '2026-11-16T00:00:00Z', comments: [],
      page: { size: 'A4', orientation: 'portrait', margin: 30.00375, pageNumbers: false, header: '', footer: '' },
      content: parseDocumentHtml(html),
    };
  }, followingRows);
}

async function measure(root: Locator) {
  return root.evaluate(async (root, { targetLabel, figureLabel }) => {
    const { documentScale } = await import('/src/document-scale.ts');
    const print = root.classList.contains('preview');
    const stack = root.closest<HTMLElement>('.document-stack');
    const paragraphs = [...root.querySelectorAll('p')];
    const target = paragraphs.find(p => p.textContent?.startsWith(targetLabel))!;
    const earlier = paragraphs.find(p => p.textContent === 'LEGITIMATE PRECEDING WRAP')!;
    const image = [...root.querySelectorAll('img')].find(img => img.alt === figureLabel)!;
    const sheets = print ? root.querySelectorAll('.print-page').length : Number(stack?.dataset.pages);
    const flowWidgets = (element: HTMLElement) => element.querySelectorAll('[data-flow-widget]').length;
    function box(element: HTMLElement) {
      const sheet = (print ? element.closest<HTMLElement>('.print-page') : stack)!;
      const origin = sheet.getBoundingClientRect(), rect = element.getBoundingClientRect();
      const scale = documentScale(element), stride = Number(sheet.dataset.pageHeight) + 24;
      const top = (rect.top - origin.top) / scale;
      const page = print ? Number(sheet.dataset.page) : Math.max(1, 1 + Math.floor((top - Number(sheet.dataset.margin) + .5) / stride));
      return { page, y: top - (print ? 0 : (page - 1) * stride), height: rect.height / scale };
    }
    return {
      sheets,
      authoredBreaks: root.querySelectorAll('[data-page-break]').length,
      earlierWidgets: flowWidgets(earlier), targetWidgets: flowWidgets(target),
      target: box(target), image: box(image), anchor: box(image.closest('p')!),
      rows: paragraphs.filter(p => p.closest('td') && p.textContent?.startsWith('FOLLOW ROW '))
        .map(p => ({ text: p.textContent, page: box(p).page })),
    };
  }, { targetLabel, figureLabel });
}

type Measurement = Awaited<ReturnType<typeof measure>>;

function verifyFrame(frame: Measurement, followingRows: number, figurePage: number) {
  const samePage = figurePage === 1;
  expect(frame.sheets).toBe(figurePage);
  expect(frame.authoredBreaks).toBe(0);
  expect(frame.earlierWidgets).toBeGreaterThan(0);
  expect(frame.target.page).toBe(1);
  expect(frame.anchor.page).toBe(figurePage);
  expect(frame.image.page).toBe(figurePage);
  expect(frame.image.y).toBeCloseTo(524.5, 0);
  expect(frame.rows).toEqual(Array.from({ length: followingRows }, (_, index) => ({ text: 'FOLLOW ROW ' + index, page: 1 })));
  if (samePage) {
    expect(frame.targetWidgets).toBeGreaterThan(0);
    expect(frame.target.height).toBeGreaterThan(24);
  } else {
    expect(frame.targetWidgets).toBe(0);
    expect(frame.target.height).toBeCloseTo(24, 0);
  }
}

test('a blank floating anchor after tables does not wrap an earlier paper', async ({ page }) => {
  for (const { followingRows, figurePage } of [{ followingRows: 6, figurePage: 2 }, { followingRows: 3, figurePage: 1 }]) {
    // More table rows put the same authored square-wrap picture on paper 2.
    // The shorter fixture controls for a legitimate wrap on the same paper.
    await page.goto('/');
    const doc = await syntheticDocument(page, followingRows);
    await page.locator('input[type=file]').first().setInputFiles({
      name: 'synthetic-table-float.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(doc)),
    });
    await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(doc.title);
    await page.evaluate(() => document.fonts.ready);
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
    await expect(editor.locator('img[alt="' + figureLabel + '"]')).toHaveCount(1);
    let settled!: Measurement;
    await expect(async () => {
      const frame = await measure(editor);
      verifyFrame(frame, followingRows, figurePage);
      settled = frame;
    }).toPass({ timeout: 5000 });

    const original = await editor.evaluate(root => (root as HTMLElement & { editor: Editor }).editor.getJSON());
    const target = editor.locator('p').filter({ hasText: targetLabel });
    await target.scrollIntoViewIfNeeded();
    const textPoint = await target.evaluate(paragraph => {
      const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
      let last: Text | undefined;
      while (walker.nextNode()) {
        const node = walker.currentNode as Text;
        if (node.data.trim() && !node.parentElement?.closest('[data-flow-widget]')) last = node;
      }
      if (!last) throw new Error('Missing visible text for click');
      const end = last.data.trimEnd().length;
      const character = document.createRange();
      character.setStart(last, end - 1);
      character.setEnd(last, end);
      const box = character.getBoundingClientRect();
      return { x: box.right - box.width / 2, y: box.top + box.height / 2 };
    });
    await page.mouse.click(textPoint.x, textPoint.y);
    await page.keyboard.press('End');
    await page.keyboard.insertText(' EDIT');
    await expect(target).toContainText(' EDIT');
    await page.keyboard.press('ControlOrMeta+z');
    await expect.poll(() => editor.evaluate(root => (root as HTMLElement & { editor: Editor }).editor.getJSON())).toEqual(original);
    await expect.poll(async () => (await measure(editor)).targetWidgets).toBe(settled.targetWidgets);
    for (let sample = 0; sample < 4; sample++) {
      await page.waitForTimeout(300);
      const reopened = await measure(editor);
      verifyFrame(reopened, followingRows, figurePage);
      expect(reopened.target.height).toBeCloseTo(settled.target.height, 1);
      expect(reopened.image.y).toBeCloseTo(settled.image.y, 1);
    }

    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const preview = page.locator('.print-layout.preview');
    await expect(preview.locator('img[alt="' + figureLabel + '"]')).toHaveCount(1);
    const printed = await measure(preview);
    verifyFrame(printed, followingRows, figurePage);
    expect(printed.rows).toEqual(settled.rows);
    expect(printed.target.height).toBeCloseTo(settled.target.height, 0);
    expect(printed.target.y).toBeCloseTo(settled.target.y, 0);
  }
});

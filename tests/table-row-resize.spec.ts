import { test, expect, type Page, type Locator } from '@playwright/test';

const simple = '<table><tr><th><p>제목 A</p></th><th><p>제목 B</p></th></tr><tr><td><p>내용 A</p></td><td><p>내용 B</p></td></tr><tr><td><p>마지막 A</p></td><td><p>마지막 B</p></td></tr></table><p>표 아래 문단</p>';
async function open(page: Page, html = simple) {
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({ name: 'row-height.html', mimeType: 'text/html', buffer: Buffer.from(html) });
  await expect(page.locator('.tiptap tr')).toHaveCount(3);
}
async function edge(page: Page, cell: Locator, outside = false) {
  await cell.scrollIntoViewIfNeeded();
  const box = (await cell.boundingBox())!;
  const point = { x: box.x + box.width / 2, y: box.y + box.height + (outside ? 1 : -1) };
  await page.mouse.move(point.x, point.y);
  await expect(page.locator('.tiptap')).toHaveClass(/row-resize-cursor/);
  return point;
}
async function drag(page: Page, cell: Locator, distance: number, outside = false) {
  const point = await edge(page, cell, outside);
  await page.mouse.down(); await page.mouse.move(point.x, point.y + distance, { steps: 8 }); await page.mouse.up();
}
const height = async (row: Locator) => (await row.boundingBox())!.height;

test('horizontal dragging preserves selection, records one undo and persists to file and print', async ({ page }) => {
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await open(page);
  const rows = page.locator('.tiptap tr'), cells = page.locator('.tiptap td,.tiptap th');
  await cells.nth(0).click(); await cells.nth(3).click({ modifiers: ['Shift'] });
  await expect(page.locator('.selectedCell')).toHaveCount(4);
  const original = await height(rows.first()), second = await height(rows.nth(1));
  await drag(page, cells.first(), 80);
  await expect.poll(() => height(rows.first())).toBeCloseTo(original + 80, 0);
  await expect.poll(() => height(rows.nth(1))).toBeCloseTo(second, 0);
  await expect(page.locator('.selectedCell')).toHaveCount(4);
  await expect(page.locator('.tiptap')).toContainText('제목 A제목 B내용 A내용 B마지막 A마지막 B');
  const style = await rows.first().getAttribute('style');
  await page.keyboard.press('ControlOrMeta+z');
  await expect.poll(() => height(rows.first())).toBeCloseTo(original, 0);
  await page.keyboard.press('ControlOrMeta+y');
  await expect(rows.first()).toHaveAttribute('style', style!);
  const download = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
  await page.locator('input[type=file]').first().setInputFiles((await (await download).path())!);
  await expect(rows.first()).toHaveAttribute('style', style!);
  await page.reload();
  await expect(rows.first()).toHaveAttribute('style', style!);
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  await expect(page.locator('.print-layout tr').first()).toHaveAttribute('style', style!);
  expect(errors).toEqual([]);
});

test('Escape cancels a preview without losing a stored height; double click resets and content never clips', async ({ page }) => {
  await open(page, simple.replace('<tr>', '<tr style="height:140px">'));
  const first = page.locator('.tiptap tr').first(), cell = first.locator('th').first();
  const before = await height(first), style = await first.getAttribute('style');
  let point = await edge(page, cell);
  await page.mouse.down(); await page.mouse.move(point.x, point.y + 50, { steps: 4 });
  await expect.poll(() => height(first)).toBeGreaterThan(before + 40);
  await page.keyboard.press('Escape'); await page.mouse.up();
  await expect(first).toHaveAttribute('style', style!);
  await expect.poll(() => height(first)).toBeCloseTo(before, 0);
  await expect(page.locator('.row-height-preview')).toHaveCount(0);
  await expect(page.locator('body')).not.toHaveClass(/table-row-resizing/);
  point = await edge(page, cell);
  await page.mouse.dblclick(point.x, point.y);
  await expect(first).not.toHaveAttribute('style', /height/);
  await expect.poll(() => height(first)).toBeLessThan(before - 40);
  // The editor selection remains where it was before the gesture.
  await page.locator('.tiptap').focus(); await page.keyboard.press('ControlOrMeta+z');
  await expect(first).toHaveAttribute('style', style!);
  await drag(page, cell, -120);
  const contained = await cell.evaluate(el => el.querySelector('p')!.getBoundingClientRect().bottom <= el.getBoundingClientRect().bottom);
  expect(contained).toBe(true);
  await expect(page.locator('.tiptap')).toContainText('제목 A');
});

test('merged and final-row borders resize at 200% while column handles and read mode keep working', async ({ page }) => {
  await open(page, '<table><tr><td rowspan="2"><p>병합</p></td><td><p>첫 행</p></td></tr><tr><td><p>둘째 행</p></td></tr><tr><td><p>마지막</p></td><td><p>끝</p></td></tr></table><p>다음 문단</p>');
  const cells = page.locator('.tiptap td'), rows = page.locator('.tiptap tr');
  const box = (await cells.first().boundingBox())!;
  await page.mouse.move(box.x + box.width - 1, box.y + 15);
  await expect(page.locator('.tiptap')).toHaveClass(/resize-cursor/);
  await expect(page.locator('.tiptap')).not.toHaveClass(/row-resize-cursor/);
  await page.mouse.down(); await page.mouse.move(box.x + box.width + 30, box.y + 15, { steps: 4 }); await page.mouse.up();
  await expect(cells.first()).toHaveAttribute('colwidth', /\d+/);
  await page.getByRole('slider', { name: '확대 축소' }).focus(); await page.keyboard.press('End');
  await expect(page.getByRole('slider', { name: '확대 축소' })).toHaveValue('200');
  await cells.first().scrollIntoViewIfNeeded();
  const oldFirst = await height(rows.first()), oldSecond = await height(rows.nth(1));
  await drag(page, cells.first(), 100);
  await expect.poll(() => height(rows.nth(1))).toBeCloseTo(oldSecond + 100, 0);
  await expect.poll(() => height(rows.first())).toBeCloseTo(oldFirst, 0);
  await expect(rows.nth(1)).toHaveAttribute('style', new RegExp(`height: ${Math.round(oldSecond / 2 + 50)}px`));
  const lastCell = rows.last().locator('td').first();
  const last = await height(rows.last());
  await drag(page, lastCell, 60, true);
  await expect.poll(() => height(rows.last())).toBeCloseTo(last + 60, 0);
  const stored = await rows.last().getAttribute('style');
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '읽기 모드', exact: true }).click();
  await lastCell.scrollIntoViewIfNeeded(); const bounds = (await lastCell.boundingBox())!;
  await page.mouse.move(bounds.x + bounds.width / 2, bounds.y + bounds.height - 1);
  await expect(page.locator('.tiptap')).not.toHaveClass(/row-resize-cursor/);
  await expect(page.locator('.table-row-resize-guide')).toBeHidden();
  await page.mouse.down(); await page.mouse.move(bounds.x + bounds.width / 2, bounds.y + bounds.height + 50); await page.mouse.up();
  await expect(rows.last()).toHaveAttribute('style', stored!);
});

test('HTML and DOCX roundtrips preserve row heights as content-safe minimums', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    // @ts-expect-error Vite source module.
    const { exportDocx, importFile, sanitizeImportedHtml, validateDocument } = await import('/src/document-io.ts');
    // @ts-expect-error Vite source module.
    const { editorExtensions } = await import('/src/extensions.ts');
    // @ts-expect-error Vite dependency module.
    const { generateJSON, generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    // @ts-expect-error Vite dependency module.
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    let content = generateJSON(sanitizeImportedHtml('<table><tr style="height:72pt"><th colspan="2"><p>머리글</p></th></tr><tr style="height:144px"><td rowspan="2"><p>병합 A</p></td><td rowspan="2"><p>병합 B</p></td></tr><tr style="height:64px"></tr><tr><td colspan="2"><p>자동</p></td></tr></table>').html, editorExtensions);
    const heights = (doc: any) => doc.content[0].content.map((row: any) => row.attrs.rowHeight);
    const htmlHeights = heights(generateJSON(generateHTML(content, editorExtensions), editorExtensions));
    const source = { format: 'word-studio', version: 1, title: '행 높이', updatedAt: new Date().toISOString(),
      page: { size: 'A4', orientation: 'portrait', margin: 25.4, header: '', footer: '', pageNumbers: true }, comments: [] };
    let xml = '';
    for (let i = 0; i < 2; i++) {
      const blob = await exportDocx({ ...source, content });
      xml = await (await JSZip.loadAsync(await blob.arrayBuffer())).file('word/document.xml').async('string');
      const loaded = await importFile(new File([blob], 'heights.docx'));
      content = generateJSON(loaded.html, editorExtensions);
    }
    const invalid = structuredClone(content); invalid.content[0].content[0].attrs.rowHeight = -5;
    let rejected = false;
    try { validateDocument({ ...source, content: invalid }); } catch { rejected = true; }
    const empty = structuredClone(content); empty.content[0].content[0] = { type: 'tableRow', attrs: { rowHeight: 50 } };
    let emptyRejected = false;
    try { validateDocument({ ...source, content: empty }); } catch { emptyRejected = true; }
    return { htmlHeights, docxHeights: heights(content), xml, rejected, emptyRejected };
  });
  expect(result.htmlHeights).toEqual([96, 144, 64, null]);
  expect(result.docxHeights).toEqual([96, 144, 64, null]);
  expect(result.xml).toContain('w:hRule="atLeast"');
  expect(result.xml).not.toContain('w:hRule="exact"');
  expect(result.rejected).toBe(true);
  expect(result.emptyRejected).toBe(true);
});

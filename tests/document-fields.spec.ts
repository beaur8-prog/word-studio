import { test, expect } from '@playwright/test';

test('body PAGE and NUMPAGES fields update after page insertion, undo and redo without replacing their document nodes', async ({ page }) => {
  await page.goto('/');
  const doc = await page.evaluate(async () => {
    const { defaultPage } = await import('/src/types.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    return { format:'word-studio', version:1, title:'Page fields', comments:[], updatedAt:new Date().toISOString(), page:{ ...defaultPage, pageNumbers:false },
      content:parseDocumentHtml('<p>FIRST <span data-page-number="PAGE">37</span> of <span data-page-number="NUMPAGES">99</span></p><div data-page-break="true"></div><p>SECOND <span data-page-number="PAGE">37</span></p>') };
  });
  await page.locator('input[type=file]').first().setInputFiles({ name:'fields.wordstudio', mimeType:'application/json', buffer:Buffer.from(JSON.stringify(doc)) });
  const editor = page.getByRole('textbox', { name:'문서 편집 영역' });
  await expect.poll(() => editor.locator('[data-page-number]').allTextContents()).toEqual(['1','2','2']);
  await editor.locator('p').first().click();
  await editor.locator('p').first().evaluate(p => {
    const range=document.createRange(); range.setStart(p.firstChild!,0); range.collapse(true);
    const selection=window.getSelection()!; selection.removeAllRanges(); selection.addRange(range);
  });
  await page.getByRole('button',{name:'삽입',exact:true}).click();
  await page.getByRole('button',{name:'페이지 나누기',exact:true}).click();
  await expect.poll(() => editor.locator('[data-page-number]').allTextContents()).toEqual(['2','3','3']);
  await page.keyboard.press('ControlOrMeta+z');
  await expect.poll(() => editor.locator('[data-page-number]').allTextContents()).toEqual(['1','2','2']);
  await page.keyboard.press('ControlOrMeta+Shift+z');
  await expect.poll(() => editor.locator('[data-page-number]').allTextContents()).toEqual(['2','3','3']);
  await page.getByRole('button', {name:'보기',exact:true}).click(); await page.getByRole('button',{name:'인쇄 미리보기',exact:true}).click();
  await expect(page.locator('.print-page')).toHaveCount(3);
  expect(await page.locator('.print-content [data-page-number]').allTextContents()).toEqual(['2','3','3']);
});

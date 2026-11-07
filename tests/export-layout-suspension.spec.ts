import { test, expect } from '@playwright/test';

test('nested measurement failures release the final layout lock exactly once', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { withoutEditorReflow, editorLayoutSuspended, editorLayoutResume } = await import('/src/editor-layout-task.ts');
    const states: boolean[] = []; let resumes = 0;
    const resumed = () => resumes++;
    document.addEventListener(editorLayoutResume, resumed);
    try {
      await withoutEditorReflow(async () => {
        states.push(editorLayoutSuspended());
        try { await withoutEditorReflow(async () => { throw new Error('measurement failed'); }); } catch {}
        states.push(editorLayoutSuspended()); states.push(resumes > 0);
      });
      states.push(editorLayoutSuspended());
      return { states, resumes };
    } finally { document.removeEventListener(editorLayoutResume, resumed); }
  });
  expect(result).toEqual({ states: [true, true, false, false], resumes: 1 });
});

test('live pagination resumes after a failed measurement without changing editable text or history', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).waitFor();
  await page.evaluate(async () => {
    const { withoutEditorReflow } = await import('/src/editor-layout-task.ts');
    (window as any).pendingLayout = withoutEditorReflow(() => new Promise<void>((_resolve, reject) => { (window as any).releaseLayout = () => reject(new Error('measurement failed')); })).catch(() => {});
  });
  const text = 'LAYOUT RESUME 정부 문서의 본문과 표를 보존합니다. '.repeat(500);
  await page.locator('input[type=file]').first().setInputFiles({ name: 'layout.txt', mimeType: 'text/plain', buffer: Buffer.from(text) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  await expect(editor).toHaveText(text);
  await page.evaluate(async () => { (window as any).releaseLayout(); await (window as any).pendingLayout; });
  await expect.poll(async () => Number(await page.locator('.document-stack').getAttribute('data-pages'))).toBeGreaterThan(2);
  await expect(editor).toHaveText(text);
  await editor.click(); await page.keyboard.press('ControlOrMeta+Home'); await page.keyboard.insertText('EDIT ');
  await page.keyboard.press('ControlOrMeta+z'); await expect(editor).toHaveText(text);
});

test('two HWP table DOCX exports complete with the same live editor present and release layout afterwards', async ({ page }) => {
  test.setTimeout(90000);
  await page.goto('/');
  const source = await page.evaluate(async () => {
    const { defaultPage } = await import('/src/types.ts');
    return { id: 'export-live-layout', format: 'word-studio', version: 1, title: 'Export live table', updatedAt: new Date().toISOString(), page: defaultPage, comments: [], content: { type: 'doc', content: [{ type: 'table', content: Array.from({ length: 60 }, (_, i) => ({ type: 'tableRow', content: [{ type: 'tableCell', attrs: { colwidth: [520] }, content: [{ type: 'paragraph', attrs: { hwpCellLine: '1 Percent 16 24 16 16', lineSpacing: '16px', paragraphSize: '16px' }, content: [{ type: 'text', text: `ROW ${i} 정부 문서의 표 본문과 숫자 12345` }] }] }] })) }] } };
  });
  await page.locator('input[type=file]').first().setInputFiles({ name: 'layout.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(source)) });
  for (let save = 1; save <= 2; save++) {
    const output = await page.evaluate(async doc => {
      const { exportDocx, importFile } = await import('/src/document-io.ts');
      const { parseDocumentHtml } = await import('/src/extensions.ts');
      const { editorLayoutSuspended } = await import('/src/editor-layout-task.ts');
      const blob = await exportDocx(doc as any); const imported = await importFile(new File([blob], 'table.docx'));
      return { bytes: [...new Uint8Array(await blob.arrayBuffer())], doc: { ...imported.document, content: parseDocumentHtml(imported.html!) }, suspended: editorLayoutSuspended() };
    }, source);
    expect(output.suspended).toBe(false);
    const text = (node: any): string => (node.text || '') + (node.content || []).map(text).join('');
    expect(text(output.doc.content)).toBe(text(source.content));
    await page.locator('input[type=file]').first().setInputFiles({ name: `layout-save${save}.docx`, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(output.bytes) });
    await expect(page.getByRole('textbox', { name: '문서 편집 영역', exact: true })).toContainText('ROW 59');
    Object.assign(source, output.doc);
  }
  await expect.poll(async () => Number(await page.locator('.document-stack').getAttribute('data-pages'))).toBeGreaterThan(1);
});

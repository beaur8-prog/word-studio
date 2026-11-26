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

test('desktop Hangul saving holds live layout until the review closes', async ({ page }) => {
  await page.goto('/');
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  await editor.fill('Keep this editable during the native format review.');
  await page.evaluate(async () => {
    const { editorLayoutSuspended } = await import('/src/editor-layout-task.ts');
    const layoutStates: boolean[] = [];
    const observer = new MutationObserver(() => {
      if (/쪽 배치|문서의 쪽 나누기/.test(document.querySelector('.hwp-progress-panel strong')?.textContent ?? ''))
        layoutStates.push(editorLayoutSuspended());
    });
    const check = () => ({ layoutStates: layoutStates.slice(), suspended: editorLayoutSuspended() });
    (window as any).readHwpLayoutLease = check;
    observer.observe(document.body, { subtree: true, childList: true, characterData: true });
  });
  await page.getByRole('button', { name: '파일', exact: true }).click();
  await page.getByRole('button', { name: '한글 표준 문서로 내보내기 HWPX', exact: true }).click();
  const review = page.getByRole('dialog', { name: '한글 문서로 내보내기' });
  await expect(review).toBeVisible();
  const states = await page.evaluate(() => (window as any).readHwpLayoutLease());
  expect(states.layoutStates.length).toBeGreaterThan(0);
  expect(states.layoutStates.every(Boolean)).toBe(true);
  expect(states.suspended).toBe(true);
  await review.getByRole('button', { name: '취소', exact: true }).click();
  await expect(review).toBeHidden();
  await expect.poll(() => page.evaluate(() => (window as any).readHwpLayoutLease().suspended)).toBe(false);
  await expect(editor).toHaveText('Keep this editable during the native format review.');
});

test('native save progress does not rescan the unchanged live document', async ({ page }) => {
  await page.goto('/');
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  await editor.fill('A document stays unchanged while preparing a native save.');
  await expect(page.getByText('9개 단어', { exact: true })).toBeVisible();
  await editor.evaluate(el => {
    const live = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
    const original = live.getText.bind(live);
    let reads = 0;
    live.getText = options => { reads++; return original(options); };
    Object.assign(window, { summaryReads: () => reads });
  });
  await page.getByRole('button', { name: '파일', exact: true }).click();
  await page.getByRole('button', { name: '한글 표준 문서로 내보내기 HWPX', exact: true }).click();
  const review = page.getByRole('dialog', { name: '한글 문서로 내보내기' });
  await expect(review).toBeVisible();
  const reads = await page.evaluate(() =>
    (window as Window & typeof globalThis & { summaryReads: () => number }).summaryReads());
  expect(reads).toBe(0);
  await review.getByRole('button', { name: '취소', exact: true }).click();
  await editor.fill('Four new words now');
  await expect(page.getByText('4개 단어', { exact: true })).toBeVisible();
});

test('live Hancom body and cell metrics wait for a detached layout and resume on the current document', async ({ page }) => {
  await page.goto('/');
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  await expect(editor).toBeVisible();
  const source = await page.evaluate(async () => {
    const { defaultPage } = await import('/src/types.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { suspendEditorReflow } = await import('/src/editor-layout-task.ts');
    Object.assign(window, { resumeHancomLineMetrics: suspendEditorReflow() });
    const content = parseDocumentHtml('<p data-hwp-body-line="1 150 24 36" style="font-size:24px;line-height:36px">Body sample</p>' +
      '<table><tbody><tr><td><p data-hwp-cell-line="1 Percent 16 24 16 16" style="font-size:16px;line-height:16px">Cell sample</p></td></tr></tbody></table>');
    return { id: 'suspended-hancom-metrics', format: 'word-studio', version: 1,
      title: 'Suspended line metrics', updatedAt: new Date().toISOString(), page: defaultPage, comments: [], content };
  });
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'layout.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(source)),
  });
  await expect(editor).toContainText('Cell sample');
  const lineDecorationCounts = () => editor.evaluate(el => {
    const { state } = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
    return state.plugins.filter(plugin => /^hwp(?:Body|Cell)LineMetrics/.test((plugin as any).key))
      .map(plugin => plugin.getState(state).find().length);
  });
  await page.evaluate(() => new Promise<void>(resolve =>
    requestAnimationFrame(() => requestAnimationFrame(() => resolve()))));
  expect(await lineDecorationCounts()).toEqual([0, 0]);
  await page.evaluate(() => (window as any).resumeHancomLineMetrics());
  await expect.poll(async () => (await lineDecorationCounts()).every(count => count > 0)).toBe(true);
  await expect(editor).toContainText('Body sample');
  await expect(editor).toContainText('Cell sample');
});

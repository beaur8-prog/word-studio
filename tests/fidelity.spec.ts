import { test, expect } from '@playwright/test';

test('DOCX roundtrip preserves section margins, anchored comments and resolved state', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    // @ts-expect-error Vite serves source modules directly.
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    // @ts-expect-error Vite serves source modules directly.
    const { editorExtensions } = await import('/src/extensions.ts');
    // @ts-expect-error Vite dependency import.
    const { generateJSON } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const first = { size: 'A4', orientation: 'portrait', margin: 20,
      margins: { top: 20, right: 16, bottom: 22, left: 30 }, header: '첫 구역', footer: '보고서', pageNumbers: true };
    const second = { size: 'Letter', orientation: 'landscape', margin: 12.7, header: '둘째 구역', footer: '', pageNumbers: true };
    const doc = { format: 'word-studio', version: 1, id: 'roundtrip', title: '구역과 메모', updatedAt: new Date().toISOString(), page: first,
      comments: [{ id: 'comment-1', text: '이 문장을 검토했습니다.', quote: '같은 문장', resolved: true, createdAt: new Date().toISOString() }],
      content: { type: 'doc', content: [
        { type: 'paragraph', content: [{ type: 'text', text: '같은 문장' }] },
        { type: 'paragraph', content: [{ type: 'text', text: '같은 문장', marks: [{ type: 'commentAnchor', attrs: { ids: 'comment-1' } }] }] },
        { type: 'sectionBreak', attrs: { page: JSON.stringify(second) } },
        { type: 'paragraph', content: [{ type: 'text', text: '가로 페이지 내용' }] },
      ] } };
    const loaded = await importFile(new File([await exportDocx(doc)], 'sections.docx'));
    const content = generateJSON(loaded.html, editorExtensions);
    const again = await importFile(new File([await exportDocx({ ...loaded.document, content })], 'again.docx'));
    return { page: loaded.document.page, comments: loaded.document.comments, content, secondComments: again.document.comments, html: loaded.html };
  });
  expect(result.page.margins.left).toBeCloseTo(30, 1);
  expect(result.page.margins.right).toBeCloseTo(16, 1);
  expect(result.comments).toHaveLength(1);
  expect(result.comments[0]).toMatchObject({ text: '이 문장을 검토했습니다.', quote: '같은 문장', resolved: true });
  expect(result.secondComments[0].quote).toBe('같은 문장');
  const paragraphs = result.content.content.filter((n: { type: string }) => n.type === 'paragraph');
  expect((paragraphs[0].content[0].marks ?? []).filter((m: { type: string }) => m.type === 'commentAnchor')).toHaveLength(0);
  expect(paragraphs[1].content[0].marks.some((m: { type: string }) => m.type === 'commentAnchor')).toBe(true);
  const section = result.content.content.find((n: { type: string }) => n.type === 'sectionBreak');
  expect(JSON.parse(section.attrs.page)).toMatchObject({ size: 'Letter', orientation: 'landscape', header: '둘째 구역' });
});

test('print layout splits long Korean paragraphs and tables without losing text or overflowing', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    // @ts-expect-error Vite source module.
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const paragraph = '한글과 English 혼합 문장 👩‍💻 조합 문자를 안전하게 보존합니다. '.repeat(240);
    const rows = Array.from({ length: 90 }, (_, i) => `<tr><td><p>${i + 1}</p></td><td><p>긴 표의 내용 ${i + 1}</p></td></tr>`).join('');
    const layout = await createPrintLayout(`<p>${paragraph}</p><table><tr><th><p>번호</p></th><th><p>내용</p></th></tr>${rows}</table>`,
      { size: 'A4', orientation: 'portrait', margin: 25.4, header: '검증', footer: '끝', pageNumbers: true });
    const host = layout.element as HTMLElement;
    const fragments = Array.from(host.querySelectorAll('.print-content > p')).map(p => p.textContent).join('');
    const overflow = Array.from(host.querySelectorAll('.print-content')).filter(el => {
      const last = el.lastElementChild;
      return last && last.getBoundingClientRect().bottom > el.getBoundingClientRect().bottom + 1;
    }).length;
    const tables = Array.from(host.querySelectorAll('table'));
    host.classList.replace('measuring', 'preview');
    return { pages: layout.pages, textMatches: fragments === paragraph, overflow,
      tableParts: tables.length, repeatedHeaders: tables.every(t => t.rows[0].cells[0].textContent === '번호'),
      dataRows: tables.flatMap(t => Array.from(t.rows).slice(1)).length };
  });
  expect(result.pages).toBeGreaterThan(5);
  expect(result.textMatches).toBe(true);
  expect(result.overflow).toBe(0);
  expect(result.tableParts).toBeGreaterThan(1);
  expect(result.repeatedHeaders).toBe(true);
  expect(result.dataRows).toBe(90);
  await page.screenshot({ path: 'test-results/long-document-preview.png' });
});

test('same-titled documents keep separate recovery records and comment anchors move with edits', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    // @ts-expect-error Vite source module.
    const { saveRecovery, readHistory } = await import('/src/recovery.ts');
    const base = { format: 'word-studio', version: 1, title: '같은 제목', updatedAt: new Date().toISOString(),
      page: { size: 'A4', orientation: 'portrait', margin: 25.4, header: '', footer: '', pageNumbers: true }, comments: [],
      content: { type: 'doc', content: [{ type: 'paragraph' }] } };
    await saveRecovery({ ...base, id: 'first-document' });
    await saveRecovery({ ...base, id: 'second-document' });
    return readHistory().filter((d: { title: string }) => d.title === '같은 제목').map((d: { id: string }) => d.id);
  });
  expect(result).toEqual(['second-document', 'first-document']);
  await page.keyboard.press('Control+n');
  await page.getByRole('button', { name: /빈 문서/ }).click();
  const edit = page.locator('.tiptap');
  await expect(edit).toHaveText('');
  await edit.fill('위치를 유지하는 한글 메모');
  await page.keyboard.press('ControlOrMeta+a');
  await page.getByRole('button', { name: '검토', exact: true }).click();
  await page.getByRole('button', { name: '새 메모', exact: true }).click();
  await page.getByRole('textbox', { name: '메모 내용' }).fill('검토');
  await page.getByRole('button', { name: '메모 추가', exact: true }).click();
  await expect(edit.locator('[data-comment-ids]')).toHaveText('위치를 유지하는 한글 메모');
  await edit.click(); await page.keyboard.press(process.platform === 'darwin' ? 'Meta+ArrowUp' : 'Control+Home'); await page.keyboard.insertText('앞 문장. ');
  await expect(edit.locator('[data-comment-ids]')).toContainText('위치를 유지하는 한글 메모');
});

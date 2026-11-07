import { test, expect } from '@playwright/test';
import { mkdir, writeFile } from 'node:fs/promises';

test('missing Korean serif glyphs keep full em advances without changing the requested face or point size', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).waitFor();
  const result = await page.evaluate(async () => {
    let installed = false;
    try { await new FontFace('Native requested face', 'local("휴먼명조")').load(); installed = true; } catch {}
    const root = document.createElement('div'); root.className = 'tiptap'; document.body.append(root);
    try {
      const span = document.createElement('span'); span.style.cssText = 'font-family:휴먼명조;font-size:15pt;letter-spacing:0;white-space:pre'; span.textContent = '가나다라마바사아자차'; root.append(span);
      await document.fonts.load('15pt "Noto Serif KR"', span.textContent);
      await document.fonts.ready;
      const range = document.createRange(); range.selectNodeContents(span);
      const control = document.createElement('span'); control.style.fontFamily = 'BatangChe'; root.append(control);
      const authored = document.createElement('span'); authored.style.fontFamily = '휴먼명조, Arial'; root.append(authored);
      return { installed, width: range.getBoundingClientRect().width, size: getComputedStyle(span).fontSize, source: span.style.fontFamily, control: getComputedStyle(control).fontFamily, authored: getComputedStyle(authored).fontFamily };
    } finally { root.remove(); }
  });
  if (!result.installed) expect(result.width).toBeCloseTo(200, 0);
  expect(result.size).toBe('20px'); expect(result.source).toBe('휴먼명조');
  expect(result.control).toBe('BatangChe'); expect(result.authored).toBe('휴먼명조, Arial');
});

test('Korean serif fallback survives editing, undo, redo, cold print and two DOCX saves with original font names', async ({ page }) => {
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({ name: 'serif.html', mimeType: 'text/html', buffer: Buffer.from('<p style="font-family:휴먼명조;font-size:15pt;line-height:20pt">정부 문서의 본문과 표 내용 123</p>') });
  const p = page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).locator('p').first();
  await p.click(); await page.keyboard.press('End'); await page.keyboard.type(' EDIT');
  await page.keyboard.press('ControlOrMeta+z'); await expect(p).not.toContainText('EDIT');
  await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(p).toContainText('EDIT');
  const states = await page.evaluate(async () => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    let model: any = { format: 'word-studio', version: 1, id: 'serif', title: 'Serif', updatedAt: new Date().toISOString(), page: defaultPage, comments: [], content: parseDocumentHtml(document.querySelector('.tiptap')!.innerHTML) };
    const rows = [];
    for (let save = 0; save < 3; save++) {
      const html = generateHTML(model.content, editorExtensions), layout = await createPrintLayout(html, model.page);
      try { const el = layout.element.querySelector('.print-content p')!; rows.push({ text: el.textContent, pages: layout.pages, family: getComputedStyle(el).fontFamily, source: html, xml: '', bytes: [] as number[] }); } finally { layout.dispose(); }
      const blob = await exportDocx(model), zip = await JSZip.loadAsync(await blob.arrayBuffer()); rows.at(-1)!.xml = await zip.file('word/document.xml')!.async('string'); rows.at(-1)!.bytes = [...new Uint8Array(await blob.arrayBuffer())];
      const loaded = await importFile(new File([blob], 'serif.docx')); model = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
    }
    return rows;
  });
  await mkdir('test-results/korean-font-schema', { recursive: true });
  for (const [index, row] of states.entries()) {
    await writeFile(`test-results/korean-font-schema/serif-save${index}.docx`, Buffer.from(row.bytes));
    expect(row.text).toBe('정부 문서의 본문과 표 내용 123 EDIT'); expect(row.pages).toBe(1);
    expect(row.family).toContain('휴먼명조'); expect(row.source).not.toContain('Noto Serif');
    expect(row.xml).toContain('휴먼명조'); expect(row.xml).not.toMatch(/Noto Serif|AppleMyungjo/);
  }
});

import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';

test('a short column continuation below a heading never receives a second outer page gap', async ({ page }) => {
  await page.goto('/');
  const state = await page.evaluate(async () => {
    const { importFile } = await import('/src/document-io.ts'), { defaultPage } = await import('/src/types.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const html = '<table style="margin:0"><tr style="height:450px"><td style="padding:0"><p style="margin:0">HEADING</p></td></tr></table>' +
      '<div data-column-count="2" data-column-space="24" data-column-balance="false">' +
      Array.from({ length: 6 }, (_, i) => `<p style="line-height:80px;margin:0;font-size:10pt">LINE ${i}</p>`).join('') + '</div>';
    const loaded = await importFile(new File([html], 'continuation.html'));
    const doc = { ...loaded.document, page: { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10, pageNumbers: false }, content: parseDocumentHtml(html) };
    const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
    try { return { doc, pages: layout.pages, starts: [...layout.element.querySelectorAll('p')].map(p => ({ text: p.textContent, page: Number(p.closest<HTMLElement>('.print-page')!.dataset.page) })) }; }
    finally { layout.dispose(); }
  });
  expect(state.pages).toBe(2);
  await page.locator('input[type=file]').first().setInputFiles({ name: 'continuation.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(state.doc)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => editor.evaluate((root, zoom) => {
      const sheet = root.closest<HTMLElement>('.document-stack')!, origin = root.getBoundingClientRect().top;
      const stride = Number(sheet.dataset.pageHeight) + 24;
      return [...root.querySelectorAll('p')].map(p => ({ text: p.textContent, page: Math.floor(((p.getBoundingClientRect().top - origin) / (zoom / 100) + .5) / stride) + 1 }));
    }, zoom)).toEqual(state.starts);
    await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe(String(state.pages));
  }
  await editor.getByText('LINE 5', { exact: true }).fill('EDIT LAST');
  await page.keyboard.press('ControlOrMeta+z'); await expect(editor).toContainText('LINE 5');
  const download = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
  const saved = JSON.parse(await fs.readFile((await (await download).path())!, 'utf8'));
  expect(saved.content).toEqual(state.doc.content);
});

for (const leading of [600, 700]) test(`a table overflowing the last column starts on the same sheet as print after ${leading}px of preceding table`, async ({ page }) => {
  await page.goto('/');
  const state = await page.evaluate(async leading => {
    const { importFile } = await import('/src/document-io.ts'), { defaultPage } = await import('/src/types.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts'), { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const table = (name: string, height: number) => `<table style="margin:0"><tr style="height:${height}px"><td><p style="margin:0;line-height:20px">${name}</p></td></tr></table>`;
    const html = `<div data-column-count="2" data-column-space="24" data-column-balance="false">${table('FIRST FORM', leading)}${table('NEXT FORM', 620)}<p>TAIL</p></div>`;
    const imported = await importFile(new File([html], 'forms.html'));
    const doc = { ...imported.document, page: { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10 }, content: parseDocumentHtml(imported.html!) };
    const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
    try { return { doc, pages: layout.pages, starts: [...layout.element.querySelectorAll('p')].filter(p => p.textContent!.trim()).map(p => ({ text: p.textContent, page: Number(p.closest<HTMLElement>('.print-page')!.dataset.page) })) }; }
    finally { layout.dispose(); }
  }, leading);
  await page.locator('input[type=file]').first().setInputFiles({ name: 'forms.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(state.doc)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe(String(state.pages));
    await expect.poll(() => editor.evaluate((root, zoom) => {
      const sheet = root.closest<HTMLElement>('.document-stack')!, origin = root.getBoundingClientRect().top;
      const stride = Number(sheet.dataset.pageHeight) + 24;
      return [...root.querySelectorAll('p')].filter(p => p.textContent!.trim()).map(p => ({ text: p.textContent, page: Math.floor(((p.getBoundingClientRect().top - origin) / (zoom / 100) + .5) / stride) + 1 }));
    }, zoom)).toEqual(state.starts);
  }
  const paragraph = editor.getByText('NEXT FORM', { exact: true }); await paragraph.fill('EDIT FORM');
  await page.keyboard.press('ControlOrMeta+z'); await expect(paragraph).toHaveText('NEXT FORM');
  const download = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
  const saved = JSON.parse(await fs.readFile((await (await download).path())!, 'utf8'));
  expect(saved.content).toEqual(state.doc.content);
});

test('explicit page breaks in unbalanced columns retain all physical pages through two DOCX saves', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const html = '<div data-column-count="2" data-column-space="24" data-column-balance="false"><p>ONE</p><div data-page-break="true"></div><p>TWO</p><div data-page-break="true"></div><p>THREE</p></div>';
    const loaded = await importFile(new File([html], 'columns.html'));
    let doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; const states = [];
    for (let save = 0; save < 3; save++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try { states.push({ pages: layout.pages, text: [...layout.element.querySelectorAll('.print-content')].map(p => p.textContent), content: doc.content }); }
      finally { layout.dispose(); }
      if (save < 2) { const again = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...again.document, content: parseDocumentHtml(again.html!) }; }
    }
    return { states, doc };
  });
  for (const state of result.states) { expect(state.pages).toBe(3); expect(state.text).toEqual(['ONE', 'TWO', 'THREE']); }
  for (const state of result.states.slice(1)) {
    expect(state.content.content.filter(n => n.type === 'pageBreak')).toHaveLength(2);
    expect(state.content.content.filter(n => n.type === 'columnGroup')).toHaveLength(3);
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'columns.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.doc)) });
  await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe('3');
});

test('a final column group does not acquire an extra blank page from an outside trailing paragraph', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { importFile } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { defaultPage } = await import('/src/types.ts'), { createPrintLayout } = await import('/src/print-layout.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const html = '<div data-column-count="2" data-column-space="24" data-column-balance="false">' + Array.from({ length: 110 }, (_, i) => `<p style="line-height:20px;margin:0;font-size:10pt">LINE ${i}</p>`).join('') + '</div>';
    const loaded = await importFile(new File([html], 'end.html'));
    const doc = { ...loaded.document, page: { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10, pageNumbers: false }, content: parseDocumentHtml(html) };
    const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
    try { return { doc, pages: layout.pages }; } finally { layout.dispose(); }
  });
  expect(result.pages).toBe(2);
  await page.locator('input[type=file]').first().setInputFiles({ name: 'end.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.doc)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe('2');
  await expect(editor.locator(':scope > p')).toHaveCount(0);
  await editor.locator('p').last().fill('EDIT LAST LINE');
  await page.keyboard.press('ControlOrMeta+z'); await expect(editor.locator('p').last()).toHaveText('LINE 109');
  await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe('2');
});

for (const grouped of [false, true]) test(`a column break in ${grouped ? 'an explicit one-column section' : 'ordinary body text'} advances one page through editing and two DOCX saves`, async ({ page }) => {
  await page.goto('/');
  const state = await page.evaluate(async grouped => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const body = '<p>FIRST PAGE</p><div data-column-break="true"></div><p>SECOND PAGE</p><div data-column-break="true"></div><p>THIRD PAGE</p>';
    const html = grouped ? `<div data-column-count="1">${body}</div>` : body;
    const loaded = await importFile(new File([html], 'single-column.html'));
    let doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
    const original = doc, states = [];
    for (let save = 0; save < 3; save++) {
      const print = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try { states.push([...print.element.querySelectorAll('.print-content')].map(p => p.textContent)); }
      finally { print.dispose(); }
      if (save < 2) { const again = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...again.document, content: parseDocumentHtml(again.html!) }; }
    }
    return { doc: original, states };
  }, grouped);
  for (const pages of state.states) expect(pages).toEqual(['FIRST PAGE', 'SECOND PAGE', 'THIRD PAGE']);
  await page.locator('input[type=file]').first().setInputFiles({ name: 'single-column.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(state.doc)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe('3');
    await expect.poll(() => editor.evaluate((root, zoom) => {
      const height = Number(root.closest<HTMLElement>('.document-stack')!.dataset.pageHeight), origin = root.getBoundingClientRect().top;
      return [...root.querySelectorAll('p')].filter(p => p.textContent!.trim()).map(p => Math.floor(((p.getBoundingClientRect().top - origin) / (zoom / 100) + .5) / (height + 24)) + 1);
    }, zoom)).toEqual([1, 2, 3]);
  }
  await editor.getByText('SECOND PAGE', { exact: true }).fill('EDITED');
  await page.keyboard.press('ControlOrMeta+z'); await expect(editor).toContainText('SECOND PAGE');
  const download = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
  const saved = JSON.parse(await fs.readFile((await (await download).path())!, 'utf8'));
  expect(saved.content).toEqual(state.doc.content);
});

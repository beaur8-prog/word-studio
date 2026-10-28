import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';

for (const font of ['inherit', 'Studio Text']) test(`${font}: final glyphs fit while trailing leading stays in document flow through zoom, undo and two DOCX saves`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async font => {
    const { defaultPage } = await import('/src/types.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const paper = { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10, pageNumbers: false };
    if (font !== 'inherit') await document.fonts.load(`40pt "${font}"`);
    const body = 160 * 96 / 25.4;
    const html = `<p style="font-size:1px;margin:0;line-height:${body - 70}px"></p><p style="font-family:${font};font-size:40pt;line-height:80px;margin:0">FINAL</p>`;
    let doc = { format: 'word-studio', version: 1, id: 'fit-test', title: 'Fit', updatedAt: new Date().toISOString(), page: paper, comments: [], content: parseDocumentHtml(html) };
    const original = doc, states = [];
    for (let save = 0; save < 3; save++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), paper);
      try {
        const last = [...layout.element.querySelectorAll('p')].find(p => p.textContent === 'FINAL')!;
        const content = last.closest('.print-content')!;
        states.push({ pages: layout.pages, text: last.textContent, html: last.outerHTML, height: last.getBoundingClientRect().height,
          leadingOutside: last.getBoundingClientRect().bottom - content.getBoundingClientRect().bottom });
      } finally { layout.dispose(); }
      if (save < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
    }
    const guards = [];
    for (const tail of ['<p style="font-size:40pt;line-height:80px">FINAL</p><p>AFTER</p>', '<p style="font-size:40pt;line-height:80px"><mark>FINAL</mark></p>', '<p style="font-size:40pt;line-height:80px">FINAL<br></p>']) {
      const layout = await createPrintLayout(generateHTML(parseDocumentHtml(html.replace(/<p style="font-family:[\s\S]*$/, tail)), editorExtensions), paper);
      try { guards.push({ pages: layout.pages, first: layout.element.querySelector('.print-page')!.textContent }); } finally { layout.dispose(); }
    }
    return { original, states, guards };
  }, font);
  if (process.env.HWP_LAYOUT_DIAGNOSTICS) console.log(JSON.stringify(result));
  for (const state of result.states) {
    expect(state.pages).toBe(1); expect(state.text).toBe('FINAL');
    expect(state.height).toBeCloseTo(80, 0); expect(state.leadingOutside).toBeGreaterThan(9);
  }
  expect(result.guards[0].pages).toBe(2); expect(result.guards[0].first).toContain('FINAL'); expect(result.guards[0].first).not.toContain('AFTER');
  expect(result.guards[1].pages).toBe(2); expect(result.guards[2].pages).toBe(2);
  await page.locator('input[type=file]').first().setInputFiles({ name: 'fit.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.original)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  for (const zoom of [50, 100, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe('1');
  }
  const last = editor.locator('p').filter({ hasText: /^FINAL$/ });
  await last.fill('WRAP '.repeat(100)); await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).not.toBe('1');
  await page.keyboard.press('ControlOrMeta+z'); await expect(last).toHaveText('FINAL'); await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe('1');
  const download = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
  const saved = JSON.parse(await fs.readFile((await (await download).path())!, 'utf8'));
  expect(saved.content).toEqual(result.original.content);
});

test('the first print loads a requested bundled face before measuring pages', async ({ page }) => {
  await page.route('**/NotoSansMono.ttf', async route => { await new Promise(resolve => setTimeout(resolve, 150)); await route.continue(); });
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { defaultPage } = await import('/src/types.ts'), { createPrintLayout } = await import('/src/print-layout.ts');
    const html = '<p style="font-family:Studio Text;font-size:40pt;line-height:80px">FIRST PRINT</p>';
    const states = [];
    for (let i = 0; i < 2; i++) {
      const layout = await createPrintLayout(html, defaultPage);
      try { states.push({ loaded: document.fonts.check('40pt "Studio Text"'), pages: layout.pages, height: layout.element.querySelector('.print-content p')!.getBoundingClientRect().height }); }
      finally { layout.dispose(); }
      await document.fonts.ready;
    }
    return states;
  });
  expect(result[0].loaded).toBe(true); expect(result[0]).toEqual(result[1]);
});

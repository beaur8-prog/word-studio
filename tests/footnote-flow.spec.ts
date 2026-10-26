import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';

for (const kind of ['paragraph', 'row', 'merge']) test(`a previous footnote reserves space while a long ${kind} continues, including zoom and undo`, async ({ page }) => {
  await page.goto('/');
  const state = await page.evaluate(async kind => {
    const { defaultPage } = await import('/src/types.ts'), { importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const p = (text: string) => `<p style="font-size:10pt;line-height:20px;margin:0">${text}</p>`;
    const reference = '<span data-note-reference="foot" data-note-kind="footnote" data-note-label="1">1</span>';
    const body = kind === 'paragraph' ? p(Array.from({ length: 420 }, (_, i) => `Token${String(i).padStart(3, '0')} `).join(''))
      : kind === 'row' ? `<table style="width:400px;margin:0"><colgroup><col style="width:200px"><col style="width:200px"></colgroup><tr data-row-split="allow"><td style="padding:0">${Array.from({ length: 42 }, (_, i) => p('LEFT ' + i)).join('')}</td><td style="padding:0">${Array.from({ length: 38 }, (_, i) => p('RIGHT ' + i)).join('')}</td></tr></table>`
      : `<table style="width:400px;margin:0"><colgroup><col style="width:100px"><col style="width:300px"></colgroup>${Array.from({ length: 6 }, (_, i) => `<tr style="height:140px">${i ? '' : '<td rowspan="6" data-cell-split="allow" style="vertical-align:top"><p>MERGED ONCE</p></td>'}<td style="vertical-align:top">${p('ROW ' + i)}</td></tr>`).join('')}</table>`;
    const html = p('REFERENCE' + reference) + `<p style="line-height:100px;margin:0">LEAD</p>` + body + `<div data-document-note="foot" data-note-kind="footnote" data-note-label="1">${Array.from({ length: 6 }, (_, i) => p('NOTE ' + i)).join('')}</div><p></p>`;
    const loaded = await importFile(new File([html], 'notes.html'));
    const doc = { ...loaded.document, page: { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10 }, content: parseDocumentHtml(loaded.html!) };
    const print = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
    try { return { doc, pages: [...print.element.querySelectorAll('.print-content')].map(p => p.textContent!.replace(/\s/g, '')), notes: print.element.querySelector('.print-footnotes')!.textContent }; }
    finally { print.dispose(); }
  }, kind);
  expect(state.notes).toContain('NOTE 5');
  await page.locator('input[type=file]').first().setInputFiles({ name: 'notes.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(state.doc)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => editor.evaluate((root, zoom) => {
      const sheet = root.closest<HTMLElement>('.document-stack')!, origin = root.getBoundingClientRect().top, scale = zoom / 100;
      const height = Number(sheet.dataset.pageHeight), top = Number(sheet.dataset.margin), bottom = Number(sheet.dataset.bottomMargin);
      const note = root.querySelector<HTMLElement>('[data-document-note]')!, noteBox = note.getBoundingClientRect();
      const noteTop = (noteBox.top - origin) / scale;
      const notePage = Math.max(0, Math.floor((noteTop + .5) / (height + 24)));
      const pages: string[] = [], walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) {
        const node = walker.currentNode as Text;
        if (node.parentElement?.closest('.pagination-cell-gap,.pagination-gap,[data-document-note]')) continue;
        for (let i = 0; i < node.length; i++) {
          if (/\s/.test(node.data[i])) continue;
          const range = document.createRange(); range.setStart(node, i); range.setEnd(node, i + 1);
          const rect = range.getBoundingClientRect(), y = (rect.top - origin) / scale;
          const page = Math.max(0, Math.floor((y + .5) / (height + 24))), local = y - page * (height + 24);
          if (local < top - 1 || local + rect.height / scale > height - bottom + 1) return ['INK IN MARGIN'];
          if (page === notePage && y + rect.height / scale > noteTop + .5) return ['BODY OVERLAPS FOOTNOTE'];
          pages[page] = (pages[page] || '') + node.data[i];
        }
      }
      return pages;
    }, zoom)).toEqual(state.pages);
  }
  await editor.locator(':scope > p').first().click({ position: { x: 8, y: 8 } }); await page.keyboard.press('Home');
  await page.keyboard.type('EDIT'); await expect(editor).toContainText('EDIT');
  await page.keyboard.press('ControlOrMeta+z'); await expect(editor).not.toContainText('EDIT');
  const download = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
  const saved = JSON.parse(await fs.readFile((await (await download).path())!, 'utf8'));
  expect(saved.content).toEqual(state.doc.content);
});

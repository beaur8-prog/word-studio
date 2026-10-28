import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';

for (const [kind, font] of [['paragraph', 'inherit'], ['paragraph', 'Studio Text'], ['row', 'inherit'], ['merge', 'inherit'],
  ['inline-start', 'inherit'], ['inline-middle', 'inherit'], ['inline-end', 'inherit'],
  ['row-note-start', 'inherit'], ['row-note-middle', 'inherit'], ['row-note-end', 'inherit'], ['row-note-boundary', 'inherit'], ['row-note-multiple', 'inherit']]) test(`${font}: footnote space follows its reference while a long ${kind} continues, including zoom and undo`, async ({ page }) => {
  await page.goto('/');
  if (font !== 'inherit') await page.addStyleTag({ content: `.tiptap { font-family: '${font}' !important; }` });
  const state = await page.evaluate(async kind => {
    const { defaultPage } = await import('/src/types.ts'), { importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const p = (text: string) => `<p style="font-size:10pt;line-height:20px;margin:0">${text}</p>`;
    const reference = '<span data-note-reference="foot" data-note-kind="footnote" data-note-label="1">1</span>';
    const reference2 = '<span data-note-reference="second" data-note-kind="footnote" data-note-label="2">2</span>';
    const rowReference = kind.startsWith('row-note'), ownReference = kind.startsWith('inline') || rowReference, at = kind === 'inline-start' ? 0 : kind === 'inline-middle' ? 120 : 300;
    const body = kind === 'paragraph' || kind.startsWith('inline') ? p(Array.from({ length: 420 }, (_, i) => `${ownReference && i === at ? reference : ''}Token${String(i).padStart(3, '0')} `).join(''))
      : kind === 'row' || rowReference ? `<table style="width:400px;margin:0"><colgroup><col style="width:200px"><col style="width:200px"></colgroup><tr data-row-split="allow"><td style="padding:0">${Array.from({ length: 42 }, (_, i) => p('LEFT ' + i + (kind === 'row-note-multiple' && i === 22 ? reference2 : ''))).join('')}</td><td style="padding:0">${Array.from({ length: 38 }, (_, i) => p('RIGHT ' + i + (rowReference && i === (kind === 'row-note-start' || kind === 'row-note-multiple' ? 0 : kind === 'row-note-middle' ? 15 : kind === 'row-note-boundary' ? 22 : 32) ? reference : ''))).join('')}</td></tr></table>`
      : `<table style="width:400px;margin:0"><colgroup><col style="width:100px"><col style="width:300px"></colgroup>${Array.from({ length: 6 }, (_, i) => `<tr style="height:140px">${i ? '' : '<td rowspan="6" data-cell-split="allow" style="vertical-align:top"><p>MERGED ONCE</p></td>'}<td style="vertical-align:top">${p('ROW ' + i)}</td></tr>`).join('')}</table>`;
    const html = (ownReference ? '' : p('REFERENCE' + reference)) + `<p style="line-height:100px;margin:0">LEAD</p>` + body + `<div data-document-note="foot" data-note-kind="footnote" data-note-label="1">${Array.from({ length: 6 }, (_, i) => p('NOTE ' + i)).join('')}</div>${kind === 'row-note-multiple' ? `<div data-document-note="second" data-note-kind="footnote" data-note-label="2">${Array.from({ length: 3 }, (_, i) => p('SECOND NOTE ' + i)).join('')}</div>` : ''}<p></p>`;
    const loaded = await importFile(new File([html], 'notes.html'));
    const doc = { ...loaded.document, page: { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10 }, content: parseDocumentHtml(loaded.html!) };
    const print = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
    try { return { doc, pages: [...print.element.querySelectorAll('.print-content')].map(p => p.textContent!.replace(/\s/g, '')), notes: [...print.element.querySelectorAll('.print-footnotes')].map(n => n.textContent).join('') }; }
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
      const notes = [...root.querySelectorAll<HTMLElement>('[data-document-note]')].map(note => {
        const top = (note.getBoundingClientRect().top - origin) / scale;
        return { id: note.dataset.documentNote, top, page: Math.max(0, Math.floor((top + .5) / (height + 24))) };
      });
      for (const reference of root.querySelectorAll<HTMLElement>('[data-note-reference]')) {
        const y = (reference.getBoundingClientRect().top - origin) / scale;
        if (notes.find(note => note.id === reference.dataset.noteReference)?.page !== Math.max(0, Math.floor((y + .5) / (height + 24)))) return ['FOOTNOTE ON WRONG PAGE'];
      }
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
          const noteTop = Math.min(Infinity, ...notes.filter(note => note.page === page).map(note => note.top));
          if (y + rect.height / scale > noteTop + .5) return [`BODY OVERLAPS FOOTNOTE: ${node.data.slice(Math.max(0, i - 12), i + 15)} at ${y.toFixed(2)} + ${(rect.height / scale).toFixed(2)}; note ${noteTop.toFixed(2)}`];
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

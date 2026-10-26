import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';

for (const font of ['inherit', 'Studio Text']) test(`${font}: a body paragraph continues at printed line boundaries across zoom, editing and undo without splitting the model`, async ({ page }) => {
  await page.goto('/');
  if (font !== 'inherit') await page.addStyleTag({ content: `.tiptap { font-family: '${font}' !important; }` });
  const state = await page.evaluate(async () => {
    const { defaultPage } = await import('/src/types.ts'), { importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const text = Array.from({ length: 240 }, (_, i) => `Token${String(i).padStart(3, '0')} `).join('');
    const loaded = await importFile(new File([`<p style="line-height:400px;margin:0">LEAD</p><p style="font-size:10pt;line-height:20px;margin:0;text-indent:12px">${text}</p>`], 'flow.html'));
    const doc = { ...loaded.document, page: { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10 }, content: parseDocumentHtml(loaded.html!) };
    const print = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
    try { const p = print.element.querySelectorAll('.print-content p')[1] as HTMLElement; const range = document.createRange(); range.selectNodeContents(p); return { doc, metrics: { css: p.style.cssText, font: getComputedStyle(p).font, width: p.getBoundingClientRect().width, height: p.getBoundingClientRect().height, rects: [...range.getClientRects()].slice(0,3).map(r => ({x:r.x,y:r.y,w:r.width,h:r.height})) }, pages: [...print.element.querySelectorAll('.print-content')].map(p => p.textContent!.replace(/\s/g, '')) }; }
    finally { print.dispose(); }
  });
  await page.locator('input[type=file]').first().setInputFiles({ name: 'flow.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(state.doc)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  if (process.env.HWP_LAYOUT_DIAGNOSTICS) console.log('PRINT METRICS', state.metrics);
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    if (process.env.HWP_LAYOUT_DIAGNOSTICS) console.log('EDITOR METRICS', zoom, await editor.evaluate(root => { const p = root.querySelectorAll('p')[1] as HTMLElement, range = document.createRange(); range.selectNodeContents(p); return { css:p.style.cssText, font:getComputedStyle(p).font, width:p.getBoundingClientRect().width, height:p.getBoundingClientRect().height, rects:[...range.getClientRects()].slice(0,3).map(r=>({x:r.x,y:r.y,w:r.width,h:r.height})) }; }));
    await expect.poll(() => editor.evaluate((root, zoom) => {
      const sheet = root.closest<HTMLElement>('.document-stack')!, origin = root.getBoundingClientRect().top;
      const height = Number(sheet.dataset.pageHeight), top = Number(sheet.dataset.margin), bottom = Number(sheet.dataset.bottomMargin);
      const pages: string[] = [], walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) {
        const node = walker.currentNode as Text;
        if (node.parentElement?.closest('.pagination-cell-gap,.pagination-gap')) continue;
        for (let i = 0; i < node.length; i++) {
          if (/\s/.test(node.data[i])) continue;
          const range = document.createRange(); range.setStart(node, i); range.setEnd(node, i + 1);
          const rect = range.getBoundingClientRect(), y = (rect.top - origin) / (zoom / 100);
          const page = Math.max(0, Math.floor((y + .5) / (height + 24))), local = y - page * (height + 24);
          if (local < top - 1 || local + rect.height / (zoom / 100) > height - bottom + 1) return ['INK IN MARGIN'];
          pages[page] = (pages[page] || '') + node.data[i];
        }
      }
      return pages;
    }, zoom)).toEqual(state.pages);
  }
  await editor.click(); await page.keyboard.press(process.platform === 'darwin' ? 'Meta+ArrowDown' : 'Control+End');
  await page.keyboard.type('EDIT'); await expect(editor).toContainText('EDIT');
  await page.keyboard.press('ControlOrMeta+z'); await expect(editor).not.toContainText('EDIT');
  const download = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
  const saved = JSON.parse(await fs.readFile((await (await download).path())!, 'utf8'));
  expect(saved.content).toEqual(state.doc.content);
});

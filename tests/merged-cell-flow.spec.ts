import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';

for (const alignment of ['top', 'middle', 'bottom']) for (const kind of ['paragraphs', 'wrapped', 'four-pages', 'margins', 'long-paragraph']) test(`${alignment} ${kind}: rich vertical merge keeps paragraphs and pictures on the same pages in editor and print`, async ({ page }) => {
  await page.goto('/');
  const state = await page.evaluate(async ({ alignment, kind }) => {
    const { defaultPage } = await import('/src/types.ts'), { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const p = (text: string) => `<p style="font-size:10pt;line-height:20px;margin:${kind === 'margins' ? '3px 0 7px' : '0'}">${text}</p>`;
    const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';
    const picture = `<img data-inline-image="true" src="data:image/png;base64,${png}" width="80" height="40" alt="MERGED IMAGE">`;
    const rows = kind === 'four-pages' ? 12 : 6;
    const left = kind === 'long-paragraph' ? p(Array.from({ length: 120 }, (_, i) => `LEFT${i} ` + (i === 90 ? picture : '')).join('')) : kind === 'wrapped' ? p(Array.from({ length: 60 }, (_, i) => `LEFT${i} ` + (i === 30 ? picture : '')).join('')) + p(Array.from({ length: 60 }, (_, i) => `AFTER${i} `).join('')) : Array.from({ length: kind === 'four-pages' ? 70 : kind === 'margins' ? 20 : 28 }, (_, i) => p(`LEFT ${i}`) + (i === 13 ? p(picture) : '')).join('');
    const table = `<table style="width:400px;margin:0"><colgroup><col style="width:160px"><col style="width:240px"></colgroup>${Array.from({ length: rows }, (_, i) => `<tr style="height:140px">${i ? '' : `<td rowspan="${rows}" data-cell-split="allow" style="vertical-align:${alignment};padding:4px">${left}</td>`}<td style="vertical-align:top;padding:4px">${p('RIGHT ' + i)}</td></tr>`).join('')}</table>`;
    const loaded = await importFile(new File([`<p style="line-height:100px;margin:0">LEAD</p>${table}<p></p>`], 'merge.html'));
    const doc = { ...loaded.document, page: { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10 }, content: parseDocumentHtml(loaded.html!) };
    const print = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
    try {
      const ink = [...print.element.querySelectorAll<HTMLElement>('.print-content')].map(content => {
        const top = content.closest('.print-page')!.getBoundingClientRect().top, result: number[] = [];
        const walker = document.createTreeWalker(content, NodeFilter.SHOW_TEXT);
        while (walker.nextNode()) {
          const node = walker.currentNode as Text;
          for (let i = 0; i < node.length; i++) {
            if (/\s/.test(node.data[i])) continue;
            const range = document.createRange(); range.setStart(node, i); range.setEnd(node, i + 1);
            result.push(range.getBoundingClientRect().top - top);
          }
        }
        return result;
      });
      const roundtrips = []; let current = doc;
      for (let save = 0; save < 2; save++) {
        const loaded = await importFile(new File([await exportDocx(current)], 'again.docx'));
        current = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
        const layout = await createPrintLayout(generateHTML(current.content, editorExtensions), current.page);
        try { roundtrips.push({ pages: [...layout.element.querySelectorAll('.print-content')].map(p => p.textContent!.replace(/\s/g, '')),
          picturePage: Number(layout.element.querySelector('img')!.closest<HTMLElement>('.print-page')!.dataset.page) - 1,
          images: layout.element.querySelectorAll('img').length,
          scaled: [...layout.element.querySelectorAll<HTMLElement>('[style]')].some(el => !!el.style.zoom) }); }
        finally { layout.dispose(); }
      }
      return { doc, ink, roundtrips, pages: [...print.element.querySelectorAll('.print-content')].map(p => p.textContent!.replace(/\s/g, '')),
      picturePage: Number(print.element.querySelector('img')!.closest<HTMLElement>('.print-page')!.dataset.page) - 1,
      scaled: [...print.element.querySelectorAll<HTMLElement>('[style]')].some(el => !!el.style.zoom) }; }
    finally { print.dispose(); }
  }, { alignment, kind });
  expect(state.scaled).toBe(false);
  for (const roundtrip of state.roundtrips) expect(roundtrip).toEqual({ pages: state.pages, picturePage: state.picturePage, images: 1, scaled: false });
  expect(state.pages.join('')).toContain(kind === 'long-paragraph' ? 'LEFT119' : kind === 'wrapped' ? 'LEFT59' : kind === 'margins' ? 'LEFT19' : 'LEFT27'); expect(state.pages.join('')).toContain(kind === 'four-pages' ? 'RIGHT11' : 'RIGHT5');
  await page.locator('input[type=file]').first().setInputFiles({ name: 'merged.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(state.doc)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => editor.evaluate((root, { zoom, ink }) => {
      const sheet = root.closest<HTMLElement>('.document-stack')!, origin = root.getBoundingClientRect().top, scale = zoom / 100;
      const height = Number(sheet.dataset.pageHeight), top = Number(sheet.dataset.margin), bottom = Number(sheet.dataset.bottomMargin), stride = height + 24;
      const pages: string[] = [], counts: number[] = []; let error = 0, errorAt = ""; const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) {
        const node = walker.currentNode as Text;
        if (node.parentElement?.closest('.ProseMirror-widget')) continue;
        for (let i = 0; i < node.length; i++) {
          if (/\s/.test(node.data[i])) continue;
          const range = document.createRange(); range.setStart(node, i); range.setEnd(node, i + 1);
          const rect = range.getBoundingClientRect(), y = (rect.top - origin) / scale;
          const page = Math.max(0, Math.floor((y + .5) / stride)), local = y - page * stride;
          if (local < top - 1 || local + rect.height / scale > height - bottom + 1) return { pages: [`INK IN MARGIN: ${node.data} at ${y}, height ${rect.height / scale}; ${node.parentElement?.getAttribute('style')}`], picturePage: -1 };
          const difference = local - (ink[page]?.[counts[page] || 0] ?? Infinity);
          if (Math.abs(difference) > error) { error = Math.abs(difference); errorAt = `${zoom}% ${node.data}: ${local} vs ${ink[page]?.[counts[page] || 0]}, page ${page}`; } counts[page] = (counts[page] || 0) + 1;
          pages[page] = (pages[page] || '') + node.data[i];
        }
      }
      const image = root.querySelector('img[data-inline-image]')!, rect = image.getBoundingClientRect(), y = (rect.top - origin) / scale;
      const picturePage = Math.max(0, Math.floor((y + .5) / stride)), local = y - picturePage * stride;
      if (local < top - 1 || local + rect.height / scale > height - bottom + 1) return { pages: [`IMAGE IN MARGIN: ${y}`], picturePage };
      return { pages, picturePage, positionMatches: error < 1 || errorAt + ' rows=' + [...root.querySelectorAll('tr')].map(r => r.getBoundingClientRect().height / scale).join(',') };
    }, { zoom, ink: state.ink })).toEqual({ pages: state.pages, picturePage: state.picturePage, positionMatches: true });
  }
  await editor.locator(':scope > p').first().click({ position: { x: 8, y: 8 } }); await page.keyboard.press('Home');
  await page.keyboard.type('EDIT'); await expect(editor).toContainText('EDIT');
  await page.keyboard.press('ControlOrMeta+z'); await expect(editor).not.toContainText('EDIT');
  const download = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
  const saved = JSON.parse(await fs.readFile((await (await download).path())!, 'utf8'));
  expect(saved.content).toEqual(state.doc.content);
});

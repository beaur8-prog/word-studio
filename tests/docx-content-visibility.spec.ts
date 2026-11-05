import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';

for (const story of ['header', 'footer'] as const) test(`${story} right tab keeps text clear of its floating logo in editor, print and two DOCX saves`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async story => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { defaultPage } = await import('/src/types.ts');
    const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';
    const floating = JSON.stringify({ horizontalPosition: { relative: 'column', offset: 0 }, verticalPosition: { relative: 'paragraph', offset: 0 }, wrap: { type: 1, side: 'bothSides' }, behindDocument: true }).replaceAll('"', '&quot;');
    const html = `<p style="font-size:18pt;font-family:Arial;line-height:1.1;margin:0" data-tab-stops='[{"position":600,"type":"right"}]'><img data-inline-image="true" data-image-floating="${floating}" src="data:image/png;base64,${png}" width="72" height="32"><span data-docx-tab="true">\t</span><strong>READABLE COMPANY TITLE</strong></p>`;
    const original = await importFile(new File(['<p>BODY BEFORE</p><div data-page-break="true"></div><p>BODY AFTER</p>'], 'body.html'));
    let doc = { ...original.document, page: { ...defaultPage, size: 'Letter', margin: 19.05, pageNumbers: false, headerDistance: 8, footerDistance: 8, [`${story}Html`]: html }, content: parseDocumentHtml(original.html!) };
    const input = [...new Uint8Array(await (await exportDocx(doc)).arrayBuffer())];
    const first = await importFile(new File([new Uint8Array(input)], 'logo.docx'));
    doc = { ...first.document, content: parseDocumentHtml(first.html!) };
    const states = [], saves: number[][] = [];
    for (let stage = 0; stage < 3; stage++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try {
        states.push({ stage, pages: layout.pages, stories: [...layout.element.querySelectorAll(story === 'header' ? '.print-header' : '.print-footer .document-furniture')].map(root => {
          const image = root.querySelector('img')!, p = root.querySelector('p')!, range = document.createRange();
          const text = [...p.querySelectorAll('strong,b')].find(e => e.textContent === 'READABLE COMPANY TITLE')!;
          range.selectNodeContents(text);
          const a = image.getBoundingClientRect(), b = range.getBoundingClientRect(), c = p.getBoundingClientRect();
          return { text: root.textContent, imageLeft: a.left - c.left, imageRight: a.right - c.left, textLeft: b.left - c.left, textRight: b.right - c.left, tabWidth: root.querySelector<HTMLElement>('[data-docx-tab]')!.getBoundingClientRect().width, imageWidth: a.width };
        }) });
      } finally { layout.dispose(); }
      if (stage < 2) {
        const bytes = [...new Uint8Array(await (await exportDocx(doc)).arrayBuffer())]; saves.push(bytes);
        const loaded = await importFile(new File([new Uint8Array(bytes)], 'again.docx'));
        doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
      }
    }
    return { input, states, saves };
  }, story);
  await writeFile(test.info().outputPath(`${story}-visibility-states.json`), JSON.stringify(result.states, null, 2));
  for (const [i, bytes] of result.saves.entries()) await writeFile(test.info().outputPath(`textbox-save-${story}-visibility-${i + 1}.docx`), Buffer.from(bytes));
  for (const state of result.states) {
    expect(state.pages).toBe(2);
    for (const row of state.stories) {
      expect(row.text).toContain('READABLE COMPANY TITLE');
      expect(row.imageWidth).toBeCloseTo(72, 1);
      expect(row.textLeft).toBeGreaterThan(row.imageRight + 2);
      expect(row.textRight).toBeCloseTo(600, 0);
      expect(row.tabWidth).toBeGreaterThan(72);
    }
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'logo.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(result.input) });
  const headers = page.locator(story === 'header' ? '.paper-header' : '.paper-footer .document-furniture');
  await expect(headers).toHaveCount(2);
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(async () => headers.first().evaluate(root => {
      const p = root.querySelector('p')!, image = root.querySelector('img')!, tab = root.querySelector<HTMLElement>('[data-docx-tab]')!;
      const text = [...p.querySelectorAll('strong,b')].find(e => e.textContent === 'READABLE COMPANY TITLE')!, range = document.createRange(); range.selectNodeContents(text);
      const a = image.getBoundingClientRect(), b = range.getBoundingClientRect(), box = p.getBoundingClientRect(), paper = root.closest<HTMLElement>('.paper')!, scale = paper.getBoundingClientRect().width / paper.offsetWidth;
      return { noOverlap: b.left > a.right + 1, right: Math.round((b.right - box.left) / scale), tab: tab.getBoundingClientRect().width > 0 };
    })).toEqual({ noOverlap: true, right: 600, tab: true });
  }
});

test('a nearly fitting word beside a picture never escapes onto an unreserved line through two DOCX saves', async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async () => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { defaultPage } = await import('/src/types.ts');
    const phrase = 'NEAR LIMIT FIRST LINE LASTWORD';
    const measure = document.createElement('span'); Object.assign(measure.style, { position: 'fixed', whiteSpace: 'pre', font: '16px Arial' }); measure.textContent = phrase; document.body.append(measure);
    const width = measure.getBoundingClientRect().width; measure.remove();
    // Exercise a genuine subpixel overrun. Permitting another .15px here can
    // move the complete final word onto a line that has no picture exclusion.
    const imageWidth = 624 - width + .1;
    const geometry = JSON.stringify({ horizontalPosition: { relative: 'column', offset: 0 }, verticalPosition: { relative: 'paragraph', offset: 0 }, wrap: { type: 1, side: 'bothSides' }, behindDocument: false }).replaceAll('"', '&quot;');
    const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';
    const text = phrase + ' NEXT SENTENCE has to remain readable beside the picture. '.repeat(6);
    const html = `<p style="font-family:Arial;font-size:12pt;line-height:20px;margin:0"><img data-inline-image="true" data-image-floating="${geometry}" src="data:image/png;base64,${png}" width="${imageWidth}" height="100">${text}</p>`;
    const loaded = await importFile(new File([html], 'wrap.html'));
    let doc = { ...loaded.document, page: { ...defaultPage, size: 'Letter', margin: 25.4, pageNumbers: false }, content: parseDocumentHtml(loaded.html!) };
    const states = [];
    for (let stage = 0; stage < 3; stage++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try {
        const image = layout.element.querySelector('img')!, ir = image.getBoundingClientRect(), glyphs = [], walker = document.createTreeWalker(layout.element.querySelector('.print-content')!, NodeFilter.SHOW_TEXT);
        while (walker.nextNode()) {
          const node = walker.currentNode as Text; if (node.parentElement?.closest('[data-flow-widget]')) continue;
          for (const s of new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(node.data)) {
            if (!s.segment.trim()) continue; const r = document.createRange(); r.setStart(node, s.index); r.setEnd(node, s.index + s.segment.length); const b = r.getBoundingClientRect();
            glyphs.push({ text: s.segment, left: b.left, right: b.right, top: b.top, bottom: b.bottom });
          }
        }
        states.push({ stage, pages: layout.pages, text: [...layout.element.querySelectorAll('.print-content')].map(p => p.textContent).join(''), expected: text,
          overlaps: glyphs.filter(g => Math.min(g.right, ir.right) - Math.max(g.left, ir.left) > .5 && Math.min(g.bottom, ir.bottom) - Math.max(g.top, ir.top) > .5), imageWidth: ir.width });
      } finally { layout.dispose(); }
      if (stage < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
    }
    return states;
  });
  await writeFile(test.info().outputPath('subpixel-wrap-states.json'), JSON.stringify(states, null, 2));
  for (const state of states) { expect(state.text).toBe(state.expected); expect(state.pages).toBe(1); expect(state.overlaps).toEqual([]); }
});

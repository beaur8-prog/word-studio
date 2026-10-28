import { test, expect } from '@playwright/test';

for (const placement of ['body', 'cell']) test(`${placement}: exact lines use their measured strut at a fractional page boundary through zoom, editing and two DOCX saves`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async placement => {
    const { importFile, exportDocx } = await import('/src/document-io.ts'), { defaultPage } = await import('/src/types.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts'), { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const settings = { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10, pageNumbers: false };
    const text = Array.from({ length: 16 }, (_, i) => `Boundary${i} `).join('');
    const paragraph = `<p style="font-family:Arial;font-size:10pt;line-height:17.4px;margin:0">${text}</p>`;
    const html = placement === 'body' ? paragraph : `<table style="width:491px;margin:0"><colgroup><col style="width:100px"><col style="width:391px"></colgroup><tr data-row-split="allow"><td style="padding:1.88px;border:1px solid black;vertical-align:top"><p style="font-family:Arial;font-size:10pt;line-height:17.4px;margin:0">LABEL</p></td><td style="padding:1.88px;border:1px solid black;vertical-align:top">${paragraph}</td></tr></table>`;
    const normalized = generateHTML(parseDocumentHtml(html), editorExtensions);
    const probe = await createPrintLayout(normalized, settings);
    let pitch: number, lead: number, underestimate: number;
    try {
      const p = (placement === 'body' ? probe.element.querySelector('.print-content > p') : probe.element.querySelector('td:last-child p')) as HTMLElement;
      const lines: { top: number; bottom: number }[] = [], walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) {
        const range = document.createRange(); range.selectNodeContents(walker.currentNode);
        for (const rect of range.getClientRects()) {
          if (rect.width < 1 || rect.height < 1) continue;
          const same = lines.find(l => Math.abs(l.top - rect.top) < .5);
          if (same) same.bottom = Math.max(same.bottom, rect.bottom); else lines.push({ top: rect.top, bottom: rect.bottom });
        }
      }
      if (lines.length < 2) throw Error('Fixture requires several visual lines');
      lines.sort((a, b) => a.top - b.top);
      pitch = p.getBoundingClientRect().height / lines.length;
      const oldAdvance = lines[0].bottom - p.getBoundingClientRect().top + Math.max(0, lines[1].top - lines[0].bottom) / 2;
      underestimate = pitch - oldAdvance;
      let insets = 0;
      if (placement === 'cell') {
        const cell = p.closest('td')!, table = p.closest('table')!, css = getComputedStyle(cell);
        insets = p.getBoundingClientRect().top - table.getBoundingClientRect().top + parseFloat(css.paddingBottom) + parseFloat(css.borderBottomWidth);
      }
      // Print rejects this full line strut. An asymmetric glyph box must not
      // make the editor accept a shorter synthetic half-leading budget.
      const overflow = .5 + Math.min(.15, Math.max(.08, underestimate / 2));
      lead = 160 * 96 / 25.4 - pitch - insets + overflow;
    } finally { probe.dispose(); }
    const full = `<p style="font-family:Arial;font-size:10pt;line-height:${lead}px;margin:0">LEAD</p>` + html;
    const loaded = await importFile(new File([full], 'boundary.html'));
    let doc = { ...loaded.document, page: settings, content: parseDocumentHtml(full) }, first = doc;
    const states = [];
    for (let save = 0; save < 3; save++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try { states.push({ pages: layout.pages, text: [...layout.element.querySelectorAll('.print-content')].map(p => p.textContent!.replace(/\s/g, '')) }); }
      finally { layout.dispose(); }
      if (save < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
    }
    return { first, states, pitch, lead, underestimate, text };
  }, placement);
  if (process.env.HWP_LAYOUT_DIAGNOSTICS) console.log(JSON.stringify({ placement, ...result, first: undefined }));
  for (const state of result.states) {
    expect(state.pages).toBe(2); expect(state.text.join('')).toBe(result.states[0].text.join(''));
    expect(state.text[0]).not.toContain('Boundary0'); expect(state.text[1]).toContain('Boundary0');
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'boundary.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.first)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  const body = placement === 'body' ? editor.locator(':scope > p').last() : editor.locator('td:last-child p');
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => body.evaluate((p, zoom) => {
      const sheet = p.closest<HTMLElement>('.document-stack')!, origin = sheet.getBoundingClientRect().top;
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) {
        const node = walker.currentNode as Text; if (!node.data.trim() || node.parentElement?.closest('.ProseMirror-widget')) continue;
        const range = document.createRange(); range.setStart(node, 0); range.setEnd(node, 1);
        return Math.floor(((range.getBoundingClientRect().top - origin) / (zoom / 100) + .5) / (Number(sheet.dataset.pageHeight) + 24));
      }
      return -1;
    }, zoom)).toBe(1);
  }
  await body.fill(result.text + 'EDIT'); await page.keyboard.press('ControlOrMeta+z'); await expect(body).toHaveText(result.text);
});

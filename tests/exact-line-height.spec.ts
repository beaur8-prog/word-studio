import { test, expect } from '@playwright/test';

for (const nested of [false, true]) test(`mixed-size exact lines in ${nested ? 'nested cells' : 'body text'} keep one advance through zoom, editing and two DOCX saves`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async nested => {
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { defaultPage } = await import('/src/types.ts');
    const line = '<p style="font-family:Arial;font-size:13pt;line-height:22.53px;margin:0"><span style="font-size:13pt"><strong>■</strong></span><span style="font-size:12pt"><strong> MIXED TITLE</strong></span><br><span style="font-size:10pt">SECOND LINE</span></p>';
    const cell = 'style="padding:0;border:0"', table = (body: string) => `<table style="margin:0"><tr><td ${cell}>${body}</td></tr></table>`;
    const html = nested ? table(table(line + '<p style="line-height:20px;margin:0">AFTER</p>')) : line + '<p style="line-height:20px;margin:0">AFTER</p>';
    let doc = { format: 'word-studio', version: 1, id: 'exact-lines', title: 'Exact lines', updatedAt: new Date().toISOString(), page: { ...defaultPage, pageNumbers: false }, comments: [], content: parseDocumentHtml(html) };
    const first = doc, states = [];
    for (let save = 0; save < 3; save++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try {
        const p = [...layout.element.querySelectorAll('p')].find(p => p.textContent?.includes('MIXED TITLE'))!, after = p.nextElementSibling!;
        states.push({ pages: layout.pages, text: p.textContent, height: p.getBoundingClientRect().height, advance: after.getBoundingClientRect().top - p.getBoundingClientRect().top });
      } finally { layout.dispose(); }
      if (save < 2) { const again = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...again.document, content: parseDocumentHtml(again.html!) }; }
    }
    return { first, states };
  }, nested);
  for (const state of result.states) {
    expect(state.pages).toBe(1); expect(state.text).toBe('■ MIXED TITLESECOND LINE');
    expect(state.height).toBeCloseTo(45.06, 0); expect(state.advance).toBeCloseTo(45.06, 0);
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'exact.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.first)) });
  const paragraph = page.getByRole('textbox', { name: '문서 편집 영역' }).locator('p').filter({ hasText: 'MIXED TITLE' });
  for (const zoom of [50, 100, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => paragraph.evaluate((p, zoom) => p.getBoundingClientRect().height / (zoom / 100), zoom)).toBeCloseTo(45.06, 0);
  }
  await paragraph.fill('EDITED'); await page.keyboard.press('ControlOrMeta+z'); await expect(paragraph).toHaveText('■ MIXED TITLESECOND LINE');
  await expect.poll(() => paragraph.evaluate(p => p.getBoundingClientRect().height)).toBeCloseTo(45.06, 0);
});

test('auto/minimum leading and atomic pictures retain their own height', async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async () => {
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts'), { createPrintLayout } = await import('/src/print-layout.ts');
    const { defaultPage } = await import('/src/types.ts'), { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const picture = 'data:image/svg+xml,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="30" height="80"><rect width="30" height="80" fill="red"/></svg>');
    const html = '<p data-word-line-rule="auto" style="font-size:12pt;line-height:1;margin:0"><span style="font-size:40pt">AUTO</span></p>' +
      '<p data-word-line-rule="atLeast" style="font-size:12pt;line-height:12px;margin:0"><span style="font-size:40pt">MINIMUM</span></p>' +
      `<p style="font-size:12pt;line-height:20px;margin:0"><span>PICTURE</span><img data-inline-image="true" width="30" height="80" src="${picture}"></p>` +
      '<p style="font-size:12pt;line-height:20px;margin:0"><span style="font-size:40pt">ENLARGED</span></p>';
    const layout = await createPrintLayout(generateHTML(parseDocumentHtml(html), editorExtensions), defaultPage);
    try { return [...layout.element.querySelectorAll('.print-content > p')].map(p => ({ height: p.getBoundingClientRect().height, exact: p.hasAttribute('data-exact-line-height') })); }
    finally { layout.dispose(); }
  });
  expect(states[0].exact).toBe(false); expect(states[1].exact).toBe(false);
  expect(states[0].height).toBeGreaterThan(50); expect(states[1].height).toBeGreaterThan(50); expect(states[2].height).toBeGreaterThanOrEqual(80);
  expect(states[3].height).toBeGreaterThan(30);
});

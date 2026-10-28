import { test, expect } from '@playwright/test';

for (const width of [235, 240]) test(`column table with a ${width}px grid uses its real content width and preserves visible overflow checks`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async width => {
    const { createPrintLayout } = await import('/src/print-layout.ts'), { defaultPage } = await import('/src/types.ts');
    const { importFile, exportDocx } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const html = `<p style="line-height:20px;margin:0">PREFIX</p><div data-column-count="2" data-column-space="20" data-column-balance="false"><p style="line-height:20px;margin:0">FIRST COLUMN</p><div data-column-break="true" class="column-break"></div><p style="line-height:20px;margin:0">BEFORE</p><table style="width:${width}px;margin:0"><colgroup><col style="width:${width}px"></colgroup><tr><td colwidth="${width}" style="border:1px solid black;padding:4px"><p style="line-height:20px;margin:0">BORDER TABLE</p></td></tr></table><p style="line-height:20px;margin:0">AFTER</p></div>`;
    const loaded = await importFile(new File([html], 'columns.html'));
    let doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!), page: { ...defaultPage, size: 'Custom', width: 149.8, height: 180, margin: 10, pageNumbers: false } };
    const first = doc, states = [];
    // The wide control exercises the original grid; DOCX intentionally fits
    // oversized tables to a column. The fitting case must survive both saves.
    const saves = width === 235 ? 3 : 1;
    for (let save = 0; save < saves; save++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try {
        const table = layout.element.querySelector('table')!, cell = table.querySelector('p')!, range = document.createRange(); range.selectNodeContents(cell);
        const column = table.closest('[data-column-count]')!, right = column.getBoundingClientRect().right;
        states.push({ pages: layout.pages, tablePage: Number(table.closest<HTMLElement>('.print-page')!.dataset.page), overhang: table.getBoundingClientRect().right - right,
          textOverhang: Math.max(...[...range.getClientRects()].map(r => r.right - right)), text: [...layout.element.querySelectorAll('.print-content')].map(c => c.textContent).join('') });
      } finally { layout.dispose(); }
      if (save + 1 < saves) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
    }
    return { first, states };
  }, width);
  if (process.env.COLUMN_LAYOUT_DIAGNOSTICS) console.log(JSON.stringify({ width, states: result.states }));
  for (const state of result.states) {
    expect(state.pages).toBe(width === 235 ? 1 : 2); expect(state.tablePage).toBe(width === 235 ? 1 : 2);
    expect(state.textOverhang).toBeLessThanOrEqual(.5); expect(state.text).toBe('PREFIXFIRST COLUMNBEFOREBORDER TABLEAFTER');
    if (width === 235 && state === result.states[0]) { expect(state.overhang).toBeGreaterThan(.5); expect(state.overhang).toBeLessThanOrEqual(1.51); }
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'columns.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.first)) });
  if (width === 235) for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe('1');
  }
});

test('a fitting grid does not authorize inline text outside the column edge', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { columnGroupFits } = await import('/src/column-layout.ts');
    const group = document.createElement('div'); group.className = 'tiptap';
    group.style.cssText = 'position:fixed;left:-10000px;top:0;width:235.3px;padding:0;min-height:0';
    group.innerHTML = '<table style="width:235px;margin:0"><colgroup><col style="width:235px"></colgroup><tr><td style="border:1px solid black;padding:0"><p style="text-align:right;line-height:20px;margin:0"><span style="position:relative;left:5px">VISIBLE EDGE</span></p></td></tr></table>';
    document.body.append(group);
    try {
      const text = group.querySelector('span')!, range = document.createRange(); range.selectNodeContents(text);
      return { fits: columnGroupFits(group), textOverhang: range.getBoundingClientRect().right - group.getBoundingClientRect().right };
    } finally { group.remove(); }
  });
  expect(result.textOverhang).toBeGreaterThan(.5); expect(result.fits).toBe(false);
});

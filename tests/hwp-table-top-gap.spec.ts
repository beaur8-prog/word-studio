import { expect, test } from '@playwright/test';

test('native table clearance accompanies every repeated header through editing and two DOCX saves', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { wordLayoutOptions } = await import('/src/word-trailing-spaces.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const expectedGap = 5.25;
    const body = Array.from({ length: 24 }, (_, index) => `<tr><td><p style="line-height:23px">QA row ${index + 1}</p></td></tr>`).join('');
    const html = `<table data-hwp-table-top-gap="${expectedGap}" style="width:260px;margin-bottom:0"><tbody>` +
      '<tr><th><p>Repeated source header</p></th></tr>' + body + '</tbody></table>';
    const input = await importFile(new File([html], 'Repeat gap.html', { type: 'text/html' }));
    let doc = { ...input.document, content: parseDocumentHtml(input.html!), page: {
      ...defaultPage, size: 'Custom' as const, width: 110, height: 95, margin: 10, pageNumbers: false,
    } };
    const original = doc, states = [];
    let removedHeaderGap: number | null | undefined;
    for (let stage = 0; stage < 3; stage++) {
      const output = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page, undefined, undefined, wordLayoutOptions(doc));
      try {
        const tables = [...output.element.querySelectorAll<HTMLElement>('.print-content > table[data-hwp-table-top-gap]')];
        states.push({
          pages: output.pages,
          tops: tables.map(table => ({
            top: table.getBoundingClientRect().top - table.closest('.print-content')!.getBoundingClientRect().top,
            native: Number(table.dataset.hwpTableTopGap),
          })),
          literalRows: [...output.element.querySelectorAll('td p')].map(p => p.textContent),
        });
      } finally { output.dispose(); }
      if (stage < 2) {
        const blob = await exportDocx(doc);
        const zip = await JSZip.loadAsync(await blob.arrayBuffer());
        const xml = await zip.file('word/document.xml')!.async('string');
        if (!xml.includes('ws:topGap="5.25"')) throw Error('The optional table clearance is missing from DOCX');
        const reopened = await importFile(new File([blob], 'Again.docx'));
        doc = { ...reopened.document, content: parseDocumentHtml(reopened.html!) };
        if (stage === 1) {
          zip.file('word/document.xml', xml.replace(/<w:tblHeader\b[^>]*\/>/g, ''));
          const changed = await importFile(new File([await zip.generateAsync({ type: 'blob' })], 'Repeats removed.docx'));
          removedHeaderGap = parseDocumentHtml(changed.html!).content?.find(n => n.type === 'table')?.attrs?.hwpTableTopGap;
        }
      }
    }
    return { original, expectedGap, states, removedHeaderGap };
  });
  expect(result.removedHeaderGap).toBeNull();
  for (const state of result.states) {
    expect(state.pages).toBeGreaterThan(2);
    expect(state.tops).toHaveLength(state.pages);
    expect(state.literalRows).toEqual(Array.from({ length: 24 }, (_, index) => `QA row ${index + 1}`));
    for (const fragment of state.tops) {
      expect(fragment.native).toBe(result.expectedGap);
      expect(fragment.top).toBeCloseTo(result.expectedGap, 0);
    }
  }
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'Repeat gap.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.original)),
  });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  await expect(editor.locator('table')).toHaveAttribute('data-hwp-table-top-gap', '5.25');
  await expect.poll(() => page.locator('.pagination-headers table[data-hwp-table-top-gap]').count()).toBeGreaterThan(1);
  const offsets = await page.locator('.pagination-headers').evaluate(root => {
    const stack = root.closest<HTMLElement>('.document-stack')!;
    const pitch = Number(stack.dataset.pageHeight) + 24, top = Number(stack.dataset.margin);
    return [...root.children].map(wrapper => (parseFloat((wrapper as HTMLElement).style.top) - top) % pitch);
  });
  for (const offset of offsets) expect(offset).toBeCloseTo(result.expectedGap, 0);
});

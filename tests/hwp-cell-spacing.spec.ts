import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-cell-spacing-fixture.cjs');
for (const nested of [false, true]) for (const format of ['hwp', 'hwpx']) test(`${format} ${nested ? 'nested' : 'ordinary'} table keeps authored paragraph spacing and text through two DOCX saves`, async ({ page }, info) => {
  const sample = (await fixture())[`${format}-${nested}`];
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'spacing.' + format));
      let model = importEditableHwp(await session.editable(), 'spacing.' + format);
      const states = [], files = [];
      for (let save = 0; save < 3; save++) {
        const layout = await createPrintLayout(generateHTML(model.content, editorExtensions), model.page);
        try {
          const paragraphs = [...layout.element.querySelectorAll<HTMLParagraphElement>('.print-content p')];
          const target = paragraphs.find(p => p.textContent === 'SPACED 한글')!;
          const css = getComputedStyle(target), previous = paragraphs[paragraphs.indexOf(target) - 1]!;
          states.push({ before: parseFloat(css.marginTop), after: parseFloat(css.marginBottom), advance: target.getBoundingClientRect().top - previous.getBoundingClientRect().top, pages: layout.pages, paragraphs: paragraphs.map(p => p.textContent) });
        } finally { layout.dispose(); }
        if (save === 2) break;
        const blob = await exportDocx(model), zip = await JSZip.loadAsync(await blob.arrayBuffer());
        files.push({ bytes: [...new Uint8Array(await blob.arrayBuffer())], xml: await zip.file('word/document.xml')!.async('string') });
        const loaded = await importFile(new File([blob], 'again.docx')); model = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
      }
      return { states, files };
    } finally { session.close(); }
  }, { bytes: [...sample.bytes], format });
  await fs.writeFile(info.outputPath('spacing-states.json'), JSON.stringify(result.states, null, 2));
  for (const [i, file] of result.files.entries()) await fs.writeFile(info.outputPath(`textbox-save-cell-spacing-${format}-${nested}-${i + 1}.docx`), Buffer.from(file.bytes));
  for (const state of result.states) {
    expect(state.before).toBeCloseTo(12, 2); expect(state.after).toBeCloseTo(4, 2);
    expect(state.advance).toBeCloseTo(32, 1); expect(state.pages).toBe(1);
    expect(state.paragraphs).toEqual(result.states[0].paragraphs);
  }
  for (const file of result.files) expect(file.xml).toMatch(/<w:spacing\b[^>]*w:after="60"[^>]*w:before="180"/);
});

import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-fixed-spaces-fixture.cjs');
for (const placement of ['body', 'cell', 'nested']) for (const format of ['hwp', 'hwpx']) test(`${format} ${placement} fwSpace keeps native quarter-em width without deleting spaces through two DOCX saves`, async ({ page }, info) => {
  const input = (await fixture())[`${format}-${placement}`];
  await page.goto('/');
  await page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).waitFor();
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const session = new HwpReaderSession(), states = [], files = [];
    try {
      await session.open(new File([new Uint8Array(bytes)], 'fixed.' + format));
      let doc = importEditableHwp(await session.editable(), 'fixed.' + format);
      for (let save = 0; save < 3; save++) {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const p = [...layout.element.querySelectorAll('p')].find(p => p.textContent?.includes('본문'))!;
          const spans = document.createTreeWalker(p, NodeFilter.SHOW_TEXT), widths = [];
          while (spans.nextNode()) { const n = spans.currentNode; for (let i = 0; i < n.textContent!.length; i++) if (n.textContent![i] === '\u2005') { const range = document.createRange(); range.setStart(n, i); range.setEnd(n, i + 1); widths.push({ width: range.getBoundingClientRect().width, font: parseFloat(getComputedStyle(n.parentElement!).fontSize) }); } }
          states.push({ text: p.textContent, pages: layout.pages, widths });
        } finally { layout.dispose(); }
        if (save === 2) break;
        const blob = await exportDocx(doc); files.push([...new Uint8Array(await blob.arrayBuffer())]);
        const opened = await importFile(new File([blob], 'fixed.docx')); doc = { ...opened.document, content: parseDocumentHtml(opened.html!) };
      }
      // An ordinary Word/HTML figure space is not an HWP fwSpace control.
      const literal = await importFile(new File(['<p><span style="font-family:Arial;font-size:15pt">LITERAL END</span></p>'], 'literal.html'));
      const model = { ...literal.document, content: parseDocumentHtml(literal.html!) };
      const blob = await exportDocx(model), back = await importFile(new File([blob], 'literal.docx'));
      const holder = document.createElement('div'); holder.innerHTML = back.html!;
      return { states, files, literal: holder.textContent };
    } finally { session.close(); }
  }, { bytes: [...input.bytes], format });
  await fs.writeFile(info.outputPath('fixed-space-states.json'), JSON.stringify({ native: input.widths, states: result.states }, null, 2));
  for (const [i, bytes] of result.files.entries()) await fs.writeFile(info.outputPath(`textbox-save-fw-space-${format}-${placement}-${i + 1}.docx`), Buffer.from(bytes));
  for (const state of result.states) {
    expect(state.text).toBe(input.text.replaceAll('\u2007', '\u2005')); expect(state.pages).toBe(1);
    expect(state.widths).toHaveLength(4);
    for (const space of state.widths) { expect(space.font).toBeCloseTo(20, 2); expect(space.width).toBeCloseTo(space.font / 4, 1); }
  }
  expect(result.literal).toBe('LITERAL END');
});

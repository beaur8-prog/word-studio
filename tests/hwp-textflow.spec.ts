import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createRequire } from 'node:module';
const createFixture = createRequire(import.meta.url)('./hwp-textflow-fixture.cjs');
let directory: string, files: Record<string, string>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-studio-hwp-textflow-')); files = await createFixture(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

for (const format of ['hwp', 'hwpx']) test(`${format}: zero-leading heading and half-em cut line stay editable without overlap through two DOCX saves`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'textflow.' + format));
      const source = await session.editable();
      let doc = importEditableHwp(source, 'textflow.' + format);
      const states = [];
      for (let save = 0; save < 3; save++) {
        const html = generateHTML(doc.content, editorExtensions), layout = await createPrintLayout(html, doc.page);
        try {
          const paragraphs = [...layout.element.querySelectorAll('td p')];
          const upper = paragraphs.find(p => p.textContent?.includes('UPPER TITLE'))!, lower = paragraphs.find(p => p.textContent?.includes('BIG TITLE'))!;
          const ink = (p: Element) => { const range = document.createRange(); range.selectNodeContents(p); const r = range.getBoundingClientRect(); return { top: r.top, bottom: r.bottom }; };
          const ruler = [...layout.element.querySelectorAll('span')].find(s => s.textContent === '┈'.repeat(80))!;
          states.push({ upper: ink(upper), lower: ink(lower), titleParagraphs: paragraphs.length, rulerWidth: ruler.getBoundingClientRect().width,
            line: parseFloat(getComputedStyle(lower).lineHeight), after: parseFloat(getComputedStyle(lower).marginBottom),
            glyphs: (layout.element.textContent?.match(/┈/g) ?? []).length, text: layout.element.textContent });
        } finally { layout.dispose(); }
        if (save < 2) { const again = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...again.document, content: parseDocumentHtml(again.html!) }; }
      }
      return { states, rawMarkerRemains: source.html.includes('data-hwp-zero-line'), textNodes: JSON.stringify(doc.content).includes('BIG TITLE') };
    } finally { session.close(); }
  }, { bytes: [...await fs.readFile(files[format])], format });
  expect(result.rawMarkerRemains).toBe(false); expect(result.textNodes).toBe(true);
  for (const state of result.states) {
    expect(state.lower.top).toBeGreaterThanOrEqual(state.upper.bottom - .5);
    expect(state.titleParagraphs).toBe(2); expect(state.line + state.after).toBeCloseTo(80, 0);
    expect(state.rulerWidth).toBeCloseTo(80 * 10 * 4 / 3 / 2, 0);
    expect(state.glyphs).toBe(81); expect(state.text).toContain('NORMAL ┈'); expect(state.text).toContain('BIG TITLE     ');
  }
});

test('zero-line normalization keeps unmarked zero leading and never consumes following real text or objects', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { prepareEditableHwp } = await import('/src/hwp/svg-images.ts');
    const { defaultPage } = await import('/src/types.ts');
    const input = '<p style="line-height:0">UNMARKED</p><p> </p>' +
      '<p data-hwp-zero-line="40,20" style="line-height:0">HEADING</p><p>REAL TEXT</p>' +
      '<p data-hwp-zero-line="40,20" style="line-height:0">ANOTHER</p><p><br></p>' +
      '<p data-hwp-zero-line="-5,20" style="line-height:0">INVALID</p><p> </p>';
    const source = await prepareEditableHwp({ html: input, page: defaultPage, warnings: [] });
    const root = document.createElement('div'); root.innerHTML = source.html;
    return { paragraphs: root.querySelectorAll('p').length, text: root.textContent, firstLine: (root.firstElementChild as HTMLElement).style.lineHeight, marker: source.html.includes('data-hwp-zero-line'), breaks: root.querySelectorAll('br').length };
  });
  expect(result.paragraphs).toBe(8); expect(result.text).toContain('REAL TEXT'); expect(result.firstLine).toBe('0');
  expect(result.marker).toBe(false); expect(result.breaks).toBe(1);
});

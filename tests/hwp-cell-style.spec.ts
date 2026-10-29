import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-cell-style-fixture.cjs');
let directory: string, samples: any[];
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-cell-style-')); samples = await fixture(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

for (const format of ['hwp', 'hwpx']) for (const mode of ['bold', 'smaller', 'family'])
test(`${format} ${mode}: compatible mixed runs keep the native terminal height and glyphs through two DOCX saves`, async ({ page }, info) => {
  const sample = samples.find(s => s.format === format && s.mode === mode);
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'style.' + format));
      let doc = importEditableHwp(await session.editable(), 'style.' + format);
      const original = structuredClone(doc), states = [];
      for (let save = 0; save < 3; save++) {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const p = [...layout.element.querySelectorAll<HTMLElement>('td p')].find(p => p.textContent!.includes('ALPHA0'))!, top = p.getBoundingClientRect().top;
          const glyphs: { text: string; top: number; font: number }[] = [], walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
          while (walker.nextNode()) { const n = walker.currentNode as Text; if (n.parentElement?.closest('[data-flow-widget]')) continue;
            for (let i = 0; i < n.length; i++) if (!/\s/.test(n.data[i])) { const range = document.createRange(); range.setStart(n, i); range.setEnd(n, i + 1); glyphs.push({ text: n.data[i], top: range.getBoundingClientRect().top - top, font: parseFloat(getComputedStyle(n.parentElement!).fontSize) }); }
          }
          states.push({ height: p.getBoundingClientRect().height, glyphs, breaks: p.querySelectorAll('br').length, pages: layout.pages });
        } finally { layout.dispose(); }
        if (save < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
      }
      return { original, states };
    } finally { session.close(); }
  }, { bytes: [...await fs.readFile(sample.file)], format });
  await fs.writeFile(info.outputPath('native-height.json'), JSON.stringify({ native: sample.nativeHeight, states: result.states }, null, 2));
  const expected = result.states[0];
  expect(expected.glyphs.map(g => g.text).join('')).toBe(sample.logicalText);
  expect(expected.glyphs.some(g => g.font < 12)).toBe(mode === 'smaller');
  for (const state of result.states) {
    expect(Math.abs(state.height - sample.nativeHeight)).toBeLessThan(.6); expect(state.breaks).toBe(11); expect(state.pages).toBe(1);
    expect(state.glyphs.map(g => g.text)).toEqual(expected.glyphs.map(g => g.text));
    expect(Math.max(...state.glyphs.map((g, i) => Math.abs(g.font - expected.glyphs[i].font)))).toBeLessThan(.05);
    expect(Math.max(...state.glyphs.map((g, i) => Math.abs(g.top - expected.glyphs[i].top)))).toBeLessThan(.6);
  }
  await page.evaluate(async doc => {
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { editorExtensions } = await import('/src/extensions.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    (window as any).cellPdfLayout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
    document.body.classList.add('studio-printing');
  }, result.original);
  try { await page.pdf({ path: info.outputPath('cell.pdf'), printBackground: true, preferCSSPageSize: true }); }
  finally { await page.evaluate(() => { (window as any).cellPdfLayout.dispose(); document.body.classList.remove('studio-printing'); }); }
  await fs.writeFile(info.outputPath('cell-pdf-expected.txt'), sample.logicalText);
  await page.locator('input[type=file]').first().setInputFiles({ name: 'style.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.original)) });
  const paragraphs = page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).locator('td p');
  const paragraph = paragraphs.nth(await paragraphs.evaluateAll(ps => ps.findIndex(p => p.textContent!.includes('ALPHA0'))));
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => paragraph.evaluate((p, { zoom, expected }) => {
      const scale = zoom / 100, top = p.getBoundingClientRect().top, glyphs: { text: string; top: number; font: number }[] = [], walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) { const n = walker.currentNode as Text; if (n.parentElement?.closest('.ProseMirror-widget')) continue;
        for (let i = 0; i < n.length; i++) if (!/\s/.test(n.data[i])) { const r = document.createRange(); r.setStart(n, i); r.setEnd(n, i + 1); glyphs.push({ text: n.data[i], top: (r.getBoundingClientRect().top - top) / scale, font: parseFloat(getComputedStyle(n.parentElement!).fontSize) }); }
      }
      return Math.abs(p.getBoundingClientRect().height / scale - expected.height) < .6 && glyphs.length === expected.glyphs.length && glyphs.every((g, i) => g.text === expected.glyphs[i].text && Math.abs(g.top - expected.glyphs[i].top) < .6 && Math.abs(g.font - expected.glyphs[i].font) < .05);
    }, { zoom, expected })).toBe(true);
  }
  await paragraph.fill('EDIT'); await page.keyboard.press('ControlOrMeta+z'); await expect(paragraph).toContainText('ALPHA0');
  await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(paragraph).toHaveText('EDIT');
});

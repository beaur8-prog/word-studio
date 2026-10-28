import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-hard-breaks-fixture.cjs');
let directory: string, files: Record<string, any>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-hard-breaks-')); files = await fixture(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

for (const format of ['hwp', 'hwpx']) test(`${format}: explicit body/cell/nested breaks retain their slots through two DOCX and two Hancom saves`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp, exportEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts'), { createPrintLayout } = await import('/src/print-layout.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts'), { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const read = async (file: File) => { const session = new HwpReaderSession(); try { await session.open(file); return importEditableHwp(await session.editable(), file.name); } finally { session.close(); } };
    let doc = await read(new File([new Uint8Array(bytes)], 'breaks.' + format)), original = doc;
    const states = [];
    for (let save = 0; save < 5; save++) {
      const html = generateHTML(doc.content, editorExtensions), root = document.createElement('div'); root.innerHTML = html;
      const text = (node: Node): string => node.nodeName === 'BR' ? '\n' : node.nodeType === Node.TEXT_NODE ? node.textContent || '' : [...node.childNodes].map(text).join('');
      const layout = await createPrintLayout(html, doc.page);
      states.push({ paragraphs: [...root.querySelectorAll('p')].map(text).filter(s => /_ALPHA|^TAIL$/.test(s)), breaks: root.querySelectorAll('br').length, images: root.querySelectorAll('img').length, pages: layout.pages }); layout.dispose();
      if (save < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
      else if (save < 4) { const out = await exportEditableHwp(doc, format); doc = await read(new File([out.blob], 'again.' + format)); }
    }
    return { states, original };
  }, { bytes: [...await fs.readFile(files[format])], format });
  const expected = ['BODY & 😀_ALPHA\nBODY_BETA', 'MULTI_ALPHA\n\nMULTI_BETA\n', 'CELL_ALPHA\nCELL_BETA', 'NESTED_ALPHA\nNESTED_BETA', 'PICTURE_ALPHA\n😀&PICTURE_BETA\nPICTURE_END', 'TAIL'];
  for (const state of result.states) { expect(state.paragraphs).toEqual(expected); expect(state.breaks).toBe(8); expect(state.images).toBe(1); expect(state.pages).toBe(result.states[0].pages); }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'breaks.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.original)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' }), paragraph = editor.locator('p').filter({ hasText: 'BODY_BETA' });
  await expect(paragraph.locator('br:not(.ProseMirror-trailingBreak)')).toHaveCount(1);
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    const gap = (target) => target.evaluate((p, zoom) => {
      const first = document.createRange(), last = document.createRange(), walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT), nodes: Node[] = [];
      while (walker.nextNode()) if (walker.currentNode.textContent?.trim()) nodes.push(walker.currentNode);
      first.selectNodeContents(nodes[0]); last.selectNodeContents(nodes.at(-1)!);
      return (last.getBoundingClientRect().top - first.getBoundingClientRect().top) / (zoom / 100);
    }, zoom);
    await expect.poll(() => gap(paragraph)).toBeGreaterThan(10);
    await expect.poll(() => gap(editor.locator('p').filter({ hasText: 'MULTI_BETA' }))).toBeCloseTo(files.multiLineGap, 0);
  }
  await paragraph.click(); await page.keyboard.press('End'); await page.keyboard.type(' EDIT');
  await expect(paragraph).toContainText('EDIT'); await page.keyboard.press('ControlOrMeta+z'); await expect(paragraph).not.toContainText('EDIT');
  await expect(paragraph.locator('br:not(.ProseMirror-trailingBreak)')).toHaveCount(1);
  await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(paragraph).toContainText('EDIT');
});

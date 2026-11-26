import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-script-fonts-fixture.cjs');
let directory: string, files: Record<string, string>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-script-fonts-')); files = await fixture(directory, { latinSize: 125, hangulSpacing: -5, hangulWidth: 95, latinWidth: 110, columns: 2 }); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

for (const format of ['hwp', 'hwpx']) test(`${format}: source script fonts, sizes and widths reach standard DOCX and survive two saves`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const session = new HwpReaderSession(); let doc;
    try { await session.open(new File([new Uint8Array(bytes)], 'scripts.' + format)); doc = importEditableHwp(await session.editable(), 'scripts.' + format); }
    finally { session.close(); }
    const original = doc, states = [];
    const logical = (n: any): string => n.type === 'hardBreak' ? '\n' : (n.text || '') + (n.content || []).map(logical).join('');
    for (let save = 0; save < 3; save++) {
      const blob = await exportDocx(doc), zip = await JSZip.loadAsync(await blob.arrayBuffer());
      const xml = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
      const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      const runs = [...xml.getElementsByTagNameNS(W, 'r')].map(r => ({ text: [...r.getElementsByTagNameNS(W, 't')].map(t => t.textContent).join(''),
        font: r.getElementsByTagNameNS(W, 'rFonts')[0]?.getAttributeNS(W, 'ascii'),
        size: r.getElementsByTagNameNS(W, 'sz')[0]?.getAttributeNS(W, 'val'),
        tracking: r.getElementsByTagNameNS(W, 'spacing')[0]?.getAttributeNS(W, 'val'),
        width: r.getElementsByTagNameNS(W, 'w')[0]?.getAttributeNS(W, 'val') }));
      states.push({ body: logical(doc.content), runs });
      if (save < 2) { const opened = await importFile(new File([blob], 'again.docx')); doc = { ...opened.document, content: parseDocumentHtml(opened.html!) }; }
    }
    return { original, states };
  }, { bytes: [...await fs.readFile(files[format])], format });
  for (const state of result.states) {
    expect(state.body).toBe(result.states[0].body);
    expect(state.body).toContain('한글 Latin 123, Bold words\n다음 Next line.');
    const latin = state.runs.filter(r => /[A-Za-z]/.test(r.text));
    expect(latin.length).toBeGreaterThan(2);
    // The second source paragraph uses a 10pt run. The same 125% shape
    // attribute is relative to each run's base, including that paragraph.
    for (const run of latin) expect(run, run.text).toMatchObject({ font: 'Courier New', size: run.text === 'AFTER' ? '25' : '30', tracking: '0', width: '110' });
    for (const run of state.runs.filter(r => /[가-힣]/.test(r.text))) expect(run, run.text).toMatchObject({ font: 'Arial', size: '24', tracking: '-12', width: '95' });
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'scripts.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.original)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  await expect(editor).toContainText('Next line.');
  await expect.poll(() => editor.evaluate(root => {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) if (walker.currentNode.textContent?.includes('Latin 123,')) {
      const css = getComputedStyle(walker.currentNode.parentElement!);
      return [css.fontFamily, css.fontSize, css.letterSpacing === 'normal' ? 0 : parseFloat(css.letterSpacing)];
    }
    return [];
  })).toEqual(['"Courier New"', '20px', 0]);
  await editor.locator('p').filter({ hasText: 'Latin 123,' }).click(); await page.keyboard.press('End'); await page.keyboard.type(' EDIT');
  await expect(editor).toContainText('EDIT'); await page.keyboard.press('ControlOrMeta+z'); await expect(editor).not.toContainText('EDIT');
  await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(editor).toContainText('EDIT');
});

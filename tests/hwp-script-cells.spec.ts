import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-script-cells-fixture.cjs');
let directory: string, files: Record<string, string>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-script-cells-')); files = await fixture(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

for (const format of ['hwp', 'hwpx']) test(`${format}: script fonts in cells and nested cells survive two DOCX saves`, async ({ page }) => {
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
        font: r.getElementsByTagNameNS(W, 'rFonts')[0]?.getAttributeNS(W, 'ascii') }));
      states.push({ body: logical(doc.content), runs });
      if (save < 2) { const opened = await importFile(new File([blob], 'again.docx')); doc = { ...opened.document, content: parseDocumentHtml(opened.html!) }; }
    }
    // An external editor's ordinary font edit must win over any imported
    // HWP layout metadata, including inside the nested cell.
    const zip = await JSZip.loadAsync(await (await exportDocx(doc)).arrayBuffer());
    const xml = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    let changed = 0;
    for (const face of xml.getElementsByTagNameNS(W, 'rFonts')) if (face.getAttributeNS(W, 'ascii') === 'Courier New') {
      for (const name of ['ascii', 'hAnsi', 'cs', 'eastAsia']) face.setAttributeNS(W, 'w:' + name, 'Georgia');
      changed++;
    }
    zip.file('word/document.xml', new XMLSerializer().serializeToString(xml));
    const opened = await importFile(new File([await zip.generateAsync({ type: 'blob' })], 'external.docx'));
    doc = { ...opened.document, content: parseDocumentHtml(opened.html!) };
    const saved = await JSZip.loadAsync(await (await exportDocx(doc)).arrayBuffer());
    const edited = new DOMParser().parseFromString(await saved.file('word/document.xml')!.async('string'), 'application/xml');
    const externalFonts = [...edited.getElementsByTagNameNS(W, 'r')].filter(r => /[A-Za-z]/.test([...r.getElementsByTagNameNS(W, 't')].map(t => t.textContent).join('')))
      .map(r => r.getElementsByTagNameNS(W, 'rFonts')[0]?.getAttributeNS(W, 'ascii'));
    return { original, states, changed, externalBody: logical(doc.content), externalFonts };
  }, { bytes: [...await fs.readFile(files[format])], format });
  for (const state of result.states) {
    expect(state.body).toBe(result.states[0].body);
    expect(state.body).toContain('한글 Latin 123, Bold words\n다음 Next line.');
    expect(state.body).toContain('😀 After astral body');
    expect(state.body).toContain('셀 😀 Table Latin 456');
    expect(state.body).toContain('안쪽 😀 Nested Latin 789');
    expect(state.body).toContain('표뒤 After nested table');
    const latin = state.runs.filter(r => /[A-Za-z]/.test(r.text));
    expect(latin.length).toBeGreaterThan(2);
    for (const run of latin) expect(run.font, run.text).toBe('Courier New');
    for (const run of state.runs.filter(r => /[가-힣]/.test(r.text))) expect(run.font, run.text).toBe('Arial');
  }
  expect(result.changed).toBeGreaterThan(2);
  expect(result.externalBody).toBe(result.states[0].body);
  for (const font of result.externalFonts) expect(font).toBe('Georgia');
  await page.locator('input[type=file]').first().setInputFiles({ name: 'scripts.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.original)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  await expect(editor).toContainText('Next line.');
  await expect.poll(() => editor.evaluate(root => {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) if (walker.currentNode.textContent?.includes('Latin 123,'))
      return getComputedStyle(walker.currentNode.parentElement!).fontFamily;
    return '';
  })).toBe('"Courier New"');
  await editor.locator('p').filter({ hasText: 'Latin 123,' }).click(); await page.keyboard.press('End'); await page.keyboard.type(' EDIT');
  await expect(editor).toContainText('EDIT'); await page.keyboard.press('ControlOrMeta+z'); await expect(editor).not.toContainText('EDIT');
  await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(editor).toContainText('EDIT');
});

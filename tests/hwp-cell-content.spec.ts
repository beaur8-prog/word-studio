import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-cell-content-fixture.cjs');
let directory: string, files: Record<string, string>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-studio-hwp-cell-browser-')); files = await fixture(directory); });
test.afterAll(async () => { await fs.rm(directory, { recursive: true }); });
for (const format of ['hwp', 'hwpx']) test(`${format}: inline cell controls retain text order, nested shapes and pixels through two DOCX saves`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'cell.' + format));
      let doc = importEditableHwp(await session.editable(), 'cell.' + format);
      const states = [];
      for (let save = 0; save < 3; save++) {
        const html = generateHTML(doc.content, editorExtensions), root = document.createElement('div'); root.innerHTML = html;
        const first = root.querySelector('table td')!;
        const order = first.innerHTML.replace(/<img\b[^>]*alt="([^"]*)"[^>]*>/g, '[$1]').replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim();
        const painted = [];
        for (const img of root.querySelectorAll('img')) {
          await img.decode(); const canvas = document.createElement('canvas'); canvas.width = img.naturalWidth; canvas.height = img.naturalHeight;
          const ctx = canvas.getContext('2d')!; ctx.drawImage(img, 0, 0);
          const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
          let count = 0; for (let i = 3; i < data.length; i += 4) if (data[i]) count++;
          painted.push(count);
        }
        states.push({ order, painted, scripts: [...root.querySelectorAll('img')].map(i => i.alt), fields: root.querySelectorAll('[data-page-number="PAGE"]').length, text: root.textContent, raw: /data-hwp-(svg|object)=/.test(html) });
        if (save < 2) { const reopened = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...reopened.document, content: parseDocumentHtml(reopened.html!) }; }
      }
      return states;
    } finally { session.close(); }
  }, { bytes: [...await fs.readFile(files[format])], format });
  for (const state of result) {
    expect(state.order).toBe('BEFORE [수식: x^2] MIDDLE [수식: y over z] AFTER');
    expect(state.painted).toHaveLength(7); expect(state.painted.every(n => n > 10)).toBe(true);
    for (const script of ['a over b', 'x^2', 'y over z', 'sqrt {q}']) expect(state.scripts).toContain('수식: ' + script);
    expect(state.text).toContain('UNCHANGED 한글'); expect(state.text).not.toContain('변환하지 못한'); expect(state.raw).toBe(false);
    expect(state.fields).toBe(1);
  }
});

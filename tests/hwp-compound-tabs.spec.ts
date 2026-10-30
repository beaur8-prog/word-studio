import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-compound-tabs-fixture.cjs');
let directory: string, files: Record<string, string>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-compound-tabs-')); files = await fixture(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

for (const format of ['hwp', 'hwpx']) for (const saveAs of ['docx', 'hwp', 'hwpx'])
test(`${format} → ${saveAs}: both tab stops in body and nested cells survive two saves`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format, saveAs }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp, exportEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const session = new HwpReaderSession(); let doc;
    try { await session.open(new File([new Uint8Array(bytes)], 'tabs.' + format)); doc = importEditableHwp(await session.editable(), 'tabs.' + format); }
    finally { session.close(); }
    const original = doc, states = [];
    const logical = (n: any): string => n.type === 'documentTab' ? '\t' : (n.text || '') + (n.content || []).map(logical).join('');
    for (let save = 0; save < 3; save++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try {
        states.push({ save, body: logical(doc.content), pages: layout.pages,
          paragraphs: [...layout.element.querySelectorAll<HTMLElement>('.print-content p[data-tab-stops]')].map(p => {
            const tabs = [...p.querySelectorAll('[data-docx-tab]')];
            return { text: p.textContent, stops: JSON.parse(p.dataset.tabStops!), positions: tabs.map((tab, i) => {
              const r = document.createRange(); r.selectNodeContents(p); r.setStartAfter(tab);
              if (tabs[i + 1]) r.setEndBefore(tabs[i + 1]);
              const rect = r.getBoundingClientRect(), left = p.getBoundingClientRect().left;
              return { start: rect.left - left, end: rect.right - left };
            }) };
          }) });
      } finally { layout.dispose(); }
      if (save < 2) {
        if (saveAs === 'docx') {
          const opened = await importFile(new File([await exportDocx(doc)], 'again.docx'));
          doc = { ...opened.document, content: parseDocumentHtml(opened.html!) };
        } else {
          const out = await exportEditableHwp(doc, saveAs as 'hwp' | 'hwpx'), reader = new HwpReaderSession();
          try { await reader.open(new File([out.blob], 'again.' + saveAs)); doc = importEditableHwp(await reader.editable(), 'again.' + saveAs); }
          finally { reader.close(); }
        }
      }
    }
    // Ordinary OOXML tab edits made elsewhere must survive the next save.
    let external: number[][] | undefined;
    if (saveAs === 'docx') {
      const zip = await JSZip.loadAsync(await (await exportDocx(doc)).arrayBuffer());
      const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      const xml = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
      for (const tabs of xml.getElementsByTagNameNS(W, 'tabs')) {
        const stop = [...tabs.children].find(s => s.getAttributeNS(W, 'pos') === '2700');
        stop?.setAttributeNS(W, 'w:pos', '3600');
      }
      zip.file('word/document.xml', new XMLSerializer().serializeToString(xml));
      const opened = await importFile(new File([await zip.generateAsync({ type: 'blob' })], 'external.docx'));
      const edited = { ...opened.document, content: parseDocumentHtml(opened.html!) };
      const saved = await JSZip.loadAsync(await (await exportDocx(edited)).arrayBuffer());
      const next = new DOMParser().parseFromString(await saved.file('word/document.xml')!.async('string'), 'application/xml');
      external = [...next.getElementsByTagNameNS(W, 'tabs')].map(t => [...t.children].map(s => Number(s.getAttributeNS(W, 'pos'))));
    }
    return { original, states, external };
  }, { bytes: [...await fs.readFile(files[format])], format, saveAs });
  for (const state of result.states) {
    expect(state.body).toBe(result.states[0].body);
    expect(state.pages).toBe(1);
    expect(state.paragraphs).toHaveLength(3);
    for (const p of state.paragraphs) {
      expect(p.text).toMatch(/^(BODY|CELL|NEST)\tONE\tTWO$/);
      expect(p.stops.slice(0, 2).map((s: { position: number; type: string; leader?: string }) => ({ ...s, leader: s.leader || '' })))
        .toEqual([{ position: 100, type: 'left', leader: '' }, { position: 180, type: 'right', leader: 'dot' }]);
      expect(p.positions).toHaveLength(2);
      expect(p.positions[0].start).toBeCloseTo(100, 0);
      expect(p.positions[1].end).toBeCloseTo(180, 0);
    }
  }
  if (result.external) {
    expect(result.external).toHaveLength(3);
    for (const stops of result.external) { expect(stops).toContain(3600); expect(stops).not.toContain(2700); }
  }
  if (saveAs === 'docx') {
    await page.locator('input[type=file]').first().setInputFiles({ name: 'tabs.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.original)) });
    const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
    const body = editor.locator('p').filter({ hasText: /^BODY/ });
    for (const zoom of [50, 100, 125, 200, 100]) {
      const slider = page.getByRole('slider', { name: '확대 축소' });
      await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
      await expect.poll(() => body.evaluate((p, zoom) => {
        const tab = p.querySelectorAll('[data-docx-tab]')[1], r = document.createRange();
        r.selectNodeContents(p); r.setStartAfter(tab);
        return Math.round((r.getBoundingClientRect().right - p.getBoundingClientRect().left) / (zoom / 100));
      }, zoom)).toBe(180);
    }
    await body.click(); await page.keyboard.press('End'); await page.keyboard.type(' EDIT');
    await expect(body).toContainText('EDIT'); await page.keyboard.press('ControlOrMeta+z'); await expect(body).not.toContainText('EDIT');
    await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(body).toContainText('EDIT');
  }
});

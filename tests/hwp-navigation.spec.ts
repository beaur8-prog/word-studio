import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-navigation-fixture.cjs');
let directory: string, files: Record<string, string>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-hwp-navigation-')); files = await fixture(directory); });
test.afterAll(async () => { await fs.rm(directory, { recursive: true, force: true }); });
for (const format of ['hwp', 'hwpx'] as const) for (const saveAs of ['docx', 'hwp', 'hwpx'] as const) test(`${format} → ${saveAs}: body/nested bookmarks and tab alignment survive two saves`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format, saveAs }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp, exportEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const s = new HwpReaderSession(), states = [];
    try {
      await s.open(new File([new Uint8Array(bytes)], 'navigation.' + format));
      const imported = await s.editable(); let doc = importEditableHwp(imported, 'navigation.' + format);
      const original = doc;
      for (let save = 0; save < 3; save++) {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const root = layout.element;
          const bookmarks = [...root.querySelectorAll<HTMLElement>('[data-docx-bookmark]')].map(b => { const p = b.closest('p')!, range = document.createRange(); range.selectNodeContents(p); range.setEndBefore(b); return { name: b.dataset.docxBookmark, before: range.toString(), text: p.textContent }; });
          const paragraphs = [...root.querySelectorAll<HTMLElement>('.print-content p')].filter(p => p.querySelector('[data-docx-tab]'));
          states.push({ pages: layout.pages, text: root.querySelector('.print-content')!.textContent, bookmarks,
            tabs: paragraphs.map(p => { const tab = p.querySelector('[data-docx-tab]')!, range = document.createRange(); range.selectNodeContents(p); range.setStartAfter(tab); const box = range.getBoundingClientRect(); return { text: p.textContent, stops: JSON.parse(p.dataset.tabStops || '[]'), start: box.left - p.getBoundingClientRect().left, end: box.right - p.getBoundingClientRect().left, center: (box.left + box.right) / 2 - p.getBoundingClientRect().left }; }) });
        } finally { layout.dispose(); }
        if (save < 2) {
          if (saveAs === 'docx') { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
          else { const exported = await exportEditableHwp(doc, saveAs), reader = new HwpReaderSession(); try { await reader.open(new File([exported.blob], 'again.' + saveAs)); doc = importEditableHwp(await reader.editable(), 'again.' + saveAs); } finally { reader.close(); } }
        }
      }
      return { states, original, warnings: imported.warnings };
    } finally { s.close(); }
  }, { bytes: [...await fs.readFile(files[format])], format, saveAs });
  if (process.env.HWP_LAYOUT_DIAGNOSTICS) console.log(JSON.stringify(result.states));
  for (const state of result.states) {
    expect(state.pages).toBe(1);
    expect(state.bookmarks.map(b => b.name)).toEqual(['본문 & Bookmark', '셀 & 책갈피', '중첩 책갈피']);
    expect(state.bookmarks.map(b => b.before)).toEqual(['😀', 'CELL\t', 'NEST\t']);
    expect(state.tabs).toHaveLength(5);
    expect(state.tabs[0].end).toBeCloseTo(360, 0);
    expect(state.tabs[1].center).toBeCloseTo(360, 0);
    expect(state.tabs[3].start).toBeCloseTo(180, 0);
    expect(state.tabs[4].start).toBeCloseTo(180, 0);
    expect(state.text?.replace(/\s/g, '')).toBe(result.states[0].text?.replace(/\s/g, ''));
    const stop = (t: typeof state.tabs[number]) => ({ ...t.stops[0], position: Math.round(t.stops[0].position * 100) / 100, leader: t.stops[0].leader || '' });
    expect(state.tabs.map(stop)).toEqual(result.states[0].tabs.map(stop));
  }
  expect(result.warnings.join('\n')).not.toMatch(/책갈피.*전달되지|변환하지 못한 개체/);
  await page.locator('input[type=file]').first().setInputFiles({ name: 'navigation.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.original)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  await expect(editor.locator('[data-docx-bookmark]')).toHaveCount(3);
  for (const zoom of [50, 100, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => editor.locator('p').filter({ hasText: 'RIGHT' }).evaluate((p, zoom) => { const tab = p.querySelector('[data-docx-tab]')!, range = document.createRange(); range.selectNodeContents(p); range.setStartAfter(tab); return Math.round((range.getBoundingClientRect().right - p.getBoundingClientRect().left) / (zoom / 100)); }, zoom)).toBe(360);
  }
  await editor.locator('p').filter({ hasText: 'RIGHT' }).click(); await page.keyboard.press('ControlOrMeta+a'); await page.keyboard.type('TEMP'); await page.keyboard.press('ControlOrMeta+z');
  await expect(editor.locator('[data-docx-bookmark]')).toHaveCount(3);
});

for (const separator of ['.', ',']) test(`mixed-size decimal tab aligns ${separator} after two DOCX saves`, async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async separator => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const html = `<p data-tab-stops='[{"position":360,"type":"decimal"}]' style="font-size:10pt">Label<span data-docx-tab="true">\t</span><span style="font-size:24pt">123${separator}</span><span style="font-size:8pt">45</span></p>`;
    const loaded = await importFile(new File([html], 'decimal.html'));
    let doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; const states = [];
    for (let i = 0; i < 3; i++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try {
        const p = layout.element.querySelector('.print-content p')!, walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
        let x = -1;
        while (walker.nextNode()) { const n = walker.currentNode as Text, at = n.data.indexOf(separator); if (at < 0) continue; const r = document.createRange(); r.setStart(n, at); r.setEnd(n, at + 1); x = r.getBoundingClientRect().left - p.getBoundingClientRect().left; break; }
        states.push(x);
      } finally { layout.dispose(); }
      if (i < 2) { const again = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...again.document, content: parseDocumentHtml(again.html!) }; }
    }
    return states;
  }, separator);
  for (const x of states) expect(x).toBeCloseTo(360, 0);
});

for (const format of ['hwp', 'hwpx'] as const) test(`${format}: a repeating table header saves its named destination only once`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async format => {
    const { exportEditableHwp, importEditableHwp } = await import('/src/hwp/editing.ts');
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const html = '<table><tr><th><p><span data-docx-bookmark="HEADER"></span>HEADER</p></th></tr>' + Array.from({ length: 45 }, (_, i) => `<tr><td><p>ROW_${i}_END</p></td></tr>`).join('') + '</table>';
    const doc = { format: 'word-studio' as const, version: 1 as const, id: crypto.randomUUID(), title: 'Repeated bookmark', content: parseDocumentHtml(html), page: defaultPage, comments: [], updatedAt: new Date().toISOString() };
    const output = await exportEditableHwp(doc, format), s = new HwpReaderSession();
    try { const info = await s.open(new File([output.blob], 'repeat.' + format)); const model = importEditableHwp(await s.editable(), 'repeat.' + format), root = document.createElement('div'); root.innerHTML = generateHTML(model.content, editorExtensions); return { pages: info.count, bookmarks: root.querySelectorAll('[data-docx-bookmark]').length, text: root.textContent }; }
    finally { s.close(); }
  }, format);
  expect(result.pages).toBeGreaterThan(1); expect(result.bookmarks).toBe(1);
  for (let i = 0; i < 45; i++) expect(result.text?.match(new RegExp(`ROW_${i}_END`, 'g'))).toHaveLength(1);
});

test('tab leaders do not enlarge fixed text lines or create an extra page', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { defaultPage } = await import('/src/types.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const html = Array.from({ length: 26 }, (_, i) => `<p data-tab-stops='[{"position":300,"type":"right","leader":"${['dot', 'hyphen', 'underscore'][i % 3]}"}]' style="font-size:12pt;line-height:20px;margin:0">Item ${i}<span data-docx-tab="true">\t</span>123</p>`).join('');
    const loaded = await importFile(new File([html], 'leaders.html'));
    let doc = { ...loaded.document, page: { ...defaultPage, size: 'Custom' as const, width: 150, height: 160, margin: 10, pageNumbers: false }, content: parseDocumentHtml(loaded.html!) };
    const original = doc, states = [];
    for (let i = 0; i < 3; i++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try { states.push({ pages: layout.pages, heights: [...layout.element.querySelectorAll('.print-content p')].map(p => p.getBoundingClientRect().height) }); }
      finally { layout.dispose(); }
      if (i < 2) { const a = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...a.document, content: parseDocumentHtml(a.html!) }; }
    }
    return { original, states };
  });
  for (const state of result.states) { expect(state.pages).toBe(1); expect(state.heights).toHaveLength(26); for (const h of state.heights) expect(h).toBeCloseTo(20, 1); }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'leaders.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.original)) });
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe('1');
  }
});

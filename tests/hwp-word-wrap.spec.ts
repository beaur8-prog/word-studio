import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
const createFixture = createRequire(import.meta.url)('./hwp-word-wrap-fixture.cjs');
let directory: string, files: Record<string, any>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-hwp-word-wrap-')); files = await createFixture(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

for (const context of ['body', 'ordinary', 'nested']) for (const unit of [0, 1]) for (const format of ['hwp', 'hwpx']) {
  test(`${context} ${unit ? 'character-level' : 'word-level'} ${format}: wrapping survives editing and two DOCX saves`, async ({ page }) => {
    await page.goto('/');
    const fixture = files[`${context}-${unit}-${format}`];
    const result = await page.evaluate(async ({ bytes, text, unit, format }) => {
      const { HwpReaderSession } = await import('/src/hwp/reader.ts');
      const { importEditableHwp, exportEditableHwp } = await import('/src/hwp/editing.ts');
      const { exportDocx, importFile } = await import('/src/document-io.ts');
      const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
      const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
      const { createPrintLayout } = await import('/src/print-layout.ts');
      await import('/node_modules/@fontsource/noto-sans-kr/400.css');
      await document.fonts.load('16px "Noto Sans KR"', text);
      const read = async (file: File) => {
        const session = new HwpReaderSession();
        try { await session.open(file); return importEditableHwp(await session.editable(), file.name); }
        finally { session.close(); }
      };
      let doc = await read(new File([new Uint8Array(bytes)], 'wrap.' + format));
      const original = doc, states = [];
      for (let save = 0; save < (unit === 0 ? 5 : 3); save++) {
        const html = generateHTML(doc.content, editorExtensions), layout = await createPrintLayout(html, doc.page);
        try {
          const p = [...layout.element.querySelectorAll('p')].find(p => p.textContent === text)!;
          const lines: { y: number; text: string }[] = [], walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
          while (walker.nextNode()) {
            const node = walker.currentNode as Text;
            for (let i = 0; i < node.length; i++) {
              const r = document.createRange(); r.setStart(node, i); r.setEnd(node, i + 1);
              const b = r.getBoundingClientRect();
              let line = lines.find(l => Math.abs(l.y - b.top) < 1);
              if (!line) { line = { y: b.top, text: '' }; lines.push(line); }
              line.text += node.data[i];
            }
          }
          const words = new Set(text.split(/\s+/));
          states.push({ text: p.textContent, wordBreak: getComputedStyle(p).wordBreak,
            splitWords: lines.slice(0, -1).filter(l => !words.has(l.text.trim().split(/\s+/).at(-1)!)).length,
            lines: lines.map(l => l.text), breaks: p.querySelectorAll('br').length });
        } finally { layout.dispose(); }
        if (save < 2) {
          const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx'));
          doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
        } else if (save < 4 && unit === 0) {
          const exported = await exportEditableHwp(doc, format);
          doc = await read(new File([exported.blob], 'again.' + format));
        }
      }
      return { original, states };
    }, { bytes: [...await fs.readFile(fixture.file)], text: fixture.text, unit, format });
    await test.info().attach('wrap-states', { body: Buffer.from(JSON.stringify(result.states, null, 2)), contentType: 'application/json' });
    await fs.writeFile(test.info().outputPath('wrap-states.json'), JSON.stringify(result.states, null, 2));
    for (const state of result.states) {
      expect(state.text).toBe(fixture.text); expect(state.breaks).toBe(0);
      expect(state.wordBreak).toBe(unit === 0 ? 'keep-all' : 'normal');
      if (unit === 0) expect(state.splitWords).toBe(0);
    }
    if (unit === 1) expect(result.states[0].splitWords).toBeGreaterThan(0);
    await page.locator('input[type=file]').first().setInputFiles({ name: 'wrap.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.original)) });
    const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
    const initial = editor.locator('p').filter({ hasText: fixture.text });
    const index = await initial.evaluate(p => [...p.closest('.tiptap')!.querySelectorAll('p')].indexOf(p));
    const paragraph = editor.locator('p').nth(index);
    for (const zoom of [100, 50, 125, 200, 100]) {
      const slider = page.getByRole('slider', { name: '확대 축소' });
      await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
      await expect(paragraph).toHaveCSS('word-break', unit === 0 ? 'keep-all' : 'normal');
    }
    await paragraph.click(); await page.keyboard.press('End'); await page.keyboard.type(' EDIT');
    await expect(paragraph).toContainText('EDIT'); await page.keyboard.press('ControlOrMeta+z');
    await expect(paragraph).toHaveText(fixture.text); await page.keyboard.press('ControlOrMeta+Shift+z');
    await expect(paragraph).toContainText('EDIT');
    await expect(paragraph).toHaveCSS('word-break', unit === 0 ? 'keep-all' : 'normal');
  });
}

test('DOCX word-wrap booleans inherit through styles and a direct edit overrides the inherited setting', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const doc = { format: 'word-studio', version: 1, id: 'wrap-styles', title: 'Wrap styles', page: defaultPage, comments: [],
      updatedAt: new Date().toISOString(), content: parseDocumentHtml('<p style="word-break:keep-all">상속된 어절 설정</p><p style="word-break:keep-all">외부에서 수정한 설정</p>') };
    const original = await exportDocx(doc), originalZip = await JSZip.loadAsync(await original.arrayBuffer());
    const xml = await originalZip.file('word/document.xml')!.async('string');
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const off = [...new DOMParser().parseFromString(xml, 'application/xml').getElementsByTagNameNS(W, 'wordWrap')].map(e => e.getAttributeNS(W, 'val'));
    const variants = [];
    for (const value of ['0', 'false', 'off']) {
      const zip = await JSZip.loadAsync(await original.arrayBuffer()), dom = new DOMParser().parseFromString(xml, 'application/xml');
      const ps = [...dom.getElementsByTagNameNS(W, 'p')];
      for (const [i, p] of ps.entries()) {
        const props = p.getElementsByTagNameNS(W, 'pPr')[0], style = dom.createElementNS(W, 'w:pStyle');
        style.setAttributeNS(W, 'w:val', 'WrapChild'); props.prepend(style);
        const wrap = props.getElementsByTagNameNS(W, 'wordWrap')[0];
        if (i === 0) wrap.remove(); else wrap.setAttributeNS(W, 'w:val', '1');
      }
      zip.file('word/document.xml', new XMLSerializer().serializeToString(dom));
      const styles = await zip.file('word/styles.xml')!.async('string');
      zip.file('word/styles.xml', styles.replace('</w:styles>',
        `<w:style w:type="paragraph" w:styleId="WrapBase"><w:name w:val="Wrap base"/><w:pPr><w:wordWrap w:val="${value}"/></w:pPr></w:style>` +
        '<w:style w:type="paragraph" w:styleId="WrapChild"><w:name w:val="Wrap child"/><w:basedOn w:val="WrapBase"/></w:style></w:styles>'));
      const loaded = await importFile(new File([await zip.generateAsync({ type: 'blob' })], 'external.docx'));
      const root = document.createElement('div'); root.innerHTML = loaded.html!;
      variants.push({ value, styles: [...root.querySelectorAll('p')].map(p => p.style.wordBreak) });
    }
    return { off, variants };
  });
  expect(result.off).toEqual(['0', '0']);
  for (const variant of result.variants) expect(variant.styles).toEqual(['keep-all', 'normal']);
});

for (const format of ['hwp', 'hwpx']) test(`${format}: word wrapping, tabs and bookmarks survive the same paragraph restoration`, async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async format => {
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { defaultPage } = await import('/src/types.ts');
    const { exportEditableHwp, importEditableHwp } = await import('/src/hwp/editing.ts');
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    let doc = { format: 'word-studio', version: 1, id: 'wrap-tabs', title: 'Wrap tabs', page: defaultPage, comments: [], updatedAt: new Date().toISOString(),
      content: parseDocumentHtml('<p style="word-break:keep-all" data-tab-stops=\'[{"position":180,"type":"right","leader":"dot"}]\'>' +
        '<span data-docx-bookmark="wrap-start"></span>가나다라마바<span data-docx-tab="true">\t</span>사아자차카타</p>') };
    const states = [];
    for (let save = 0; save < 3; save++) {
      const root = document.createElement('div'); root.innerHTML = generateHTML(doc.content, editorExtensions);
      const p = root.querySelector('p')!;
      states.push({ wordBreak: p.style.wordBreak, text: p.textContent, stops: JSON.parse(p.dataset.tabStops!),
        bookmarks: [...root.querySelectorAll('[data-docx-bookmark]')].map(e => e.getAttribute('data-docx-bookmark')) });
      if (save < 2) {
        const out = await exportEditableHwp(doc, format), session = new HwpReaderSession();
        try { await session.open(new File([out.blob], 'again.' + format)); doc = importEditableHwp(await session.editable(), 'again.' + format); }
        finally { session.close(); }
      }
    }
    return states;
  }, format);
  for (const state of states) {
    expect(state.wordBreak).toBe('keep-all'); expect(state.text).toBe('가나다라마바\t사아자차카타');
    expect(state.bookmarks).toEqual(['wrap-start']); expect(state.text).not.toContain('WORDSTUDIO-NAV');
    expect(state.stops).toEqual(expect.arrayContaining([{ position: 180, type: 'right', leader: 'dot' }]));
  }
});

test('a word longer than its cell stays visible without adding hard breaks through two DOCX saves', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { defaultPage } = await import('/src/types.ts');
    const text = '가나다라마바사'.repeat(30);
    let doc = { format: 'word-studio', version: 1, id: 'long-word', title: 'Long word', page: defaultPage, comments: [], updatedAt: new Date().toISOString(),
      content: parseDocumentHtml(`<table style="width:140px"><tr><td colwidth="140"><p style="word-break:keep-all;font-size:12pt">${text}</p></td></tr></table>`) };
    const states = [];
    for (let save = 0; save < 3; save++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try {
        const ps = [...layout.element.querySelectorAll('td p')], glyphs = [];
        for (const p of ps) {
          const box = p.closest('td')!.getBoundingClientRect(), walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
          while (walker.nextNode()) {
            const n = walker.currentNode as Text;
            for (let i = 0; i < n.length; i++) { const r = document.createRange(); r.setStart(n, i); r.setEnd(n, i + 1); glyphs.push(r.getBoundingClientRect().right - box.right); }
          }
        }
        states.push({ text: ps.map(p => p.textContent).join(''), overflow: Math.max(...glyphs), breaks: ps.reduce((n, p) => n + p.querySelectorAll('br').length, 0) });
      } finally { layout.dispose(); }
      if (save < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
    }
    return { text, states };
  });
  for (const state of result.states) { expect(state.text).toBe(result.text); expect(state.breaks).toBe(0); expect(state.overflow).toBeLessThanOrEqual(1); }
});

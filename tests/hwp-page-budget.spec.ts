import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-page-budget-fixture.cjs');
let directory: string, files: Record<string, any>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-page-budget-')); files = await fixture(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

for (const columns of [1, 2]) for (const rule of ['percent', 'fixed']) for (const format of ['hwp', 'hwpx']) {
  test(`${columns === 1 ? 'body' : 'columns'} ${rule} ${format}: mixed fonts retain native line advances through editing and saves`, async ({ page }) => {
    await page.goto('/'); const source = files[`${columns}-${rule}-${format}`];
    const result = await page.evaluate(async ({ bytes, format, rule, text }) => {
      const { HwpReaderSession } = await import('/src/hwp/reader.ts');
      const { importEditableHwp, exportEditableHwp } = await import('/src/hwp/editing.ts');
      const { exportDocx, importFile } = await import('/src/document-io.ts');
      const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
      const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
      const { createPrintLayout } = await import('/src/print-layout.ts');
      const read = async (file: File) => { const session = new HwpReaderSession();
        try { await session.open(file); return importEditableHwp(await session.editable(), file.name); } finally { session.close(); } };
      let doc = await read(new File([new Uint8Array(bytes)], 'body.' + format));
      const original = doc, states = [];
      for (let save = 0; save < (rule === 'percent' ? 5 : 3); save++) {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const p = [...layout.element.querySelectorAll('p')].find(p => p.textContent === text)!;
          const after = [...layout.element.querySelectorAll('p')].find(p => p.textContent === 'AFTER')!;
          const fonts: number[] = [], walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
          while (walker.nextNode()) if (walker.currentNode.textContent?.trim()) fonts.push(parseFloat(getComputedStyle(walker.currentNode.parentElement!).fontSize));
          states.push({ save, height: p.getBoundingClientRect().height, advance: after.getBoundingClientRect().top - p.getBoundingClientRect().top,
            fonts, text: p.textContent, body: [...layout.element.querySelectorAll('.print-content')].map(c => c.textContent).join(''),
            hardBreaks: p.querySelectorAll('br').length, policy: p.dataset.hwpBodyLine, applied: p.dataset.hwpBodyLayout });
        } finally { layout.dispose(); }
        if (save < 2) {
          const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx'));
          doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
        } else if (save < 4 && rule === 'percent') {
          const saved = await exportEditableHwp(doc, format); doc = await read(new File([saved.blob], 'again.' + format));
        }
      }
      return { original, states };
    }, { bytes: [...await fs.readFile(source.file)], format, rule, text: source.text });
    await test.info().attach('native-line-comparison', { body: Buffer.from(JSON.stringify({ source, states: result.states }, null, 2)), contentType: 'application/json' });
    await fs.writeFile(test.info().outputPath('states.json'), JSON.stringify({ source, states: result.states }, null, 2));
    for (const state of result.states) {
      expect(state.text).toBe(source.text); expect(state.body).toBe(result.states[0].body); expect(state.hardBreaks).toBe(0);
      expect(state.height).toBeCloseTo(source.nativeAdvance, 0); expect(state.advance).toBeCloseTo(source.nativeAdvance, 0);
      expect(Math.max(...state.fonts)).toBeCloseTo(24, 1); expect(Math.min(...state.fonts)).toBeCloseTo(40 / 3, 1);
      if (rule === 'percent') expect(state.applied).toBe('true'); else expect(state.policy).toBeUndefined();
    }
    await page.locator('input[type=file]').first().setInputFiles({ name: 'body.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.original)) });
    const p = page.getByRole('textbox', { name: '문서 편집 영역' }).locator('p').filter({ hasText: 'small words' });
    for (const zoom of [100, 50, 125, 200, 100]) {
      const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
      await expect.poll(() => p.evaluate((p, zoom) => p.getBoundingClientRect().height / (zoom / 100), zoom)).toBeCloseTo(source.nativeAdvance, 0);
    }
    await p.click(); await page.keyboard.press('End'); await page.keyboard.type(' EDIT'); await expect(p).toContainText('EDIT');
    await page.keyboard.press('ControlOrMeta+z'); await expect(p).toHaveText(source.text);
    await expect.poll(() => p.evaluate(p => p.getBoundingClientRect().height)).toBeCloseTo(source.nativeAdvance, 0);
    await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(p).toContainText('EDIT');
  });
}

for (const edge of ['Page', 'Paper']) for (const format of ['hwp', 'hwpx']) {
  test(`${edge} bottom overlay ${format}: table remains editable without a second flow budget through saves`, async ({ page }) => {
    await page.goto('/'); const source = files[`${edge}-${format}`];
    const result = await page.evaluate(async ({ bytes, format }) => {
      const { HwpReaderSession } = await import('/src/hwp/reader.ts');
      const { importEditableHwp, exportEditableHwp } = await import('/src/hwp/editing.ts');
      const { exportDocx, importFile } = await import('/src/document-io.ts');
      const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
      const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
      const { createPrintLayout } = await import('/src/print-layout.ts');
      const read = async (file: File) => { const session = new HwpReaderSession();
        try { await session.open(file); return importEditableHwp(await session.editable(), file.name); } finally { session.close(); } };
      let doc = await read(new File([new Uint8Array(bytes)], 'bottom.' + format));
      const original = doc, states = [];
      for (let save = 0; save < 5; save++) {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const table = layout.element.querySelector('table')!, paper = table.closest('.print-page')!, frame = table.parentElement!;
          const ps = [...layout.element.querySelectorAll('p')], before = ps.find(p => p.textContent === 'BEFORE')!, after = ps.find(p => p.textContent === 'AFTER')!;
          states.push({ save, pages: layout.pages, active: table.classList.contains('table-overlay-active'), anchor: table.dataset.tableAnchor,
            x: table.getBoundingClientRect().left - paper.getBoundingClientRect().left, y: table.getBoundingClientRect().top - paper.getBoundingClientRect().top,
            height: table.getBoundingClientRect().height, frameHeight: frame.getBoundingClientRect().height,
            bodyGap: after.getBoundingClientRect().top - before.getBoundingClientRect().bottom,
            text: [...layout.element.querySelectorAll('.print-content')].map(c => c.textContent).join('') });
        } finally { layout.dispose(); }
        if (save < 2) {
          const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
        } else if (save < 4) {
          const saved = await exportEditableHwp(doc, format); doc = await read(new File([saved.blob], 'again.' + format));
        }
      }
      return { original, states };
    }, { bytes: [...await fs.readFile(source.file)], format });
    await test.info().attach('overlay-geometry', { body: Buffer.from(JSON.stringify({ source, states: result.states }, null, 2)), contentType: 'application/json' });
    await fs.writeFile(test.info().outputPath('states.json'), JSON.stringify({ source, states: result.states }, null, 2));
    for (const state of result.states) {
      expect(state.pages).toBe(1); expect(state.active).toBe(true); expect(state.frameHeight).toBe(0);
      expect(state.text).toBe(result.states[0].text); expect(state.text).not.toContain('WORDSTUDIO-');
      expect(Math.abs(state.x - source.nativePosition.x)).toBeLessThan(1.5);
      expect(Math.abs(state.y - source.nativePosition.y)).toBeLessThan(1.5);
      expect(Math.abs(state.bodyGap)).toBeLessThan(1);
    }
    await page.locator('input[type=file]').first().setInputFiles({ name: 'bottom.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.original)) });
    const table = page.getByRole('textbox', { name: '문서 편집 영역' }).locator('table').filter({ hasText: 'BOTTOM OVERLAY' });
    for (const zoom of [100, 50, 125, 200, 100]) {
      const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
      await expect(table).toHaveClass(/table-overlay-active/);
    }
    const p = table.locator('p'); await p.click(); await page.keyboard.press('End'); await page.keyboard.type(' EDIT');
    await expect(p).toContainText('EDIT'); await page.keyboard.press('ControlOrMeta+z'); await expect(p).toHaveText('BOTTOM OVERLAY');
    await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(p).toContainText('EDIT');
  });
}

test('external DOCX spacing edits retire the native percentage policy', async ({ page }) => {
  await page.goto('/'); const source = files['1-percent-hwp'];
  const result = await page.evaluate(async bytes => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts'), { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const session = new HwpReaderSession(); let doc;
    try { await session.open(new File([new Uint8Array(bytes)], 'body.hwp')); doc = importEditableHwp(await session.editable(), 'body.hwp'); } finally { session.close(); }
    const zip = await JSZip.loadAsync(await (await exportDocx(doc)).arrayBuffer()), W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const xml = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
    const spacing = xml.getElementsByTagNameNS(W, 'spacing')[0]; spacing.setAttributeNS(W, 'w:line', '720'); spacing.setAttributeNS(W, 'w:lineRule', 'exact');
    zip.file('word/document.xml', new XMLSerializer().serializeToString(xml));
    const loaded = await importFile(new File([await zip.generateAsync({ type: 'blob' })], 'external.docx'));
    const root = document.createElement('div'); root.innerHTML = loaded.html!;
    const p = root.querySelector('p')!; return { policy: p.dataset.hwpBodyLine, lineHeight: p.style.lineHeight, text: p.textContent };
  }, [...await fs.readFile(source.file)]);
  expect(result.policy).toBeUndefined(); expect(result.lineHeight).toBe('36pt'); expect(result.text).toBe(source.text);
});

test('external DOCX floating position edits retire the old bottom anchor', async ({ page }) => {
  await page.goto('/'); const source = files['Page-hwp'];
  const result = await page.evaluate(async bytes => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts'), { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const session = new HwpReaderSession(); let doc;
    try { await session.open(new File([new Uint8Array(bytes)], 'bottom.hwp')); doc = importEditableHwp(await session.editable(), 'bottom.hwp'); } finally { session.close(); }
    const zip = await JSZip.loadAsync(await (await exportDocx(doc)).arrayBuffer()), W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const xml = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
    const position = xml.getElementsByTagNameNS(W, 'tblpPr')[0]; position.setAttributeNS(W, 'w:tblpY', String(Number(position.getAttributeNS(W, 'tblpY')) + 300));
    zip.file('word/document.xml', new XMLSerializer().serializeToString(xml));
    const loaded = await importFile(new File([await zip.generateAsync({ type: 'blob' })], 'external.docx'));
    const root = document.createElement('div'); root.innerHTML = loaded.html!;
    return { anchor: root.querySelector('table')!.dataset.tableAnchor, text: root.textContent };
  }, [...await fs.readFile(source.file)]);
  expect(result.anchor).toBeUndefined(); expect(result.text).toContain('BOTTOM OVERLAY'); expect(result.text).toContain('AFTER');
});

test('a bottom table grown taller than the body returns to splittable flow', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts'), { defaultPage } = await import('/src/types.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const anchor = JSON.stringify({ x: 60, y: 0, bottom: 0, overlay: 'page-bottom' }).replaceAll('"', '&quot;');
    const rows = Array.from({ length: 40 }, (_, i) => `<p style="font-size:10pt;line-height:24px;margin:0">KEEP ROW ${i}</p>`).join('');
    const html = `<p>BEFORE</p><table data-table-anchor="${anchor}"><tr data-row-split="allow"><td colwidth="300" style="padding:0">${rows}</td></tr></table><p>AFTER</p>`;
    const doc = { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 20, pageNumbers: false };
    const layout = await createPrintLayout(generateHTML(parseDocumentHtml(html), editorExtensions), doc);
    try { return { pages: layout.pages, active: layout.element.querySelectorAll('.table-overlay-active').length,
      text: [...layout.element.querySelectorAll('.print-content')].map(c => c.textContent).join('') }; } finally { layout.dispose(); }
  });
  expect(result.pages).toBeGreaterThan(1); expect(result.active).toBe(0);
  for (let i = 0; i < 40; i++) expect(result.text).toContain(`KEEP ROW ${i}`);
  expect(result.text).toContain('BEFORE'); expect(result.text).toContain('AFTER');
});


test('a bottom overlay resolves against the final column geometry', async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async () => {
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts'), { defaultPage } = await import('/src/types.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const anchor = JSON.stringify({ x: 60, y: 12, bottom: 0, overlay: 'page-bottom' }).replaceAll('"', '&quot;');
    const paragraph = '<p style="font-size:10pt;line-height:24px;margin:0">COLUMN FLOW</p>';
    const table = `<table data-table-anchor="${anchor}"><tr><td colwidth="220" style="height:60px;padding:0"><p>OVERLAY CELL</p></td></tr></table>`;
    const settings = { ...defaultPage, size: 'Custom', width: 180, height: 220, margin: 20, pageNumbers: false };
    const states = [];
    for (const before of [5, 15, 20]) {
      const html = `<div data-column-count="2" data-column-space="20">${paragraph.repeat(before)}${table}${paragraph.repeat(3)}<div data-column-break="true"></div>${paragraph.repeat(2)}</div>`;
      const layout = await createPrintLayout(generateHTML(parseDocumentHtml(html), editorExtensions), settings);
      try {
        const t = layout.element.querySelector('table')!, sheet = t.closest('.print-page')!, content = sheet.querySelector('.print-content')!;
        states.push({ active: t.classList.contains('table-overlay-active'), x: t.getBoundingClientRect().left - sheet.getBoundingClientRect().left,
          bottomGap: content.getBoundingClientRect().bottom - t.getBoundingClientRect().bottom,
          text: [...layout.element.querySelectorAll('.print-content')].map(c => c.textContent).join('') });
      } finally { layout.dispose(); }
    }
    return states;
  });
  for (const state of states) {
    expect(state.active).toBe(true); expect(Math.abs(state.x - 60)).toBeLessThan(.1);
    expect(Math.abs(state.bottomGap - 12)).toBeLessThan(.1); expect(state.text).toContain('OVERLAY CELL');
  }
});

test('an explicit editor line spacing change retires percentage layout', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts'), { defaultPage } = await import('/src/types.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const content = parseDocumentHtml('<p style="font-size:18pt;line-height:36px" data-hwp-body-line="1 150 24 36"><span style="font-size:18pt">A</span><span style="font-size:10pt"> smaller text</span></p>');
    content.content![0].attrs!.lineSpacing = '48px';
    const layout = await createPrintLayout(generateHTML(content, editorExtensions), defaultPage);
    try { const p = layout.element.querySelector('p')!; return { policy: p.dataset.hwpBodyLine, active: p.dataset.hwpBodyLayout, height: p.getBoundingClientRect().height, text: p.textContent }; }
    finally { layout.dispose(); }
  });
  expect(result.policy).toBeUndefined(); expect(result.active).toBeUndefined(); expect(result.height).toBe(48); expect(result.text).toBe('A smaller text');
});

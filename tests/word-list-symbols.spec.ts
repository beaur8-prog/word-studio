import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { Document, Packer, Paragraph, Header, Footer, LevelFormat } from 'docx';

const defaults = { document: { run: { font: 'Arial', size: 20 }, paragraph: { spacing: { line: 276, before: 0, after: 0 } } } };

test('a Unicode bullet keeps its source label, font and hanging body position through two DOCX saves', async ({ page }, info) => {
  const input = await Packer.toBuffer(new Document({ styles: { default: defaults }, numbering: { config: [{ reference: 'Unicode', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', style: { run: { font: 'Arial' }, paragraph: { indent: { left: 1080, hanging: 720 } } } }] }] }, sections: [{ children: [new Paragraph({ text: 'FIRST', numbering: { reference: 'Unicode', level: 0 } }), new Paragraph({ text: 'SECOND', numbering: { reference: 'Unicode', level: 0 } })] }] }));
  await page.goto('/');
  const result = await page.evaluate(async bytes => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    let file = new File([new Uint8Array(bytes)], 'unicode.docx'); const stages = [], saves = [];
    for (let i = 0; i < 3; i++) {
      const opened = await importFile(file), model = { ...opened.document, content: parseDocumentHtml(opened.html!) };
      const html = generateHTML(model.content, editorExtensions), layout = await createPrintLayout(html, model.page);
      try { stages.push({ model: model.content, html, items: [...layout.element.querySelectorAll('.print-content li')].map(li => {
        const p = li.querySelector('p')!, text = p.firstChild!, range = document.createRange(); range.setStart(text, 0); range.setEnd(text, 1);
        return { pattern: (li as HTMLElement).dataset.listPattern, label: (li as HTMLElement).dataset.listLabel, font: (li as HTMLElement).dataset.listFont, x: range.getBoundingClientRect().left - li.closest('.print-content')!.getBoundingClientRect().left, bodyFont: getComputedStyle(p).fontFamily };
      }) }); } finally { layout.dispose(); }
      if (i < 2) { const saved = [...new Uint8Array(await (await exportDocx(model)).arrayBuffer())]; saves.push(saved); file = new File([new Uint8Array(saved)], 'unicode.docx'); }
    }
    return { stages, saves };
  }, [...input]);
  for (const stage of result.stages) {
    for (const item of stage.items) { expect(item.pattern).toBe('•'); expect(item.label).toBe('•'); expect(item.font).toBe('Arial'); expect(item.x).toBeCloseTo(72, 1); expect(item.bodyFont).toContain('Arial'); }
    expect(stage.html).not.toContain('--word-list-marker');
    expect(JSON.stringify(stage.model)).not.toContain('data-list-glyph');
  }
  for (let i = 0; i < 2; i++) await fs.writeFile(info.outputPath(`list-save-unicode-${i + 1}.docx`), Buffer.from(result.saves[i]));
});

test('known legacy marker fonts display Unicode but retain original codes and fonts in both DOCX saves and furniture', async ({ page }, info) => {
  const cases = [{ font: 'Symbol', pattern: '\uf0b7', glyph: '•' }, { font: 'Wingdings', pattern: '\uf0fc', glyph: '✔' }, { font: 'Symbol', pattern: 'A', glyph: 'Α' }, { font: 'UnknownPrivateMarker', pattern: '\uf0b7', glyph: null }];
  const config = cases.map((item, i) => ({ reference: `Marker${i}`, levels: [{ level: 0, format: LevelFormat.BULLET, text: item.pattern, style: { run: { font: item.font }, paragraph: { indent: { left: 720, hanging: 360 } } } }] }));
  const input = await Packer.toBuffer(new Document({ styles: { default: defaults }, numbering: { config }, sections: [{ headers: { default: new Header({ children: [new Paragraph({ text: 'HEADER', numbering: { reference: 'Marker0', level: 0 } })] }) }, footers: { default: new Footer({ children: [new Paragraph({ text: 'FOOTER', numbering: { reference: 'Marker1', level: 0 } })] }) }, children: cases.map((_, i) => new Paragraph({ text: `BODY-${i}`, numbering: { reference: `Marker${i}`, level: 0 } })) }] }));
  await page.goto('/');
  const result = await page.evaluate(async bytes => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    let file = new File([new Uint8Array(bytes)], 'symbols.docx'); const stages = [], saves = [];
    for (let i = 0; i < 3; i++) {
      const zip = await JSZip.loadAsync(await file.arrayBuffer()), xml = new DOMParser().parseFromString(await zip.file('word/numbering.xml')!.async('string'), 'text/xml');
      const definitions = [...xml.getElementsByTagNameNS('*', 'lvl')].filter(l => l.getAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'ilvl') === '0').map(l => ({ pattern: l.getElementsByTagNameNS('*', 'lvlText')[0]?.getAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'val'), font: l.getElementsByTagNameNS('*', 'rFonts')[0]?.getAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'ascii') }));
      const opened = await importFile(file), model = { ...opened.document, content: parseDocumentHtml(opened.html!) }, html = generateHTML(model.content, editorExtensions), layout = await createPrintLayout(html, model.page);
      const snapshot = (selector: string) => [...layout.element.querySelectorAll(selector)].map(li => ({ text: li.textContent, pattern: (li as HTMLElement).dataset.listPattern, label: (li as HTMLElement).dataset.listLabel, font: (li as HTMLElement).dataset.listFont, glyph: (li as HTMLElement).dataset.listGlyph ?? null, content: getComputedStyle(li, li.hasAttribute('data-word-list-layout') ? '::before' : '::marker').content, bodyFont: getComputedStyle(li.querySelector('p')!).fontFamily }));
      try { stages.push({ definitions, body: snapshot('.print-content li'), header: snapshot('.print-header li'), footer: snapshot('.print-footer li'), html, model: model.content }); } finally { layout.dispose(); }
      if (i < 2) { const saved = [...new Uint8Array(await (await exportDocx(model)).arrayBuffer())]; saves.push(saved); file = new File([new Uint8Array(saved)], 'symbols.docx'); }
    }
    return { stages, saves };
  }, [...input]);
  for (const stage of result.stages) {
    expect(stage.body).toHaveLength(cases.length);
    for (let i = 0; i < cases.length; i++) {
      const item = stage.body[i], original = cases[i]; expect(item.font).toBe(original.font); expect(item.pattern).toBe(original.pattern); expect(item.label).toBe(original.pattern); expect(item.glyph).toBe(original.glyph); expect(item.bodyFont).toContain('Arial');
      if (original.glyph) expect(item.content).toContain(original.glyph);
      expect(stage.definitions).toContainEqual({ pattern: original.pattern, font: original.font });
    }
    expect(stage.header[0].glyph).toBe('•'); expect(stage.header[0].font).toBe('Symbol'); expect(stage.footer[0].glyph).toBe('✔'); expect(stage.footer[0].font).toBe('Wingdings');
    expect(stage.html).not.toContain('data-list-glyph'); expect(JSON.stringify(stage.model)).not.toContain('Studio Symbols');
  }
  for (let i = 0; i < 2; i++) await fs.writeFile(info.outputPath(`list-save-symbols-${i + 1}.docx`), Buffer.from(result.saves[i]));
});

test('level edits, undo and redo change the marker font with its source level without changing body fonts', async ({ page }) => {
  const input = await Packer.toBuffer(new Document({ styles: { default: defaults }, numbering: { config: [{ reference: 'Fonts', levels: ['Symbol', 'Wingdings'].map((font, level) => ({ level, format: LevelFormat.BULLET, text: level ? '\uf0fc' : '\uf0b7', style: { run: { font }, paragraph: { indent: { left: 720 + level * 360, hanging: 360 } } } })) }] }, sections: [{ children: [new Paragraph({ text: 'FIRST', numbering: { reference: 'Fonts', level: 0 } }), new Paragraph({ text: 'SECOND', numbering: { reference: 'Fonts', level: 0 } }), new Paragraph({ text: 'NESTED', numbering: { reference: 'Fonts', level: 1 } })] }] }));
  await page.goto('/'); await page.locator('input[type=file]').first().setInputFiles({ name: 'levels.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: input });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true }), paragraph = editor.locator('p').filter({ hasText: 'SECOND' }); await expect(paragraph).toBeVisible();
  const snapshot = () => paragraph.evaluate(p => { const li = p.parentElement as HTMLElement; return { glyph: li.dataset.listGlyph, font: li.dataset.listFont, bodyFont: getComputedStyle(p).fontFamily }; });
  await expect.poll(snapshot).toEqual({ glyph: '•', font: 'Symbol', bodyFont: 'Arial' });
  await paragraph.click(); await page.keyboard.press(process.platform === 'darwin' ? 'Meta+ArrowRight' : 'End'); await page.waitForTimeout(550); await page.keyboard.press('Tab');
  await expect.poll(snapshot).toEqual({ glyph: '✔', font: 'Wingdings', bodyFont: 'Arial' });
  await page.keyboard.press('ControlOrMeta+z'); await expect.poll(snapshot).toEqual({ glyph: '•', font: 'Symbol', bodyFont: 'Arial' });
  await page.keyboard.press('ControlOrMeta+Shift+z'); await expect.poll(snapshot).toEqual({ glyph: '✔', font: 'Wingdings', bodyFont: 'Arial' });
  await page.waitForTimeout(550); await page.keyboard.press('Shift+Tab'); await expect.poll(snapshot).toEqual({ glyph: '•', font: 'Symbol', bodyFont: 'Arial' });
});

test('list labels are literal metadata while links, scripts and invalid font CSS remain inert', async ({ page }) => {
  await page.goto('/');
  const actual = await page.evaluate(async () => {
    const { sanitizeImportedHtml } = await import('/src/document-html.ts'), { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts'), { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts'), { defaultPage } = await import('/src/types.ts');
    const html = sanitizeImportedHtml('<ol><li data-list-id="Literal" data-list-level="0" data-list-format="bullet" data-list-pattern="javascript:alert(1)" data-list-font="Arial; background-image:url(https://invalid.local/font)"><p data-word-line-rule="auto" style="font-family:Arial;font-size:10pt;line-height:1.15">BODY<a href="javascript:alert(2)" onclick="alert(3)">LINK</a><script>alert(4)</script></p></li></ol>').html;
    const layout = await createPrintLayout(generateHTML(parseDocumentHtml(html), editorExtensions), defaultPage);
    try { const li = layout.element.querySelector('li')!, a = layout.element.querySelector('a')!; return { label: (li as HTMLElement).dataset.listLabel, href: a?.getAttribute('href') ?? null, event: a?.getAttribute('onclick') ?? null, scripts: layout.element.querySelectorAll('script').length, background: getComputedStyle(li).backgroundImage, markerVariable: (li as HTMLElement).style.getPropertyValue('--word-list-marker-font') }; } finally { layout.dispose(); }
  });
  expect(actual.label).toBe('javascript:alert(1)'); expect(actual.href).toBeNull(); expect(actual.event).toBeNull(); expect(actual.scripts).toBe(0); expect(actual.background).toBe('none'); expect(actual.markerVariable).toBe('');
});


test('adjacent Word numbering ids keep independent starts and continue their own sequence through two saves', async ({ page }, info) => {
  const config = [{ reference: 'A', start: 4 }, { reference: 'B', start: 9 }].map(item => ({ reference: item.reference, levels: [{ level: 0, start: item.start, format: LevelFormat.DECIMAL, text: '%1.', style: { run: { font: 'Arial' }, paragraph: { indent: { left: 720, hanging: 360 } } } }] }));
  const references = ['A', 'A', 'B', 'B', 'A'];
  const input = await Packer.toBuffer(new Document({ styles: { default: defaults }, numbering: { config }, sections: [{ children: references.map((reference, i) => new Paragraph({ text: `ITEM-${i}`, numbering: { reference, level: 0 } })) }] }));
  await page.goto('/');
  const actual = await page.evaluate(async bytes => {
    const { importFile, exportDocx } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts'), { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    let file = new File([new Uint8Array(bytes)], 'starts.docx'); const stages = [], saves = [];
    for (let i = 0; i < 3; i++) {
      const opened = await importFile(file), model = { ...opened.document, content: parseDocumentHtml(opened.html!) }, layout = await createPrintLayout(generateHTML(model.content, editorExtensions), model.page);
      try { stages.push([...layout.element.querySelectorAll('.print-content li')].map(li => ({ text: li.textContent, label: (li as HTMLElement).dataset.listLabel, start: (li as HTMLElement).dataset.listStart }))); } finally { layout.dispose(); }
      if (i < 2) { const saved = [...new Uint8Array(await (await exportDocx(model)).arrayBuffer())]; saves.push(saved); file = new File([new Uint8Array(saved)], 'starts.docx'); }
    }
    return { stages, saves };
  }, [...input]);
  for (const stage of actual.stages) {
    expect(stage.map(item => item.text)).toEqual(references.map((_, i) => `ITEM-${i}`));
    expect(stage.map(item => item.label)).toEqual(['4.', '5.', '9.', '10.', '6.']);
    expect(stage.map(item => item.start)).toEqual(['4', '4', '9', '9', '4']);
  }
  for (let i = 0; i < 2; i++) await fs.writeFile(info.outputPath(`list-save-independent-${i + 1}.docx`), Buffer.from(actual.saves[i]));
});

test('a legacy bullet on a long split item appears once and is suppressed on continuation pages', async ({ page }) => {
  const input = await Packer.toBuffer(new Document({ styles: { default: defaults }, numbering: { config: [{ reference: 'Long', levels: [{ level: 0, format: LevelFormat.BULLET, text: '\uf0b7', style: { run: { font: 'Symbol' }, paragraph: { indent: { left: 720, hanging: 360 } } } }] }] }, sections: [{ children: [new Paragraph({ text: Array.from({ length: 1600 }, (_, i) => `TOKEN${i}`).join(' '), numbering: { reference: 'Long', level: 0 } })] }] }));
  await page.goto('/');
  const actual = await page.evaluate(async bytes => {
    const { importFile } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts'), { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const opened = await importFile(new File([new Uint8Array(bytes)], 'long.docx')), model = { ...opened.document, content: parseDocumentHtml(opened.html!) }, layout = await createPrintLayout(generateHTML(model.content, editorExtensions), model.page);
    try { return { pages: layout.element.querySelectorAll('.print-page').length, text: layout.element.querySelector('.print-pages')?.textContent ?? layout.element.textContent, items: [...layout.element.querySelectorAll('.print-content li')].map(li => ({ glyph: (li as HTMLElement).dataset.listGlyph, continuation: li.hasAttribute('data-print-list-continuation'), marker: getComputedStyle(li, '::marker').content })) }; } finally { layout.dispose(); }
  }, [...input]);
  expect(actual.pages).toBeGreaterThan(2);
  expect(actual.items.filter(item => !item.continuation).map(item => item.marker)).toEqual(['"• "']);
  expect(actual.items.filter(item => item.continuation).every(item => item.marker === '""')).toBe(true);
  expect(actual.items.every(item => item.glyph === '•')).toBe(true);
  for (let i = 0; i < 1600; i++) expect(actual.text?.match(new RegExp(`\\bTOKEN${i}\\b`, 'g'))).toHaveLength(1);
});

test('printing a resolved marker keeps neighboring paragraph coordinates and source reading order', async ({ page }, info) => {
  const expected = 'READ_ORDER_BEGINREAD_ORDER_LISTREAD_ORDER_END';
  const input = await Packer.toBuffer(new Document({ styles: { default: defaults }, numbering: { config: [{ reference: 'Order', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', style: { run: { font: 'Arial' }, paragraph: { indent: { left: 720, hanging: 360 } } } }] }] }, sections: [{ children: [new Paragraph('READ_ORDER_BEGIN'), new Paragraph({ text: 'READ_ORDER_LIST', numbering: { reference: 'Order', level: 0 } }), new Paragraph('READ_ORDER_END')] }] }));
  await page.goto('/');
  await page.evaluate(async bytes => {
    const { importFile } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts'), { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const opened = await importFile(new File([new Uint8Array(bytes)], 'reading-order.docx'));
    (window as any).listOrderLayout = await createPrintLayout(generateHTML(parseDocumentHtml(opened.html!), editorExtensions), opened.document.page);
    document.body.classList.add('studio-printing');
  }, [...input]);
  const bounds = () => page.locator('.print-content p').evaluateAll(ps => ps.map(p => { const content = p.closest('.print-content')!.getBoundingClientRect(), r = document.createRange(); r.selectNodeContents(p); const rect = r.getBoundingClientRect(); return { text: p.textContent, x: rect.x - content.x, y: rect.y - content.y, width: rect.width, height: rect.height }; }));
  try {
    const before = await bounds(); await page.emulateMedia({ media: 'print' }); expect(await bounds()).toEqual(before);
    expect((await bounds()).map(p => p.text).join('')).toBe(expected);
    // The PDF/expected pair is audited by the local PDF text-order check,
    // independently of this DOM/geometry check and the OOXML validator.
    await page.pdf({ path: info.outputPath('list-reading-order.pdf'), printBackground: true, preferCSSPageSize: true });
    await fs.writeFile(info.outputPath('list-reading-order-expected.txt'), expected);
  } finally { await page.evaluate(() => { (window as any).listOrderLayout.dispose(); document.body.classList.remove('studio-printing'); }); }
});

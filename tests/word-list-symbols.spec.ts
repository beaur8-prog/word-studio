import { test, expect } from '@playwright/test';
import type { Editor } from '@tiptap/core';
import fs from 'node:fs/promises';
import { Document, Packer, Paragraph, Header, Footer, Table, TableRow, TableCell, LevelFormat, LevelSuffix, LineRuleType, Tab, TabStopType, TextRun } from 'docx';
import JSZip from 'jszip';

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
  const expected = 'READ_ORDER_BEGINREAD_ORDER_LISTREAD_ORDER_MIDTABLE_ATABLE_BREAD_ORDER_END';
  const input = await Packer.toBuffer(new Document({ styles: { default: defaults }, numbering: { config: [{ reference: 'Order', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', style: { run: { font: 'Arial' }, paragraph: { indent: { left: 720, hanging: 360 } } } }] }] }, sections: [{ children: [new Paragraph('READ_ORDER_BEGIN'), new Paragraph({ text: 'READ_ORDER_LIST', numbering: { reference: 'Order', level: 0 } }), new Paragraph('READ_ORDER_MID'), new Table({ rows: [new TableRow({ children: ['TABLE_A', 'TABLE_B'].map(text => new TableCell({ shading: { fill: 'DCE6F1' }, children: [new Paragraph(text)] })) })] }), new Paragraph('READ_ORDER_END')] }] }));
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

const numberingTabCases = [
  { name: 'INHERITED', reference: 'NumberTab', effective: '[1080]', align: true },
  { name: 'DIRECT', reference: 'NoNumberTab', direct: TabStopType.NUM, effective: '[1080]', align: true },
  { name: 'ABSENT', reference: 'NoNumberTab', effective: null, align: false },
  { name: 'ORDINARY', reference: 'NumberTab', direct: TabStopType.LEFT, effective: null, align: false },
  { name: 'INLINE', reference: 'NumberTab', inline: true, effective: '[1080]', align: false },
  { name: 'CLEARED', reference: 'NumberTab', direct: TabStopType.CLEAR, effective: null, align: false },
];

async function numberingTabFixture() {
  const definitions = ['NumberTab', 'NoNumberTab'].map(reference => ({
    reference,
    levels: [{
      level: 0,
      format: LevelFormat.BULLET,
      suffix: LevelSuffix.TAB,
      text: '\uf0b7',
      style: {
        run: { font: 'Symbol' },
        paragraph: {
          indent: { left: 1080, hanging: 360 },
          ...(reference === 'NumberTab' ? { leftTabStop: 1080 } : {}),
        },
      },
    }],
  }));
  const children = numberingTabCases.map(({ name, reference, direct, inline }) => new Paragraph({
    numbering: { reference, level: 0 },
    ...(direct ? { tabStops: [{ type: direct, position: 1080 }] } : {}),
    ...(inline ? { tabStops: [{ type: TabStopType.LEFT, position: 2160 }] } : {}),
    children: [
      inline ? new TextRun({ children: [name + '-HEAD', new Tab(), 'FIELD'] }) : new TextRun(name + '-HEAD'),
      new TextRun({ text: name + '-TAIL', break: 1 }),
    ],
  }));
  const zip = await JSZip.loadAsync(await Packer.toBuffer(new Document({
    styles: { default: defaults },
    numbering: { config: definitions },
    sections: [{ children }],
  })));
  const part = 'word/numbering.xml';
  const xml = await zip.file(part)!.async('string');
  // The packer's typed level options expose a left tab only. Replace just
  // that generated level's sentinel with the independently authored NUM.
  let levelEdits = 0;
  const numbered = xml.replace(/<w:tab\b[^>]*\/>/g, tag => {
    if (!tag.includes('w:val="left"') || !tag.includes('w:pos="1080"')) return tag;
    levelEdits++;
    return tag.replace('w:val="left"', 'w:val="num"');
  });
  expect(levelEdits).toBe(1);
  zip.file(part, numbered);
  return zip.generateAsync({ type: 'nodebuffer' });
}

function measureNumberTabItems(root: HTMLElement) {
  const bodyLeft = root.getBoundingClientRect().left;
  return Array.from(root.querySelectorAll<HTMLElement>('li[data-list-id]'), li => {
    const paragraph = li.querySelector<HTMLElement>(':scope > p')!;
    const marks = new Map<string, number>();
    const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const text = walker.currentNode as Text;
      for (const match of text.data.matchAll(/[A-Z]+-(HEAD|TAIL)|FIELD/g)) {
        const range = document.createRange();
        range.setStart(text, match.index!);
        range.setEnd(text, match.index! + 1);
        marks.set(match[1] ?? 'FIELD', range.getBoundingClientRect().left - bodyLeft);
      }
    }
    return {
      name: paragraph.textContent!.split('-HEAD')[0],
      text: paragraph.textContent,
      pattern: li.dataset.listPattern,
      font: li.dataset.listFont,
      marker: getComputedStyle(li, '::marker').content,
      before: getComputedStyle(li, '::before').content,
      absoluteMarker: li.hasAttribute('data-word-list-layout'),
      ordinalStops: paragraph.dataset.wordNumberingTabs ?? null,
      bodyTabStops: JSON.parse(paragraph.dataset.tabStops ?? '[]'),
      bodyTabs: paragraph.querySelectorAll('[data-docx-tab]').length,
      first: marks.get('HEAD'),
      continued: marks.get('TAIL'),
      field: marks.get('FIELD'),
    };
  });
}

test('Symbol numbering tabs align continuation text without replacing the native marker or an authored body tab', async ({ page }, info) => {
  let bytes = await numberingTabFixture();
  const stages = [];
  await page.goto('/');
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  const title = page.getByRole('textbox', { name: '문서 제목', exact: true });
  const protect = page.getByRole('dialog', { name: '변경 내용 보관' });
  for (let pass = 0; pass < 3; pass++) {
    const name = 'symbol-number-tab-' + pass;
    await page.locator('input[type=file]').first().setInputFiles({
      name: name + '.docx',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      buffer: bytes,
    });
    await expect.poll(async () => await title.inputValue() === name || await protect.isVisible()).toBe(true);
    if (await protect.isVisible()) await protect.getByRole('button', { name: '복구본 보관 후 계속' }).click();
    await expect(title).toHaveValue(name);
    await expect(editor.locator('li[data-list-glyph]')).toHaveCount(numberingTabCases.length);
    await page.evaluate(() => document.fonts.ready);
    await expect(page.locator('.busy-indicator')).toBeHidden();
    await page.evaluate(() => new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve()))));
    const editing = await editor.evaluate(measureNumberTabItems);

    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const print = page.locator('.print-content');
    await expect(print).toHaveCount(1);
    const printing = await print.evaluate(measureNumberTabItems);
    stages.push({ editing, printing });
    await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();

    if (pass < 2) {
      await page.getByRole('button', { name: '파일', exact: true }).click();
      const pending = page.waitForEvent('download');
      await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
      const download = await pending;
      const saved = info.outputPath(`list-save-numbering-tab-${pass + 1}.docx`);
      await download.saveAs(saved);
      bytes = await fs.readFile(saved);
    }
  }

  await fs.writeFile(info.outputPath('numbering-tab-stages.json'), JSON.stringify(stages, null, 2));
  for (const [pass, stage] of stages.entries()) {
    for (const [mode, rows] of Object.entries(stage)) {
      const absent = rows.find(row => row.name === 'ABSENT')!;
      expect(rows.map(row => row.name)).toEqual(numberingTabCases.map(item => item.name));
      for (const [index, item] of numberingTabCases.entries()) {
        const row = rows[index];
        const at = `save ${pass}, ${mode}, ${item.name}`;
        expect.soft(row.ordinalStops, at).toBe(item.effective);
        expect.soft(row.pattern, at).toBe('\uf0b7');
        expect.soft(row.font, at).toBe('Symbol');
        expect.soft(row.marker, at).toBe('"• "');
        expect.soft(row.before, at).not.toContain('•');
        expect.soft(row.absoluteMarker, at).toBe(false);
        expect.soft(row.bodyTabs, at).toBe(item.inline ? 1 : 0);
        expect.soft(row.bodyTabStops.map((tab: { position: number; type: string }) => [tab.position, tab.type]), at)
          .toEqual(item.name === 'ORDINARY' ? [[72, 'left']] : item.inline ? [[144, 'left']] : []);
        expect.soft(row.text?.replace(/\s+/g, ''), at)
          .toBe(item.name + '-HEAD' + (item.inline ? 'FIELD' : '') + item.name + '-TAIL');
        expect.soft(Math.abs(row.first! - absent.first!), at + ', first line x').toBeLessThan(1);
        expect.soft(Math.abs(row.continued! - (absent.continued! - (item.align ? 25 : 0))), at + ', continuation x').toBeLessThan(1);
        if (item.inline) expect.soft(row.field!, at + ', separate body tab advances following text').toBeGreaterThan(row.first! + 30);
      }
    }
  }

  const readModel = () => editor.evaluate(root => (root as HTMLElement & { editor: Editor }).editor.getJSON());
  const beforeUnlist = await readModel();
  const inherited = editor.locator('p').filter({ hasText: 'INHERITED-HEAD' });
  const originalGeometry = stages.at(-1)!.editing.find(item => item.name === 'INHERITED');
  await page.getByRole('button', { name: '홈', exact: true }).click();
  await inherited.click();
  await page.getByRole('button', { name: '글머리 기호', exact: true }).click();
  await expect.poll(async () => {
    const document = await readModel();
    const paragraph = document.content?.find(node => node.type === 'paragraph' &&
      node.content?.some(child => child.text?.includes('INHERITED-HEAD')));
    return paragraph ? paragraph.attrs?.wordNumberingTabs ?? null : 'missing standalone paragraph';
  }).toBeNull();
  await expect(inherited).not.toHaveAttribute('data-word-numbering-tabs', '[1080]');
  await expect(inherited).not.toHaveAttribute('data-word-native-list-indent', 'true');

  await page.keyboard.press('ControlOrMeta+z');
  await expect.poll(readModel).toEqual(beforeUnlist);
  await expect(inherited).toHaveAttribute('data-word-numbering-tabs', '[1080]');
  await expect.poll(async () => (await editor.evaluate(measureNumberTabItems)).find(item => item.name === 'INHERITED'))
    .toEqual(originalGeometry);
});

test('a real tab late in a split Symbol list item keeps every fragment on its native margin', async ({ page }, info) => {
  const tabLine = 27;
  const lines = Array.from({ length: 38 }, (_, index) => 'ITEM' + String(index).padStart(2, '0'));
  const input = await Packer.toBuffer(new Document({
    styles: { default: defaults },
    numbering: { config: [{
      reference: 'LongNumberTab',
      levels: [{
        level: 0, format: LevelFormat.BULLET, suffix: LevelSuffix.TAB, text: '\uf0b7',
        style: { run: { font: 'Symbol' }, paragraph: { indent: { left: 1080, hanging: 360 } } },
      }],
    }] },
    sections: [{
      properties: { page: { size: { width: 6000, height: 6000 }, margin: { top: 750, right: 750, bottom: 750, left: 750 } } },
      children: [new Paragraph({
        numbering: { reference: 'LongNumberTab', level: 0 },
        tabStops: [{ type: TabStopType.NUM, position: 1080 }],
        spacing: { before: 0, after: 0, line: 300, lineRule: LineRuleType.EXACT },
        children: lines.map((text, index) => new TextRun({
          children: index === tabLine ? [text, new Tab(), 'FIELD'] : [text],
          break: index > 0 ? 1 : 0,
        })),
      })],
    }],
  }));
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'split-numbered-tab.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: input,
  });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  const paragraph = editor.locator('li > p');
  await expect(paragraph).toContainText(lines.at(-1)!);
  await expect(paragraph).toHaveAttribute('data-word-numbering-tabs', '[1080]');
  await expect(paragraph.locator('[data-docx-tab]')).toHaveCount(1);
  await expect(paragraph).toHaveCSS('margin-left', '72px');
  await page.evaluate(() => document.fonts.ready);
  const before = await paragraph.evaluate(p => p.textContent);
  expect(before).toBe(lines.map((line, index) => index === tabLine ? line + '\tFIELD' : line).join(''));

  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  const items = page.locator('.print-content li[data-list-id]');
  await expect.poll(() => items.count()).toBeGreaterThan(2);
  const fragments = await items.evaluateAll(items => items.map(li => {
    const p = li.querySelector(':scope > p')!;
    return {
      text: p.textContent,
      tabs: p.querySelectorAll('[data-docx-tab]').length,
      marker: getComputedStyle(li, '::marker').content,
      continuation: li.hasAttribute('data-print-list-continuation'),
      appliedNumberTabIndent: p.hasAttribute('data-word-native-list-indent'),
      marginLeft: getComputedStyle(p).marginLeft,
    };
  }));
  await fs.writeFile(info.outputPath('split-numbered-tab-fragments.json'), JSON.stringify(fragments, null, 2));

  expect.soft(fragments[0].continuation).toBe(false);
  expect.soft(fragments[0].marker).toBe('"• "');
  expect.soft(fragments.slice(1).every(fragment => fragment.continuation && fragment.marker === '""')).toBe(true);
  expect.soft(fragments[0].tabs).toBe(0);
  expect.soft(fragments.reduce((count, fragment) => count + fragment.tabs, 0)).toBe(1);
  expect.soft(fragments.map(fragment => fragment.text).join('')).toBe(before);
  expect.soft(fragments.map(fragment => fragment.appliedNumberTabIndent)).toEqual(fragments.map(() => false));
  expect(fragments.map(fragment => fragment.marginLeft)).toEqual(fragments.map(() => '72px'));
});

import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';
import JSZip from 'jszip';
import type { PageSettings, StudioDocument } from '../src/types';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, HeightRule, LineRuleType, BorderStyle, SectionType } from 'docx';

const none = { style: BorderStyle.NONE, size: 0, color: 'auto' };
const paragraph = (text: string) => new Paragraph({
  children: [new TextRun({ text, font: 'Arial', size: 18 })],
  spacing: { before: 0, after: 0, line: 300, lineRule: LineRuleType.EXACT },
  widowControl: false,
});
const row = () => new Table({
  width: { size: 4500, type: WidthType.DXA }, columnWidths: [4500],
  borders: { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none },
  rows: [new TableRow({ cantSplit: true, height: { value: 4500, rule: HeightRule.ATLEAST },
    children: [new TableCell({ margins: { top: 0, bottom: 0, left: 0, right: 0 }, children: [paragraph('CELL')] })] })],
});

for (const change of ['margins', 'custom-height'] as const) {
  test('same-kind Word paper with different ' + change + ' keeps editor and print pages through two saves', async ({ page }) => {
    const first = { size: { width: 6000, height: 6000 }, margin: { top: 750, bottom: 750, left: 750, right: 750 } };
    const second = change === 'margins'
      ? { ...first, margin: { top: 300, bottom: 300, left: 300, right: 300 } }
      : { ...first, size: { width: 6000, height: 6900 } };
    const bytes = await Packer.toBuffer(new Document({ sections: [
      { properties: { page: first }, children: [paragraph('FIRST')] },
      { properties: { type: SectionType.NEXT_PAGE, page: second }, children: [paragraph('LEAD'), row(), paragraph('AFTER')] },
    ] }));
    await writeFile(test.info().outputPath('textbox-save-section-source.docx'), bytes);
    await page.goto('/');
    const result = await page.evaluate(async bytes => {
      const { importFile, exportDocx } = await import('/src/document-io.ts');
      const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
      const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
      const { createPrintLayout } = await import('/src/print-layout.ts');
      let input = new File([new Uint8Array(bytes)], 'sections.docx');
      const states = [], saves = [];
      for (let pass = 0; pass < 3; pass++) {
        const loaded = await importFile(input);
        const doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        const printed = [...layout.element.querySelectorAll<HTMLElement>('.print-content p')].filter(p => p.textContent?.trim()).map(p => {
          const paper = p.closest<HTMLElement>('.print-page')!, box = p.getBoundingClientRect(), root = paper.getBoundingClientRect();
          return { text: p.textContent, page: Number(paper.dataset.page), x: box.left - root.left, y: box.top - root.top, height: box.height };
        });
        states.push({ doc, pages: layout.pages, printed, warnings: layout.warnings }); layout.dispose();
        if (pass < 2) { const output = await exportDocx(doc); saves.push([...new Uint8Array(await output.arrayBuffer())]); input = new File([output], 'again.docx'); }
      }
      return { states, saves };
    }, [...bytes]);
    for (const [index, bytes] of result.saves.entries()) await writeFile(test.info().outputPath('textbox-save-section-' + (index + 1) + '.docx'), Buffer.from(bytes));
    for (const state of result.states) {
      expect(state.pages).toBe(2); expect(state.warnings).toEqual([]);
      expect(state.printed.map(p => p.text)).toEqual(['FIRST', 'LEAD', 'CELL', 'AFTER']);
      await page.locator('input[type=file]').first().setInputFiles({ name: 'sections.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(state.doc)) });
      const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
      await expect(editor).toContainText('AFTER');
      await expect(page.locator('.busy-indicator')).toBeHidden();
      for (const zoom of [100, 50, 125, 100]) {
        const slider = page.getByRole('slider', { name: '확대 축소' });
        await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
        await expect.poll(async () => editor.evaluate((root, printed) => {
          const stack = root.closest<HTMLElement>('.document-stack')!, sections = [...root.querySelectorAll<HTMLElement>('.document-section-content')];
          let scale = 1;
          for (let p: Element | null = root; p; p = p.parentElement) {
            const css = getComputedStyle(p); scale *= parseFloat(css.zoom) || 1;
            if (css.transform !== 'none') { const m = new DOMMatrixReadOnly(css.transform); scale *= Math.hypot(m.a, m.b); }
          }
          const paragraphs = [...root.querySelectorAll('p')].filter(p => p.textContent?.trim());
          const actual = paragraphs.map(p => {
            const section = p.closest<HTMLElement>('.document-section-content'), local = section?.querySelector<HTMLElement>('.section-editable') || root;
            const h = Number(section?.dataset.pageHeight ?? stack.dataset.pageHeight), box = p.getBoundingClientRect(), origin = local.getBoundingClientRect();
            const y = (box.top - origin.top) / scale, localPage = Math.max(0, Math.floor((y + .5) / (h + 24)));
            const before = section ? sections.slice(0, sections.indexOf(section)).reduce((sum, s) => sum + Number(s.dataset.pages || 1), 0) : 0;
            return { text: p.textContent, page: before + localPage + 1, x: (box.left - origin.left) / scale, y: y - localPage * (h + 24), height: box.height / scale };
          });
          return { pages: Number(stack.dataset.pages),
            same: actual.length === printed.length && actual.every((p,i) => p.text === printed[i].text && p.page === printed[i].page && Math.abs(p.x - printed[i].x) <= 1 && Math.abs(p.y - printed[i].y) <= 1 && Math.abs(p.height - printed[i].height) <= .6) };
        }, state.printed), { timeout: 8000 }).toEqual({ pages: 2, same: true });
      }
    }
  });
}

test('framed section headers and PAGE/NUMPAGES use document-wide order and totals', async ({ page }) => {
  const { Header, Footer, PageNumber, PageBreak } = await import('docx');
  const header = (label: string) => new Header({ children: [paragraph(label)] });
  const footer = () => new Footer({ children: [new Paragraph({ children: [
    new TextRun('PAGE '), new TextRun({ children: [PageNumber.CURRENT] }),
    new TextRun(' OF '), new TextRun({ children: [PageNumber.TOTAL_PAGES] }),
  ] })] });
  const first = { size: { width: 6000, height: 6000 }, margin: { top: 750, bottom: 750, left: 750, right: 750, header: 100, footer: 100 } };
  const second = { ...first, margin: { top: 300, bottom: 300, left: 300, right: 300, header: 100, footer: 100 } };
  const stories = () => ({ headers: { default: header('ODD'), even: header('EVEN') }, footers: { default: footer(), even: footer() } });
  const bytes = await Packer.toBuffer(new Document({ evenAndOddHeaderAndFooters: true, sections: [
    { ...stories(), properties: { page: first }, children: [paragraph('FIRST')] },
    { ...stories(), properties: { type: SectionType.NEXT_PAGE, page: second }, children: [
      paragraph('SECOND'), new Paragraph({ children: [new PageBreak()] }), paragraph('THIRD'),
    ] },
  ] }));
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({ name: 'section-fields.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: bytes });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  await expect(editor).toContainText('THIRD');
  await expect(page.locator('.busy-indicator')).toBeHidden();
  await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe('3');
  await expect(page.locator('.section-paper .section-header')).toHaveText(['ODD', 'EVEN', 'ODD']);
  await expect(page.locator('.section-paper .section-footer')).toHaveText(['PAGE 1 OF 3', 'PAGE 2 OF 3', 'PAGE 3 OF 3']);
  const original = await editor.evaluate((el: any) => el.editor.getJSON());
  await editor.locator('.section-editable p').filter({ hasText: /^FIRST$/ }).click();
  await page.keyboard.press('End'); await page.keyboard.insertText(' EDIT');
  await page.keyboard.press('ControlOrMeta+z');
  await expect.poll(() => editor.evaluate((el: any) => el.editor.getJSON())).toEqual(original);
  await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(editor).toContainText('FIRST EDIT');
  await page.keyboard.press('ControlOrMeta+z');
  await expect.poll(() => editor.evaluate((el: any) => el.editor.getJSON())).toEqual(original);
  await expect(page.locator('.section-paper .section-footer')).toHaveText(['PAGE 1 OF 3', 'PAGE 2 OF 3', 'PAGE 3 OF 3']);
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  await expect(page.locator('.print-page')).toHaveCount(3);
  await expect(page.locator('.print-header')).toHaveText(['ODD', 'EVEN', 'ODD']);
  await expect(page.locator('.print-footer')).toHaveText(['PAGE 1 OF 3', 'PAGE 2 OF 3', 'PAGE 3 OF 3']);
});

for (const [kind, pages] of [['continuous', 1], ['oddPage', 3], ['evenPage', 2]] as const) {
  test('different-margin ' + kind + ' sections retain their shared or parity pages', async ({ page }) => {
    const type = kind === 'continuous' ? SectionType.CONTINUOUS : kind === 'oddPage' ? SectionType.ODD_PAGE : SectionType.EVEN_PAGE;
    const first = { size: { width: 6000, height: 6000 }, margin: { top: 750, bottom: 750, left: 750, right: 750 } };
    const second = { ...first, margin: { top: 300, bottom: 300, left: 300, right: 300 } };
    const input = await Packer.toBuffer(new Document({ sections: [
      { properties: { page: first }, children: [paragraph('FIRST')] },
      { properties: { type, page: second }, children: [paragraph('SECOND')] },
    ] }));
    await page.goto('/');
    const result = await page.evaluate(async bytes => {
      const { importFile } = await import('/src/document-io.ts');
      const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
      const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
      const { createPrintLayout } = await import('/src/print-layout.ts');
      const opened = await importFile(new File([new Uint8Array(bytes)], 'control.docx'));
      const doc = { ...opened.document, content: parseDocumentHtml(opened.html!) };
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      const result = { doc, pages: layout.pages, text: [...layout.element.querySelectorAll('.print-content p')].map(p => p.textContent).filter(Boolean) };
      layout.dispose(); return result;
    }, [...input]);
    expect(result.pages).toBe(pages); expect(result.text).toEqual(['FIRST', 'SECOND']);
    await page.locator('input[type=file]').first().setInputFiles({ name: 'control.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.doc)) });
    await expect(page.getByRole('textbox', { name: '문서 편집 영역', exact: true })).toContainText('SECOND');
    await expect(page.locator('.busy-indicator')).toBeHidden();
    await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe(String(pages));
  });
}

test('page settings still update the first framed section and undo preserves later sections', async ({ page }) => {
  const first = { size: { width: 6000, height: 6000 }, margin: { top: 750, bottom: 750, left: 750, right: 750 } };
  const second = { ...first, margin: { top: 300, bottom: 300, left: 300, right: 300 } };
  const bytes = await Packer.toBuffer(new Document({ sections: [
    { properties: { page:first }, children:[paragraph('FIRST')] },
    { properties: { type:SectionType.NEXT_PAGE, page:second }, children:[paragraph('SECOND')] },
  ] }));
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({ name:'page-controls.docx', mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer:bytes });
  const editor = page.getByRole('textbox',{name:'문서 편집 영역',exact:true});
  await expect(editor).toContainText('SECOND'); await expect(page.locator('.busy-indicator')).toBeHidden();
  const before = await editor.evaluate((el:any)=>el.editor.getJSON());
  const sections = editor.locator('.document-section-content');
  await expect.poll(async()=>Number(await sections.first().getAttribute('data-margin'))).toBeCloseTo(50,0);
  await expect.poll(async()=>Number(await sections.nth(1).getAttribute('data-margin'))).toBeCloseTo(20,0);
  await page.getByRole('button',{name:'레이아웃',exact:true}).click();
  await page.getByRole('spinbutton',{name:'위 여백',exact:true}).fill('25.4');
  await expect.poll(async()=>Number(await sections.first().getAttribute('data-margin'))).toBeCloseTo(96,4);
  await expect.poll(async()=>Number(await sections.nth(1).getAttribute('data-margin'))).toBeCloseTo(20,0);
  await editor.locator('.section-editable p').first().click();
  await page.keyboard.press('ControlOrMeta+z');
  await expect.poll(()=>editor.evaluate((el:any)=>el.editor.getJSON())).toEqual(before);
  await expect.poll(async()=>Number(await sections.first().getAttribute('data-margin'))).toBeCloseTo(50,0);
  await page.keyboard.press('ControlOrMeta+Shift+z');
  await expect.poll(async()=>Number(await sections.first().getAttribute('data-margin'))).toBeCloseTo(96,4);
  await expect.poll(async()=>Number(await sections.nth(1).getAttribute('data-margin'))).toBeCloseTo(20,0);
});

test('an identical continuous section does not disable the body geometry of an earlier next-page section', async ({ page }) => {
  const first = { size: { width: 6000, height: 6000 }, margin: { top: 750, bottom: 750, left: 750, right: 750 } };
  const body = { ...first, margin: { top: 300, bottom: 300, left: 300, right: 300 } };
  const bytes = await Packer.toBuffer(new Document({ sections: [
    { properties: { page: first }, children: [paragraph('FIRST')] },
    { properties: { type: SectionType.NEXT_PAGE, page: body }, children: [paragraph('LEAD'), row(), paragraph('AFTER')] },
    { properties: { type: SectionType.CONTINUOUS, page: body }, children: [paragraph('CONTINUED')] },
  ] }));
  await page.goto('/');
  const result = await page.evaluate(async bytes => {
    const { importFile, exportDocx, validateDocument } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { withoutEditorReflow } = await import('/src/editor-layout-task.ts');
    let input = new File([new Uint8Array(bytes)], 'mixed-sections.docx'); const states = [], saves = [];
    for (let pass = 0; pass < 3; pass++) {
      const loaded = await importFile(input), doc = validateDocument({ ...loaded.document, content: parseDocumentHtml(loaded.html!) });
      const sections: any[] = []; (function walk(n: any) { if (n.type === 'sectionBreak') sections.push(JSON.parse(n.attrs.page)); n.content?.forEach(walk); })(doc.content);
      const printed = await withoutEditorReflow(async () => {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try { return { pages: layout.pages, paragraphs: [...layout.element.querySelectorAll('.print-content p')].filter(p => p.textContent?.trim()).map(p => {
          const r = p.getBoundingClientRect(), sheet = p.closest<HTMLElement>('.print-page')!, o = sheet.getBoundingClientRect();
          return { text: p.textContent, page: Number(sheet.dataset.page), x: r.x - o.x, y: r.y - o.y, height: r.height };
        }) }; } finally { layout.dispose(); }
      });
      states.push({ doc, sections, printed });
      if (pass < 2) { const blob = await exportDocx(doc); saves.push([...new Uint8Array(await blob.arrayBuffer())]); input = new File([blob], 'again.docx'); }
    }
    const invalid = structuredClone(states[0].doc), frame = invalid.content.content.find((n: any) => n.type === 'sectionContent' && n.content.some((c: any) => c.type === 'sectionBreak'));
    let rejectsChangedGeometry = false;
    if (frame) { const marker = frame.content.find((n: any) => n.type === 'sectionBreak'), page = JSON.parse(marker.attrs.page); page.margins = { ...(await import('/src/types.ts')).pageMargins(page) }; page.margins.left += 1; marker.attrs.page = JSON.stringify(page); try { validateDocument(invalid); } catch { rejectsChangedGeometry = true; } }
    return { states, saves, rejectsChangedGeometry };
  }, [...bytes]);
  for (const [index, bytes] of result.saves.entries()) await writeFile(test.info().outputPath(`textbox-save-section-mixed-${index + 1}.docx`), Buffer.from(bytes));
  const xml = [];
  for (const data of [bytes, ...result.saves.map(data => Buffer.from(data))]) {
    const zip = await JSZip.loadAsync(data), part = await zip.file('word/document.xml')!.async('string');
    xml.push({ sections: [...part.matchAll(/<w:sectPr(?:\s|>)/g)].length, continuous: [...part.matchAll(/<w:type\b[^>]*w:val="continuous"/g)].length });
  }
  expect(xml).toEqual(Array.from({ length: 3 }, () => ({ sections: 3, continuous: 1 })));
  expect(result.rejectsChangedGeometry).toBe(true);
  for (const state of result.states) {
    expect(state.printed.pages).toBe(2); expect(state.sections).toHaveLength(2);
    expect(state.sections.map(p => p.breakType)).toEqual(['nextPage', 'continuous']);
    expect(state.doc.content.content.filter((n: any) => n.type === 'sectionContent')).toHaveLength(2);
    expect(state.printed.paragraphs.map(p => p.text)).toEqual(['FIRST', 'LEAD', 'CELL', 'AFTER', 'CONTINUED']);
    await page.locator('input[type=file]').first().setInputFiles({ name: 'mixed-sections.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(state.doc)) });
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true }); await expect(editor).toContainText('CONTINUED');
    for (const zoom of [100, 50, 125, 100]) {
      const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
      await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe('2');
      await expect.poll(async () => editor.evaluate(async (root, printed) => {
        const { documentScale } = await import('/src/document-scale.ts'), scale = documentScale(root), all = [...root.querySelectorAll<HTMLElement>('.document-section-content')];
        const actual = [...root.querySelectorAll('p')].filter(p => p.textContent?.trim() && !p.closest('.document-furniture')).map(p => {
          const frame = p.closest<HTMLElement>('.document-section-content')!, local = frame.querySelector('.section-editable')!, h = Number(frame.dataset.pageHeight), r = p.getBoundingClientRect(), o = local.getBoundingClientRect();
          const y = (r.y - o.y) / scale, index = Math.floor((y + .5) / (h + 24)), before = all.slice(0, all.indexOf(frame)).reduce((sum, s) => sum + Number(s.dataset.pages), 0);
          return { text: p.textContent, page: before + index + 1, x: (r.x - o.x) / scale, y: y - index * (h + 24), height: r.height / scale };
        });
        return actual.length === printed.length && actual.every((p, i) => p.page === printed[i].page && p.text === printed[i].text && ['x','y','height'].every(k => Math.abs((p as any)[k] - (printed[i] as any)[k]) < .75));
      }, state.printed.paragraphs)).toBe(true);
    }
    const before = await editor.evaluate((root: any) => root.editor.getJSON()); await editor.locator('p').filter({ hasText: /^CONTINUED$/ }).fill('CONTINUED EDIT');
    await page.keyboard.press('ControlOrMeta+z'); await expect.poll(() => editor.evaluate((root: any) => root.editor.getJSON())).toEqual(before);
    await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(editor).toContainText('CONTINUED EDIT'); await page.keyboard.press('ControlOrMeta+z');
    await expect.poll(() => editor.evaluate((root: any) => root.editor.getJSON())).toEqual(before);
  }
});

test('page controls update identical continuous settings in the first shared frame as one undo event', async ({ page }) => {
  const first = { size: { width: 6000, height: 6000 }, margin: { top: 750, bottom: 750, left: 750, right: 750 } };
  const second = { ...first, margin: { top: 300, bottom: 300, left: 300, right: 300 } };
  const input = await Packer.toBuffer(new Document({ sections: [
    { properties: { page: first }, children: [paragraph('FIRST')] },
    { properties: { type: SectionType.CONTINUOUS, page: first }, children: [paragraph('FIRST CONTINUED')] },
    { properties: { type: SectionType.NEXT_PAGE, page: second }, children: [paragraph('SECOND')] },
  ] }));
  await page.goto('/'); await page.locator('input[type=file]').first().setInputFiles({ name: 'shared-page-controls.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: input });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true }); await expect(editor).toContainText('SECOND');
  const original = await editor.evaluate((root: any) => root.editor.getJSON());
  await page.getByRole('button', { name: '레이아웃', exact: true }).click(); await page.getByRole('spinbutton', { name: '위 여백', exact: true }).fill('25.4');
  await expect.poll(() => editor.locator('.document-section-content').first().getAttribute('data-margin')).toBe('96');
  const result = await editor.evaluate(async (root: any) => {
    const ed = root.editor, { documentMetadata } = await import('/src/document-history.ts'), { exportDocx, validateDocument } = await import('/src/document-io.ts');
    const metadata = documentMetadata.getState(ed.state)!;
    const doc = validateDocument({ format: 'word-studio', version: 1, title: 'Shared controls', updatedAt: new Date().toISOString(), comments: metadata.comments, page: metadata.page, content: ed.getJSON() });
    const blob = await exportDocx(doc); return [...new Uint8Array(await blob.arrayBuffer())];
  });
  await writeFile(test.info().outputPath('textbox-save-shared-section-controls.docx'), Buffer.from(result));
  const zip = await JSZip.loadAsync(Buffer.from(result)), xml = await zip.file('word/document.xml')!.async('string');
  const topMargins = [...xml.matchAll(/<w:pgMar\b[^>]*w:top="(\d+)"/g)].map(m => Number(m[1]));
  expect(topMargins).toEqual([1440, 1440, 300]);
  await editor.locator('p').filter({ hasText: /^FIRST CONTINUED$/ }).click(); await page.keyboard.press('ControlOrMeta+z');
  await expect.poll(() => editor.evaluate((root: any) => root.editor.getJSON())).toEqual(original);
  await page.keyboard.press('ControlOrMeta+Shift+z');
  await expect.poll(() => editor.locator('.document-section-content').first().getAttribute('data-margin')).toBe('96');
  await page.keyboard.press('ControlOrMeta+z'); await expect.poll(() => editor.evaluate((root: any) => root.editor.getJSON())).toEqual(original);
});

test('precise custom print paper survives two saves and yields to an external paper-size edit', async ({ page }, testInfo) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { default: Zip } = await import('/node_modules/.vite/deps/jszip.js');
    const ns = 'https://word-studio.app/ooxml/paper-size';
    const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const authored = { ...defaultPage, size: 'Custom', width: 296.9965555555556, height: 419.989, orientation: 'portrait' } as const;
    const landscape = { ...authored, width: 249.915, height: 350.015, orientation: 'landscape' } as const;
    let document: StudioDocument = { format: 'word-studio', version: 1, title: 'Custom stationery', page: authored,
      comments: [], updatedAt: '2026-11-20T13:00:00Z',
      content: parseDocumentHtml('<p>Portrait copy</p><div data-section-page="' +
        JSON.stringify(landscape).replace(/"/g, '&quot;') + '"></div><p>Landscape copy</p>') };
    const folder = [];
    const measurements = [];
    const settings = (doc: StudioDocument) => {
      const sections = (doc.content.content ?? []).filter(node => node.type === 'sectionBreak')
        .map(node => JSON.parse(node.attrs!.page) as PageSettings);
      return [doc.page, ...sections].map(({ width, height, orientation }) => [width, height, orientation]);
    };
    const read = async (file: Blob) => {
      const loaded = await importFile(new File([file], 'custom.docx'));
      return { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
    };
    const parse = async (zip: typeof Zip.prototype) =>
      new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
    for (let pass = 0; pass < 3; pass++) {
      const print = await createPrintLayout(generateHTML(document.content, editorExtensions), document.page);
      try {
        measurements.push({ paper: settings(document), pages: print.pages, rules: print.element.querySelector('style')?.textContent, surface: [...print.element.querySelectorAll<HTMLElement>('.print-page')]
          .map(p => [p.style.width, p.style.height, p.querySelector('.print-content')?.textContent]) });
      } finally { print.dispose(); }
      if (pass >= 2) continue;
      const file = await exportDocx(document);
      const zip = await Zip.loadAsync(await file.arrayBuffer());
      const xml = await parse(zip);
      const layout = [...xml.getElementsByTagNameNS(w,'pgSz')].map(el => ({
        standard: [el.getAttributeNS(w, 'w'), el.getAttributeNS(w, 'h')],
        exact: el.getAttributeNS(ns, 'sizeMm'),
      }));
      folder.push({ layout, output: [...new Uint8Array(await file.arrayBuffer())] });
      document = await read(file);
    }
    const unchanged = new Uint8Array(folder[1].output);
    const altered = await Zip.loadAsync(unchanged);
    const xml = await parse(altered);
    xml.getElementsByTagNameNS(w, 'pgSz')[0].setAttributeNS(w, 'w:h', '24000');
    altered.file('word/document.xml', new XMLSerializer().serializeToString(xml));
    const edited = await read(await altered.generateAsync({type:'blob'}));
    const malformed = await Zip.loadAsync(unchanged);
    const bad = await parse(malformed);
    bad.getElementsByTagNameNS(w, 'pgSz')[0].setAttributeNS(ns, 'wspg:sizeMm', 'Infinity 419.989');
    malformed.file('word/document.xml', new XMLSerializer().serializeToString(bad));
    const invalid = await read(await malformed.generateAsync({type:'blob'}));
    return { measurements, folder, edited: settings(edited), invalid: settings(invalid) };
  });
  const expected = [[296.9965555555556, 419.989, 'portrait'], [249.915, 350.015, 'landscape']] as const;
  const twip = (mm: number) => Math.round(mm * 1440 / 25.4) * 25.4 / 1440;
  const standard = expected.map(([width, height, orientation]) => [twip(width), twip(height), orientation]);
  for (const [index, stage] of result.measurements.entries()) {
    expect(stage.paper).toEqual(index === 0 ? expected : standard);
    expect(stage.pages).toBe(2);
    expect(stage.rules).toContain('size:296.9965555555556mm 419.989mm');
    expect(stage.rules).toContain('size:350.015mm 249.915mm');
    expect(stage.surface.map(row => row[2])).toEqual(['Portrait copy', 'Landscape copy']);
    if (index === 2) expect(stage.surface).toEqual(result.measurements[1].surface);
  }
  for (const [index, saved] of result.folder.entries()) {
    expect(saved.layout).toEqual(result.folder[0].layout);
    expect(saved.layout.map(r => r.exact)).toEqual(['296.9965555555556 419.989', '249.915 350.015']);
    expect(saved.layout[0].standard).toEqual(['16838', '23810']);
    await writeFile(testInfo.outputPath('custom-paper-save-' + (index + 1) + '.docx'), Buffer.from(saved.output));
  }
  expect(result.edited).toEqual([[16838 * 25.4 / 1440, 24000 * 25.4 / 1440, 'portrait'], standard[1]]);
  expect(result.invalid).toEqual([[16838 * 25.4 / 1440, 23810 * 25.4 / 1440, 'portrait'], standard[1]]);
});

import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';
import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
async function fixture() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  zip.file('word/styles.xml', `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
    <w:style w:type="paragraph" w:styleId="OwnSpacing"><w:name w:val="Own spacing"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="60"/></w:pPr></w:style>
    <w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style>
    <w:style w:type="table" w:styleId="CompactBase"><w:name w:val="Compact base"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>
    <w:style w:type="table" w:styleId="CompactChild"><w:name w:val="Compact child"/><w:basedOn w:val="CompactBase"/></w:style></w:styles>`);
  const p = (text: string, properties = '') => `<w:p>${properties ? `<w:pPr>${properties}</w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`;
  const cell = (html: string, span = 1) => `<w:tc><w:tcPr><w:tcW w:w="${span * 3000}" w:type="dxa"/>${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ''}</w:tcPr>${html}</w:tc>`;
  const nested = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cell(p('NESTED DEFAULT'))}</w:tr></w:tbl>`;
  zip.file('word/document.xml', `<w:document xmlns:w="${W}"><w:body>${p('BEFORE DEFAULT')}<w:tbl><w:tblPr><w:tblStyle w:val="CompactChild"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
    <w:tr>${cell(p('INHERITED TABLE'))}${cell(p('PARAGRAPH OVERRIDE', '<w:pStyle w:val="OwnSpacing"/>') + p('DIRECT OVERRIDE', '<w:spacing w:after="90"/>'))}</w:tr>
    <w:tr>${cell(p('OUTER TABLE') + nested + p('OUTER TAIL'), 2)}</w:tr></w:tbl>${p('AFTER DEFAULT')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

test('table paragraph spacing inherits below paragraph/direct formatting, stays local to the nearest table and survives two saves', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async bytes => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    let file = new File([new Uint8Array(bytes)], 'table-style.docx');
    const states = [], saves: number[][] = []; let initial: any;
    for (let stage = 0; stage < 3; stage++) {
      const loaded = await importFile(file), doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; initial ||= doc;
      const root = document.createElement('div'); root.innerHTML = generateHTML(doc.content, editorExtensions);
      const paragraphs = [...root.querySelectorAll('p')].map(p => ({ text: p.textContent, after: p.style.marginBottom, line: p.style.lineHeight }));
      const layout = await createPrintLayout(root.innerHTML, doc.page);
      try { states.push({ paragraphs, tables: root.querySelectorAll('table').length, cells: root.querySelectorAll('td,th').length, spans: [...root.querySelectorAll('[colspan]')].map(c => c.getAttribute('colspan')), text: layout.element.textContent, pages: layout.pages }); }
      finally { layout.dispose(); }
      if (stage < 2) { const bytes = [...new Uint8Array(await (await exportDocx(doc)).arrayBuffer())]; saves.push(bytes); file = new File([new Uint8Array(bytes)], 'again.docx'); }
    }
    return { states, saves, initial };
  }, [...await fixture()]);
  for (const [i, bytes] of result.saves.entries()) await writeFile(test.info().outputPath(`textbox-save-table-style-${i + 1}.docx`), Buffer.from(bytes));
  await writeFile(test.info().outputPath('table-style-states.json'), JSON.stringify(result.states, null, 2));
  const expected: Record<string, [number, number]> = { 'BEFORE DEFAULT': [160 / 15, 259 / 240], 'AFTER DEFAULT': [160 / 15, 259 / 240], 'NESTED DEFAULT': [160 / 15, 259 / 240], 'INHERITED TABLE': [0, 1], 'OUTER TABLE': [0, 1], 'OUTER TAIL': [0, 1], 'PARAGRAPH OVERRIDE': [4, 1], 'DIRECT OVERRIDE': [6, 1] };
  for (const state of result.states) {
    expect(state.tables).toBe(2); expect(state.cells).toBe(4); expect(state.spans).toContain('2'); expect(state.pages).toBe(1);
    for (const [text, [after, line]] of Object.entries(expected)) {
      const p = state.paragraphs.find(p => p.text === text)!;
      expect(p, text).toBeTruthy(); expect(parseFloat(p.after), text + ' after').toBeCloseTo(after, 3); expect(parseFloat(p.line), text + ' line').toBeCloseTo(line, 3);
    }
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'table-style.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.initial)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true }), cell = editor.locator('td').first().locator('p').first();
  await expect(cell).toHaveCSS('margin-bottom', '0px'); await cell.click();
  await page.keyboard.press(process.platform === 'darwin' ? 'Meta+ArrowRight' : 'End');
  await expect.poll(() => cell.evaluate(p => {
    const s = getSelection(); if (!s?.isCollapsed || !s.anchorNode || !p.contains(s.anchorNode)) return null;
    const r = document.createRange(); r.selectNodeContents(p); r.setEnd(s.anchorNode, s.anchorOffset); return r.toString();
  })).toBe('INHERITED TABLE');
  await page.keyboard.type(' EDIT');
  await expect(cell).toHaveText('INHERITED TABLE EDIT'); await page.keyboard.press('ControlOrMeta+z'); await expect(cell).toHaveText('INHERITED TABLE'); await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(cell).toHaveText('INHERITED TABLE EDIT');
  await expect(editor.locator('table')).toHaveCount(2); await expect(editor.locator('td[colspan="2"]')).toHaveCount(1);
});

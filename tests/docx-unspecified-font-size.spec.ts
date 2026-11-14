import { test, expect } from '@playwright/test';
import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
async function input(options: { documentSize?: number; styleSize?: number; runSize?: number } = {}) {
  const z = new JSZip();
  z.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  z.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  z.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  z.file('word/styles.xml', `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>${options.documentSize ? `<w:sz w:val="${options.documentSize * 2}"/>` : ''}</w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>${options.styleSize ? `<w:rPr><w:sz w:val="${options.styleSize * 2}"/></w:rPr>` : ''}</w:style></w:styles>`);
  const run = `<w:r>${options.runSize ? `<w:rPr><w:sz w:val="${options.runSize * 2}"/></w:rPr>` : ''}<w:t>Unspecified size keeps Word fallback</w:t></w:r>`;
  z.file('word/document.xml', `<w:document xmlns:w="${W}"><w:body><w:p>${run}</w:p><w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p>${run}</w:p></w:tc></w:tr></w:tbl><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>`);
  return [...await z.generateAsync({ type: 'uint8array' })];
}

for (const scenario of [
  { name: 'no size in any source layer uses Word compatible ten points', options: {}, expected: 10 },
  { name: 'explicit document default eleven points is preserved', options: { documentSize: 11 }, expected: 11 },
  { name: 'paragraph style overrides the default', options: { documentSize: 11, styleSize: 12 }, expected: 12 },
  { name: 'direct run size overrides both inherited sizes', options: { documentSize: 11, styleSize: 12, runSize: 9 }, expected: 9 },
]) test(scenario.name, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async bytes => {
    const { importFile } = await import('/src/document-io.ts');
    const loaded = await importFile(new File([new Uint8Array(bytes)], 'font-size.docx'));
    const dom = new DOMParser().parseFromString(loaded.html!, 'text/html');
    return [...dom.querySelectorAll('p')].filter(p => p.textContent).map(p => {
      const run = p.querySelector<HTMLElement>('span')!;
      return { size: run.style.fontSize, family: run.style.fontFamily, text: p.textContent };
    });
  }, await input(scenario.options));
  expect(result).toHaveLength(2);
  for (const p of result) {
    expect(p.size).toBe(`${scenario.expected}pt`);
    expect(p.family).toContain('Calibri');
    expect(p.text).toBe('Unspecified size keeps Word fallback');
  }
});

test('implicit Word size and source font survive two real DOCX saves', async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async bytes => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    let loaded = await importFile(new File([new Uint8Array(bytes)], 'unspecified.docx'));
    const out = [];
    for (let i = 0; i < 3; i++) {
      const dom = new DOMParser().parseFromString(loaded.html!, 'text/html');
      out.push([...dom.querySelectorAll('p')].filter(p => p.textContent).map(p => ({ text: p.textContent, size: p.querySelector<HTMLElement>('span')?.style.fontSize, family: p.querySelector<HTMLElement>('span')?.style.fontFamily })));
      if (i < 2) loaded = await importFile(new File([await exportDocx({ ...loaded.document, content: parseDocumentHtml(loaded.html!) })], 'saved.docx'));
    }
    return out;
  }, await input());
  for (const state of states) {
    expect(state).toHaveLength(2);
    for (const p of state) expect(p).toMatchObject({ text: 'Unspecified size keeps Word fallback', size: '10pt', family: 'Calibri' });
  }
  expect(states[1]).toEqual(states[0]);
  expect(states[2]).toEqual(states[0]);
});

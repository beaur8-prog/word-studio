import { expect, test } from '@playwright/test';
import { BorderStyle, Document, Header, Packer, Paragraph, Table, TableCell, TableRow, TextRun, type IParagraphOptions } from 'docx';
import fs from 'node:fs/promises';

const horizontal = { style: BorderStyle.SINGLE, size: 3, color: '007055', space: 0 };
const side = { ...horizontal, space: 9 };
const sides = { left: side, right: side };
const open = { ...sides, top: horizontal };
const close = { ...sides, bottom: horizontal };
const whole = { ...open, ...close, between: { style: BorderStyle.NIL } };
const p = (text: string, border?: IParagraphOptions['border'], extra?: Partial<IParagraphOptions>) => new Paragraph({
  text, border, indent: { left: 171, right: 171 }, spacing: { before: 0, after: 0 }, ...extra,
});
const gap = () => p('Separator');

test('complete direct paragraph border groups become editable standard boxes without connecting ambiguous neighbors', async ({ page }, info) => {
  const first = [
    p('Connected first', open), p('Connected middle', sides), p('Connected last', close), gap(),
    p('Existing bottom border', { bottom: horizontal }), gap(),
    p('Existing complete box', whole), gap(),
    p('Incomplete opener', open), p('Incomplete content', sides), gap(),
    p('Unequal opener', open), p('Unequal side space', { ...sides, left: { ...side, space: 8 } }), p('Unequal closer', close), gap(),
    p('Shifted opener', open), p('Shifted closer', close, { indent: { left: 160, right: 171 } }), gap(),
    p('Spaced opener', open), p('Spaced closer', close, { spacing: { before: 120 } }), gap(),
    p('Break opener', open), p('', sides, { children: [new TextRun('Explicit'), new TextRun({ break: 1, text: 'Break' })] }),
    p('Break closer', close), gap(),
    p('Page opener', open), p('Page closer', close, { pageBreakBefore: true }), gap(),
    p('Table opener', open), new Table({ rows: [new TableRow({ children: [new TableCell({ children: [
      p('Cell opener', open), p('Cell closer', close),
    ] })] })] }), p('Table closer', close), gap(),
    p('Adjacent first opener', open), p('Adjacent first closer', close), p('Adjacent second opener', open), p('Adjacent second closer', close), gap(),
    p('Adjacent complete box', whole), p('Next to complete opener', open), p('Next to complete closer', close), gap(),
    p('Section opener', open),
  ];
  const fixture = new Document({ sections: [
    { headers: { default: new Header({ children: [p('Header opener', open), p('Header closer', close)] }) }, children: first },
    { children: [p('Section closer', close)] },
  ] });
  const bytes = [...await Packer.toBuffer(fixture)];
  await page.goto('/');
  const actual = await page.evaluate(async bytes => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { default: ZIP } = await import('/node_modules/.vite/deps/jszip.js');
    const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const open = async (file: File) => {
      const loaded = await importFile(file);
      return { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
    };
    let doc = await open(new File([Uint8Array.from(bytes)], 'Border groups.docx'));
    const states = [], saves = [];
    for (let stage = 0; stage < 3; stage++) {
      const html = generateHTML(doc.content, editorExtensions);
      const rendered = await createPrintLayout(html, doc.page);
      try {
        const p = [...rendered.element.querySelectorAll<HTMLElement>('.print-content p')];
        states.push({
          pages: rendered.pages,
          boxes: p.filter(el=>el.hasAttribute('data-paragraph-border-box')).map(el => [el.textContent,
            el.hasAttribute('data-paragraph-border-box-start'), el.hasAttribute('data-paragraph-border-box-end')]),
          bottom: p.filter(el=>el.hasAttribute('data-word-paragraph-bottom-border')).map(el=>el.textContent),
          text: p.map(el=>el.textContent),
          headerBox: rendered.element.querySelectorAll('.print-header [data-paragraph-border-box]').length,
        });
      } finally { rendered.dispose(); }
      if (stage === 2) break;
      const blob = await exportDocx(doc);
      const saved = await blob.arrayBuffer(), archive = await ZIP.loadAsync(saved);
      const xml = new DOMParser().parseFromString(await archive.file('word/document.xml').async('string'), 'application/xml');
      const borders = [...xml.getElementsByTagNameNS(ns, 'pBdr')].map(node =>
        [...node.children].map(edge=>[edge.localName, edge.getAttributeNS(ns,'val')]));
      saves.push({ bytes: [...new Uint8Array(saved)], borders });
      doc = await open(new File([blob], 'Saved.docx'));
    }
    return { states, saves };
  }, bytes);

  for (const state of actual.states) {
    expect(state.boxes).toEqual([
      ['Connected first', true, false], ['Connected middle', false, false], ['Connected last', false, true],
      ['Existing complete box', true, true], ['Adjacent complete box', true, true],
    ]);
    expect(state.bottom).toEqual(['Existing bottom border']);
    expect(state.headerBox).toBe(0);
    expect(state.text.filter(Boolean)).toEqual(actual.states[0].text.filter(Boolean));
    expect(state.pages).toBe(actual.states[0].pages);
  }
  for (const [index, save] of actual.saves.entries()) {
    const complete = [['top','single'],['left','single'],['bottom','single'],['right','single'],['between','nil']];
    expect(save.borders).toEqual([complete, complete, complete, [['bottom','single']], complete, complete]);
    await fs.writeFile(info.outputPath(`textbox-save-connected-paragraph-borders-${index + 1}.docx`), Buffer.from(save.bytes));
  }
});

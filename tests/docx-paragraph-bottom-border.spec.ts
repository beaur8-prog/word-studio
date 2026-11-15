import { writeFile } from 'node:fs/promises';
import { expect, test, type Page } from '@playwright/test';
import JSZip from 'jszip';

const word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const relationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const packageRelationships = 'http://schemas.openxmlformats.org/package/2006/relationships';
const docx = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
const positive = { style: 'single', size: 8, space: 4, color: '94C600' } as const;
const firstText = 'Direct  rule\tvalue\nContinuation';
const expectedText = [firstText, 'Inherited rule', 'Direct none', 'Direct nil', 'Unsupported two sides', 'Plain control',
  'Invalid size token', 'Invalid space token', 'Empty space token'];
const expectedBorders = [positive, positive, { style: 'none' }, { style: 'nil' }, null, null, null, null, null];

async function fixture() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="' + docx + 'document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="' + docx + 'styles+xml"/></Types>');
  zip.file('_rels/.rels',
    '<Relationships xmlns="' + packageRelationships + '"><Relationship Id="rId1" Type="' + relationships +
    '/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/_rels/document.xml.rels',
    '<Relationships xmlns="' + packageRelationships + '"><Relationship Id="rId1" Type="' + relationships +
    '/styles" Target="styles.xml"/></Relationships>');
  const bottom = '<w:bottom w:val="single" w:sz="8" w:space="4" w:color="94C600"/>';
  const border = (edges: string) => '<w:pBdr>' + edges + '</w:pBdr>';
  const style = '<w:pStyle w:val="BorderChild"/>';
  zip.file('word/styles.xml',
    '<w:styles xmlns:w="' + word + '">' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="BorderBase"><w:name w:val="Border Base"/><w:basedOn w:val="Normal"/>' +
    '<w:pPr>' + border(bottom) + '</w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="BorderChild"><w:name w:val="Border Child"/><w:basedOn w:val="BorderBase"/></w:style>' +
    '</w:styles>');
  const paragraph = (properties: string, content: string) => '<w:p>' +
    (properties ? '<w:pPr>' + properties + '</w:pPr>' : '') + '<w:r>' + content + '</w:r></w:p>';
  const text = (value: string) => '<w:t>' + value + '</w:t>';
  zip.file('word/document.xml', '<w:document xmlns:w="' + word + '"><w:body>' +
    paragraph(border(bottom), '<w:t xml:space="preserve">Direct  rule</w:t><w:tab/><w:t>value</w:t><w:br/><w:t>Continuation</w:t>') +
    paragraph(style, text('Inherited rule')) +
    paragraph(style + border('<w:bottom w:val="none"/>'), text('Direct none')) +
    paragraph(style + border('<w:bottom w:val="nil"/>'), text('Direct nil')) +
    paragraph(border('<w:top w:val="single" w:sz="8" w:color="4488AA"/>' + bottom), text('Unsupported two sides')) +
    paragraph('', text('Plain control')) +
    paragraph(border(bottom.replace('w:sz="8"', 'w:sz="2e0"')), text('Invalid size token')) +
    paragraph(border(bottom.replace('w:space="4"', 'w:space="0x1"')), text('Invalid space token')) +
    paragraph(border(bottom.replace('w:space="4"', 'w:space=""')), text('Empty space token')) +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" ' +
    'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function inspectStandardXml(page: Page, bytes: Buffer) {
  const zip = await JSZip.loadAsync(bytes);
  const body = await zip.file('word/document.xml')!.async('string');
  const styles = await zip.file('word/styles.xml')!.async('string');
  return page.evaluate(({ body, styles, word }) => {
    const parser = new DOMParser();
    const document = parser.parseFromString(body, 'application/xml');
    const stylesheet = parser.parseFromString(styles, 'application/xml');
    const ww = (element: Element, name: string) => element.namespaceURI === word && element.localName === name;
    const children = (element: Element | undefined, name: string) =>
      [...(element?.children ?? [])].filter(child => ww(child, name));
    const child = (element: Element | undefined, name: string) => children(element, name)[0];
    const attr = (element: Element | undefined, name: string) => element?.getAttributeNS(word, name) ?? null;
    const definitions = children(stylesheet.documentElement, 'style');
    const byId = new Map(definitions.map(style => [attr(style, 'styleId'), style]));
    const defaultStyle = definitions.find(style => attr(style, 'type') === 'paragraph' && ['1', 'true'].includes(attr(style, 'default') || ''));
    return children(child(document.documentElement, 'body'), 'p').map(paragraph => {
      const properties = child(paragraph, 'pPr');
      const direct = children(properties, 'pBdr');
      const requested = attr(child(properties, 'pStyle'), 'val');
      let inherited: Element | undefined, active = requested ? byId.get(requested) : defaultStyle;
      const seen = new Set<Element>();
      while (active && !seen.has(active) && !inherited) {
        seen.add(active);
        inherited = child(child(active, 'pPr'), 'pBdr');
        active = byId.get(attr(child(active, 'basedOn'), 'val'));
      }
      const effective = direct[0] ?? inherited;
      const ink = children(paragraph, 'r').flatMap(run => [...run.children]).map(element => {
        if (ww(element, 't')) return element.textContent;
        if (ww(element, 'tab')) return '\t';
        if (ww(element, 'br')) return '\n';
        return '';
      }).join('');
      const sides = [...(effective?.children ?? [])];
      const bottom = child(effective, 'bottom');
      const style = attr(bottom, 'val');
      return { text: ink, directCount: direct.length, directBottom: attr(child(direct[0], 'bottom'), 'val'),
        sides: sides.map(side => side.namespaceURI === word ? side.localName : side.tagName),
        border: style === 'none' || style === 'nil' ? { style } : style ? {
          style, size: Number(attr(bottom, 'sz')), space: Number(attr(bottom, 'space')), color: attr(bottom, 'color'),
        } : null };
    });
  }, { body, styles, word });
}

test('Word main-story bottom borders retain their values and explicit removal through two DOCX saves', async ({ page }, info) => {
  await page.goto('/');
  const input = await fixture();
  const original = await inspectStandardXml(page, input);
  expect(original.map(row => row.text)).toEqual(expectedText);
  expect(original[1]).toMatchObject({ directCount: 0, sides: ['bottom'], border: positive });
  expect(original[4].sides).toEqual(['top', 'bottom']);

  const { states, saves } = await page.evaluate(async source => {
    const { importFile, exportDocx, validateDocument } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    type Node = ReturnType<typeof parseDocumentHtml>;
    const textOf = (node: Node): string => {
      if (node.type === 'documentTab') return '\t';
      if (node.type === 'hardBreak') return '\n';
      return node.text ?? (node.content ?? []).map(textOf).join('');
    };
    const read = (value: string | null | undefined) => typeof value === 'string' ? JSON.parse(value) : null;
    const states = [], saves: number[][] = [];
    let file = new File([new Uint8Array(source)], 'self-authored-paragraph-border.docx');
    for (let pass = 0; pass < 3; pass++) {
      const imported = await importFile(file);
      const html = new DOMParser().parseFromString(imported.html!, 'text/html');
      const content = parseDocumentHtml(imported.html!);
      const model = validateDocument({ ...imported.document, content });
      states.push({
        markers: [...html.querySelectorAll('body > p')].map(paragraph => read(paragraph.getAttribute('data-word-paragraph-bottom-border'))),
        paragraphs: (content.content ?? []).map(node => ({ type: node.type, text: textOf(node),
          border: read(node.attrs?.wordParagraphBottomBorder) })),
      });
      if (pass < 2) {
        const blob = await exportDocx(model);
        saves.push([...new Uint8Array(await blob.arrayBuffer())]);
        file = new File([blob], 'self-authored-paragraph-border.docx');
      }
    }
    return { states, saves };
  }, [...input]);

  for (const [index, bytes] of saves.entries()) {
    const output = Buffer.from(bytes);
    await writeFile(info.outputPath('textbox-save-paragraph-bottom-border-' + (index + 1) + '.docx'), output);
    const xml = await inspectStandardXml(page, output);
    expect(xml.map(row => row.text)).toEqual(expectedText);
    expect(xml.map(row => row.border)).toEqual(expectedBorders);
    expect(xml.map(row => row.sides)).toEqual(expectedBorders.map(border => border ? ['bottom'] : []));
    expect(xml[2]).toMatchObject({ directBottom: 'none', directCount: 1 });
    expect(xml[3]).toMatchObject({ directBottom: 'nil', directCount: 1 });
  }
  for (const snapshot of states) {
    expect(snapshot.markers).toEqual(expectedBorders);
    expect(snapshot.paragraphs).toEqual(expectedText.map((text, index) => ({
      type: 'paragraph', text, border: expectedBorders[index],
    })));
  }
});

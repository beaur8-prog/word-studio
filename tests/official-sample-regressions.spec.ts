import { test, expect, type Page } from '@playwright/test';
import { Document, Packer, Paragraph } from 'docx';
import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const margins = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>';
const rels = (body: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
const relationship = (id: string, type: string, target: string, external = false) => `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`;
async function fixture(body: string, section = margins, parts: Record<string, string | Buffer> = {}) {
  const zip = await JSZip.loadAsync(await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph('fixture')] }] })));
  zip.file('word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr>${section}</w:sectPr></w:body></w:document>`);
  for (const [name, value] of Object.entries(parts)) zip.file(name, value);
  if (Object.keys(parts).some(name => name.endsWith('.png'))) {
    const types = await zip.file('[Content_Types].xml')!.async('string');
    if (!types.includes('Extension="png"')) zip.file('[Content_Types].xml', types.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>'));
  }
  return Buffer.from(await zip.generateAsync({ type: 'uint8array' }));
}
async function convert(page: Page, bytes: Buffer) {
  return page.evaluate(async bytes => {
    // @ts-expect-error Vite module.
    const io = await import('/src/document-io.ts');
    // @ts-expect-error Vite module.
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const first = await io.importFile(new File([new Uint8Array(bytes)], 'sample.docx'));
    const doc = { ...first.document, content: parseDocumentHtml(first.html) }; io.validateDocument(doc);
    const blob = await io.exportDocx(doc);
    const second = await io.importFile(new File([blob], 'again.docx'));
    const again = { ...second.document, content: parseDocumentHtml(second.html) }; io.validateDocument(again);
    return { doc, again, html: first.html, bytes: [...new Uint8Array(await blob.arrayBuffer())] };
  }, [...bytes]);
}
function nodes(root: any, type: string): any[] { return [...(root.type === type ? [root] : []), ...(root.content || []).flatMap((n: any) => nodes(n, type))]; }
async function open(page: Page, doc: unknown) {
  await page.locator('input[type=file]').first().setInputFiles({ name: 'sample.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(doc)) });
  const firstText = nodes((doc as any).content, 'text')[0]?.text;
  if (firstText) await expect(page.locator('.document-stack .tiptap')).toContainText(firstText);
}
async function preview(page: Page) {
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  await expect(page.locator('.print-page').first()).toBeVisible();
}
test.beforeEach(async ({ page }) => { await page.goto('/'); });

test('external and internal links retain their styled text and real bookmark through DOCX save', async ({ page }) => {
  const bytes = await fixture(`<w:p><w:bookmarkStart w:id="1" w:name="destination"/><w:r><w:t>Target</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>
    <w:p><w:hyperlink r:id="external"><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>External bold italic</w:t></w:r></w:hyperlink>
    <w:hyperlink w:anchor="destination"><w:r><w:rPr><w:b/></w:rPr><w:t>Internal bold</w:t></w:r></w:hyperlink></w:p>`, margins,
    { 'word/_rels/document.xml.rels': rels(relationship('external', 'hyperlink', 'https://example.com/path', true)) });
  const result = await convert(page, bytes);
  for (const doc of [result.doc, result.again]) {
    const text = nodes(doc.content, 'text');
    for (const [value, href, marks] of [['External bold italic', 'https://example.com/path', ['bold','italic']], ['Internal bold','#destination',['bold']]] as const) {
      const run = text.find(n => n.text === value); expect(run).toBeTruthy();
      expect(run.marks.find((m: any) => m.type === 'link').attrs.href).toBe(href);
      for (const mark of marks) expect(run.marks.some((m: any) => m.type === mark)).toBe(true);
    }
    expect(nodes(doc.content, 'documentBookmark').some(n => n.attrs.name === 'destination')).toBe(true);
  }
  await open(page, result.doc);
  await expect(page.locator('.tiptap a[href]')).toHaveCount(2);
});

test('inherited table borders remain visible and nested tables fit the padded parent', async ({ page }) => {
  const table = (content: string) => `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:type="pct" w:w="5000"/></w:tblPr><w:tblGrid><w:gridCol w:w="12000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>${content}</w:tc></w:tr></w:tbl>`;
  const styles = `<w:styles xmlns:w="${W}"><w:style w:type="table" w:styleId="Base"><w:tblPr><w:tblBorders>${['top','left','bottom','right','insideH','insideV'].map(side => `<w:${side} w:val="single" w:sz="8"/>`).join('')}</w:tblBorders><w:tblCellMar><w:left w:w="120"/><w:right w:w="120"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="table" w:styleId="TableGrid"><w:basedOn w:val="Base"/></w:style></w:styles>`;
  const result = await convert(page, await fixture(table(table(p('Nested first')) + p('Outer after nested')), margins, { 'word/styles.xml': styles }));
  for (const doc of [result.doc, result.again]) expect(nodes(doc.content, 'tableCell').every(cell => cell.attrs.borderTop.includes('solid'))).toBe(true);
  await open(page, result.doc);
  const sizes = await page.locator('.tiptap table').evaluateAll(tables => tables.map(t => {
    const parent = t.parentElement!.closest('td') || t.closest('.tiptap')!;
    const style = getComputedStyle(parent); return { width: t.getBoundingClientRect().width, available: parent.clientWidth - parseFloat(style.paddingLeft) - parseFloat(style.paddingRight) };
  }));
  expect(sizes).toHaveLength(2); for (const size of sizes) expect(size.width).toBeLessThanOrEqual(size.available + 2);
  await preview(page); await expect(page.locator('.print-content td').first()).toHaveCSS('border-top-style', 'solid');
});

test('character style inheritance preserves double underline, strike, script and caps', async ({ page }) => {
  const styles = `<w:styles xmlns:w="${W}">${[['Under','<w:u w:val="double"/>'],['Strike','<w:dstrike/>'],['Super','<w:vertAlign w:val="superscript"/>'],['Sub','<w:vertAlign w:val="subscript"/>'],['Caps','<w:smallCaps/>']].map(([name, props]) => `<w:style w:type="character" w:styleId="${name}"><w:rPr>${props}</w:rPr></w:style>`).join('')}</w:styles>`;
  const result = await convert(page, await fixture('<w:p>' + ['Under','Strike','Super','Sub','Caps'].map(name => `<w:r><w:rPr><w:rStyle w:val="${name}"/></w:rPr><w:t>${name}</w:t></w:r>`).join('') + '</w:p>', margins, { 'word/styles.xml': styles }));
  for (const doc of [result.doc, result.again]) {
    const text = nodes(doc.content, 'text');
    for (const [value, mark] of [['Under','underline'],['Strike','strike'],['Super','superscript'],['Sub','subscript']]) expect(text.find(n => n.text === value).marks.some((m: any) => m.type === mark)).toBe(true);
    expect(text.find(n => n.text === 'Under').marks.find((m: any) => m.type === 'textStyle').attrs.decorationStyle).toBe('double');
    expect(text.find(n => n.text === 'Caps').marks.find((m: any) => m.type === 'textStyle').attrs.smallCaps).toBe(true);
  }
});

test('separated paragraphs with the same numbering id continue and retain indentation and tabs', async ({ page }) => {
  const list = `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr><w:ind w:left="720" w:hanging="360"/><w:tabs><w:tab w:val="left" w:pos="1080"/></w:tabs></w:pPr><w:r><w:t>Numbered</w:t></w:r></w:p>`;
  const numbering = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
  const result = await convert(page, await fixture(list + p('') + list + p('') + list, margins, { 'word/numbering.xml': numbering }));
  for (const doc of [result.doc, result.again]) {
    expect(nodes(doc.content, 'orderedList').map(n => n.attrs.start)).toEqual([3,4,5]);
    const paragraphs = nodes(doc.content, 'paragraph').filter(n => nodes(n, 'text').some(t => t.text === 'Numbered'));
    for (const paragraph of paragraphs) { expect(paragraph.attrs.indent).toBe(48); expect(paragraph.attrs.firstLineIndent).toBe('-24px'); expect(JSON.parse(paragraph.attrs.tabStops)[0].position).toBe(72); }
  }
});

test('continuous sections share a page and an adjacent explicit break creates no spurious blank sheet', async ({ page }) => {
  const middle = `<w:p><w:pPr><w:sectPr>${margins}</w:sectPr></w:pPr></w:p>`;
  const result = await convert(page, await fixture(p('Before') + middle + p('After'), `<w:type w:val="continuous"/>${margins}`));
  expect(nodes(result.again.content, 'sectionBreak')[0].attrs.page).toContain('continuous');
  await open(page, result.doc); await preview(page); await expect(page.locator('.print-page')).toHaveCount(1);
  await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
  const duplicate = await convert(page, await fixture(p('First') + middle + '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' + p('Second')));
  await open(page, duplicate.doc); await preview(page); await expect(page.locator('.print-page')).toHaveCount(2);
  for (const text of await page.locator('.print-content').allTextContents()) expect(text.trim()).not.toBe('');
});

test('columns and per-section editor widths survive DOCX export', async ({ page }) => {
  const result = await convert(page, await fixture(p('Left and right column content '.repeat(30)) + `<w:p><w:pPr><w:sectPr>${margins}<w:cols w:num="2" w:space="720"/></w:sectPr></w:pPr></w:p>` + p('Following continuous section'), `<w:type w:val="continuous"/>${margins}`));
  for (const doc of [result.doc, result.again]) { expect(doc.page.columns.count).toBe(2); expect(nodes(doc.content, 'columnGroup')).toHaveLength(1); }
  await open(page, result.doc); await expect(page.locator('.tiptap .document-columns')).toHaveCSS('column-count','2');
  await preview(page); await expect(page.locator('.print-page')).toHaveCount(1);
  const exported = await JSZip.loadAsync(new Uint8Array(result.bytes)); expect(await exported.file('word/document.xml')!.async('string')).toContain('w:num="2"');
});

test('footnotes and endnotes keep references and print in their own areas', async ({ page }) => {
  const body = `<w:p><w:r><w:t>Body with footnote</w:t><w:footnoteReference w:id="1"/></w:r></w:p><w:p><w:r><w:t>Body with endnote</w:t><w:endnoteReference w:id="2"/></w:r></w:p>`;
  const result = await convert(page, await fixture(body, margins, {
    'word/_rels/document.xml.rels': rels(relationship('foot','footnotes','footnotes.xml')+relationship('end','endnotes','endnotes.xml')),
    'word/footnotes.xml': `<w:footnotes xmlns:w="${W}"><w:footnote w:id="1">${p('Footnote explanation')}</w:footnote></w:footnotes>`,
    'word/endnotes.xml': `<w:endnotes xmlns:w="${W}"><w:endnote w:id="2">${p('Endnote explanation')}</w:endnote></w:endnotes>`,
  }));
  for (const doc of [result.doc, result.again]) { expect(nodes(doc.content,'noteReference')).toHaveLength(2); expect(nodes(doc.content,'documentNote')).toHaveLength(2); }
  await open(page, result.doc); await preview(page);
  await expect(page.locator('.print-footnotes')).toContainText('Footnote explanation');
  await expect(page.locator('.print-content')).not.toContainText('Footnote explanation');
  await expect(page.locator('.print-content')).toContainText('Endnote explanation');
  const positions = await page.evaluate(() => ({ body: document.querySelector('.print-content')!.getBoundingClientRect().bottom, note: document.querySelector('.print-footnotes')!.getBoundingClientRect().top, paper: document.querySelector('.print-page')!.getBoundingClientRect().bottom }));
  expect(positions.note).toBeGreaterThanOrEqual(positions.body); expect(positions.note).toBeLessThan(positions.paper);
});

test('legacy symbol characters render Unicode equivalents and retain original codes when saved', async ({ page }) => {
  const bytes = await fixture(`<w:p><w:r><w:rPr><w:rFonts w:ascii="Wingdings"/></w:rPr><w:t>\uF0FC</w:t></w:r><w:r><w:sym w:font="Symbol" w:char="F061"/></w:r></w:p>`);
  const result = await convert(page, bytes);
  for (const doc of [result.doc,result.again]) {
    const symbols = nodes(doc.content,'documentSymbol'); expect(symbols).toHaveLength(2);
    expect(symbols.map(n => Number(n.attrs.code))).toEqual([0xf0fc,0xf061]);
    expect(symbols.map(n => n.attrs.glyph)).toEqual(['✔','α']);
  }
  await open(page,result.doc); await expect(page.locator('.document-symbol')).toHaveText(['✔','α']);
});

const pixel = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aN1sAAAAASUVORK5CYII=', 'base64');
const pict = '<w:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:shape style="width:222.75pt;height:157.4pt"><v:imagedata r:id="image"/></v:shape></w:pict>';
test('VML image dimensions remain inline without splitting the containing paragraph', async ({ page }) => {
  const result = await convert(page, await fixture(`<w:p><w:r><w:t>Before image</w:t>${pict}<w:t>After image</w:t></w:r></w:p>`, margins, {
    'word/_rels/document.xml.rels': rels(relationship('image','image','media/pixel.png')), 'word/media/pixel.png': pixel,
  }));
  for (const doc of [result.doc,result.again]) {
    expect(doc.content.content).toHaveLength(1); const image = nodes(doc.content,'inlineImage')[0];
    expect(image.attrs.width).toBe(297); expect(image.attrs.height).toBeCloseTo(209.867,0);
  }
  await open(page,result.doc); await expect(page.locator('.tiptap img')).toHaveCSS('width','297px');
});

test('header images and tables survive and reserve space above the first body line', async ({ page }) => {
  const header = `<w:hdr xmlns:w="${W}" xmlns:r="${R}"><w:p><w:r><w:t>Header logo</w:t>${pict}</w:r></w:p><w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc>${p('Header table')}</w:tc></w:tr></w:tbl></w:hdr>`;
  const result = await convert(page,await fixture(p('Body first line'), '<w:headerReference w:type="default" r:id="header"/>'+margins.replace('w:top="1440"','w:top="720"'), {
    'word/_rels/document.xml.rels':rels(relationship('header','header','header1.xml')),
    'word/header1.xml':header,'word/_rels/header1.xml.rels':rels(relationship('image','image','media/pixel.png')),'word/media/pixel.png':pixel,
  }));
  for (const doc of [result.doc,result.again]) { expect(doc.page.headerHtml).toContain('<img'); expect(doc.page.headerHtml).toContain('<table'); }
  await open(page,result.doc); await expect(page.locator('.paper-header img')).toHaveCount(1);
  await expect.poll(async()=>page.evaluate(()=>document.querySelector('.document-stack .tiptap > p')!.getBoundingClientRect().top-document.querySelector('.paper-header')!.getBoundingClientRect().bottom)).toBeGreaterThanOrEqual(0);
  await preview(page); await expect(page.locator('.print-header img')).toHaveCount(1);
  expect(await page.evaluate(()=>document.querySelector('.print-content > p')!.getBoundingClientRect().top-document.querySelector('.print-header')!.getBoundingClientRect().bottom)).toBeGreaterThanOrEqual(0);
});

test('comments retain author, initials and rich formatting through save', async ({ page }) => {
  const result = await convert(page,await fixture('<w:p><w:commentRangeStart w:id="4"/><w:r><w:t>Comment target</w:t></w:r><w:commentRangeEnd w:id="4"/></w:p>',margins,{
    'word/_rels/document.xml.rels':rels(relationship('comments','comments','comments.xml')),
    'word/comments.xml':`<w:comments xmlns:w="${W}"><w:comment w:id="4" w:author="Sample Author" w:initials="SA" w:date="2026-10-18T00:00:00Z"><w:p><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>Rich comment</w:t></w:r></w:p></w:comment></w:comments>`,
  }));
  for (const doc of [result.doc,result.again]) {
    expect(doc.comments[0].author).toBe('Sample Author');expect(doc.comments[0].initials).toBe('SA');
    expect(doc.comments[0].html).toMatch(/strong|font-weight: bold/);expect(doc.comments[0].html).toMatch(/em>|font-style: italic/);
  }
});

test('portrait and landscape sections use their own editor widths and preview paper sizes', async ({ page }) => {
  const body = p('Portrait section')+`<w:p><w:pPr><w:sectPr>${margins}</w:sectPr></w:pPr></w:p>`+p('Landscape section');
  const result = await convert(page, await fixture(body,margins.replace('w:w="11906" w:h="16838"','w:w="16838" w:h="11906" w:orient="landscape"')));
  await open(page,result.doc);await expect(page.locator('.document-section-content')).toHaveCount(2);
  await expect.poll(async()=>page.locator('.document-section-content').evaluateAll(els=>els.map(el=>Math.round(el.getBoundingClientRect().width)))).toEqual([794,1123]);
  await preview(page);await expect(page.locator('.print-page')).toHaveCount(2);
  expect(await page.locator('.print-page').evaluateAll(els=>els.map(el=>el.getBoundingClientRect().width>el.getBoundingClientRect().height))).toEqual([false,true]);
});

test('floating VML retains wrapping and stays on one page with surrounding text', async ({ page }) => {
  const floating = pict.replace('<v:shape ', '<v:shape wrapcoords="0 0 21600 21600 0 21600" ').replace('width:222.75pt', 'position:absolute;margin-left:120.75pt;margin-top:26.9pt;width:222.75pt').replace('</v:shape>', '<w10:wrap xmlns:w10="urn:schemas-microsoft-com:office:word" type="tight"/></v:shape>');
  const result = await convert(page, await fixture(`<w:p><w:r><w:t>Before</w:t>${floating}<w:t>After</w:t></w:r></w:p>` + p('Surrounding text '.repeat(120)), margins, {
    'word/_rels/document.xml.rels': rels(relationship('image','image','media/pixel.png')), 'word/media/pixel.png': pixel,
  }));
  for (const doc of [result.doc,result.again]) {
    const image = nodes(doc.content,'inlineImage')[0]; const layout = JSON.parse(image.attrs.floating);
    expect(layout.wrap.type).toBe(2); expect(layout.horizontalPosition.offset).toBe(1533525); expect(layout.polygon).toEqual([[0,0],[1,1],[0,1]]);
  }
  await open(page,result.doc); await expect(page.locator('.tiptap img[src]')).toHaveCSS('position','absolute');
  await preview(page); await expect(page.locator('.print-page')).toHaveCount(1);
});

test('long unbalanced columns continue onto another page without horizontal overflow or lost text', async ({ page }) => {
  const text = 'Column text with a safe page continuation. '.repeat(200);
  const result = await convert(page, await fixture(p(text), margins+'<w:cols w:num="2" w:space="720"/>', {
    'word/settings.xml': `<w:settings xmlns:w="${W}"><w:compat><w:noColumnBalance/></w:compat></w:settings>`,
  }));
  for (const doc of [result.doc,result.again]) expect(doc.page.columns.balance).toBe(false);
  await open(page,result.doc); await preview(page); expect(await page.locator('.print-page').count()).toBeGreaterThan(1);
  expect((await page.locator('.print-content').allTextContents()).join('')).toBe(text);
  expect(await page.locator('.print-content [data-column-count]').evaluateAll(groups=>groups.every(g=>g.scrollWidth<=g.clientWidth+2))).toBe(true);
});

test('automatic trailing whitespace does not print a blank sheet, while an explicit page break does', async ({ page }) => {
  const result = await page.evaluate(async () => {
    // @ts-expect-error Vite module.
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const page = { size: 'A4', orientation: 'portrait', margin: 25.4, header: '', footer: '', pageNumbers: false };
    const automatic = await createPrintLayout('<p>Body</p>'+Array.from({length:60},()=>'<p><br></p>').join(''),page);
    const a = automatic.pages; automatic.dispose();
    const intentional = await createPrintLayout('<p>Body</p><div data-page-break="true"></div>',page);
    const b = intentional.pages; intentional.dispose(); return [a,b];
  });
  expect(result).toEqual([1,2]);
});

test('mixed multilevel numbering uses the format of each ancestor and preserves its pattern', async ({ page }) => {
  const item = (level: number, text: string) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const numbering = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="1">${[['upperRoman','%1.'],['lowerLetter','%1.%2)']].map(([format,pattern],level)=>`<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${pattern}"/></w:lvl>`).join('')}</w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
  const result = await convert(page,await fixture(item(0,'Parent')+item(1,'Child'),margins,{'word/numbering.xml':numbering}));
  for (const doc of [result.doc,result.again]) expect(nodes(doc.content,'listItem').map(n=>n.attrs.wordListLabel)).toEqual(['I.','I.a)']);
});

test('footnotes in a section with its own paper width remain at that section page bottom in the editor', async ({ page }) => {
  const result = await convert(page, await fixture(p('Portrait body')+`<w:p><w:pPr><w:sectPr>${margins}</w:sectPr></w:pPr></w:p>`+'<w:p><w:r><w:t>Landscape reference</w:t><w:footnoteReference w:id="1"/></w:r></w:p>',margins.replace('w:w="11906" w:h="16838"','w:w="16838" w:h="11906" w:orient="landscape"'),{
    'word/_rels/document.xml.rels':rels(relationship('foot','footnotes','footnotes.xml')),
    'word/footnotes.xml':`<w:footnotes xmlns:w="${W}"><w:footnote w:id="1">${p('Section footnote')}</w:footnote></w:footnotes>`,
  }));
  await open(page,result.doc);
  const note = page.locator('.document-stack .tiptap [data-document-note]');
  await expect(note).toHaveCSS('position','absolute');
  const bounds = await page.evaluate(()=>({note:document.querySelector('.tiptap [data-document-note]')!.getBoundingClientRect().toJSON(),section:document.querySelectorAll('.document-section-content')[1].getBoundingClientRect().toJSON()}));
  expect(bounds.note.top).toBeGreaterThan(bounds.section.top); expect(bounds.note.bottom).toBeLessThanOrEqual(bounds.section.bottom+2);
});

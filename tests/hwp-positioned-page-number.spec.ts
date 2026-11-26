import { expect, test } from '@playwright/test';
import * as CFB from 'cfb';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';
import { deflateRawSync, inflateRawSync } from 'node:zlib';

const req = createRequire(import.meta.url);
const variants = ['ordinary', 'even-header', 'hidden', 'hidden-first', 'restart', 'footer', 'mid-flow', 'roman'] as const;
type Variant = typeof variants[number];

function binaryHiddenFirstPageNumber(bytes: Uint8Array) {
  const cfb = CFB.read(bytes, { type: 'array' });
  const header = CFB.find(cfb, '/FileHeader')!, section = CFB.find(cfb, '/BodyText/Section0')!;
  const compressed = !!(header.content[36] & 1);
  const source = Buffer.from(section.content);
  const stream = compressed ? inflateRawSync(source) : source;
  for (let offset = 0; offset < stream.length;) {
    const record = stream.readUInt32LE(offset);
    let length = record >>> 20; offset += 4;
    if (length === 4095) { length = stream.readUInt32LE(offset); offset += 4; }
    if ((record & 1023) === 71 && ((record >>> 10) & 1023) === 1 && length >= 8 && stream.readUInt32LE(offset) === 0x73656364) {
      stream.writeUInt32LE(stream.readUInt32LE(offset + 4) | 32, offset + 4);
      const patched = compressed ? deflateRawSync(stream) : stream;
      section.content = patched; section.size = patched.length;
      return CFB.write(cfb, { type: 'buffer' });
    }
    offset += length;
  }
  throw Error('Expected initial section definition');
}

async function sample(format: 'hwp' | 'hwpx', variant: Variant) {
  const core = path.dirname(req.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  const doc = HwpDocument.createEmpty();
  try {
    doc.createBlankDocument();
    doc.pasteHtml(0, 0, 0, '<p>FIRST BODY</p><p>SECOND BODY</p><p>THIRD BODY</p>');
    for (let p = doc.getParagraphCount(0) - 1; p > 0; p--) doc.insertPageBreak(0, p, 0);
    const entries = unzipSync(doc.exportHwpx()), section = 'Contents/section0.xml';
    let xml = strFromU8(entries[section]);
    const counter = `<hp:ctrl><hp:pageNum pos="BOTTOM_CENTER" formatType="${variant === 'roman' ? 'ROMAN_CAPITAL' : 'DIGIT'}" sideChar="-"/></hp:ctrl>`;
    const before = (needle: string, control: string) => { xml = xml.replace(`<hp:t>${needle}</hp:t>`, `${control}<hp:t>${needle}</hp:t>`); };
    if (variant === 'mid-flow') before('SECOND BODY', counter);
    else xml = xml.replace('</hp:secPr>', `</hp:secPr>${counter}`);
    if (variant === 'hidden') before('SECOND BODY', '<hp:ctrl><hp:pageHiding hideHeader="0" hideFooter="0" hideMasterPage="0" hideBorder="0" hideFill="0" hidePageNum="1"/></hp:ctrl>');
    if (variant === 'hidden-first' && format === 'hwpx') xml = xml.replace('hideFirstPageNum="0"', 'hideFirstPageNum="1"');
    if (variant === 'restart') before('SECOND BODY', '<hp:ctrl><hp:newNum num="7" numType="PAGE"/></hp:ctrl>');
    const furniture = (tag: 'header' | 'footer', id: number, parity: string, value: string) => `<hp:ctrl><hp:${tag} id="${id}" applyPageType="${parity}"><hp:subList id="0" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="${tag === 'footer' ? 'BOTTOM' : 'TOP'}" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0"><hp:p id="${id + 1}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:t>${value}</hp:t></hp:run></hp:p></hp:subList></hp:${tag}></hp:ctrl>`;
    if (variant === 'footer') before('FIRST BODY', furniture('footer', 100100, 'BOTH', 'OWN FOOTER'));
    if (variant === 'even-header') before('FIRST BODY', furniture('header', 100100, 'BOTH', 'REGULAR HEADER') + furniture('header', 100200, 'EVEN', 'EVEN HEADER'));
    expect(xml).toContain(counter);
    entries[section] = strToU8(xml);
    const hwpx = zipSync(entries), authored = new HwpDocument(hwpx);
    try {
      let bytes = format === 'hwp' ? authored.exportHwp() : hwpx;
      if (format === 'hwp' && variant === 'hidden-first') bytes = binaryHiddenFirstPageNumber(bytes);
      const original = new HwpDocument(bytes);
      try {
        const { controls } = JSON.parse(original.getPageControlLayout(0));
        expect(Array.isArray(controls)).toBe(true);
        expect(original.pageCount()).toBe(3);
        const rendered = variant === 'roman' ? 'I' : '1';
        if (variant === 'ordinary' || variant === 'even-header' || variant === 'roman') {
          const first = original.renderPageSvg(0);
          expect(first).toContain(`>${rendered}</text>`);
        }
        return [...bytes];
      } finally { original.free(); }
    } finally { authored.free(); }
  } finally { doc.free(); }
}

for (const format of ['hwp', 'hwpx'] as const) for (const variant of variants)
  test(`${format}, ${variant}: convert only a proven native page number to a live footer`, async ({ page }, info) => {
    const bytes = await sample(format, variant);
    await page.goto('/');
    const accepted = variant === 'ordinary' || variant === 'even-header';
    const result = await page.evaluate(async ({ bytes, format, accepted, keepFooter }) => {
      const { HwpReaderSession } = await import('/src/hwp/reader.ts');
      const { importEditableHwp } = await import('/src/hwp/editing.ts');
      const { exportDocx, importFile } = await import('/src/document-io.ts');
      const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
      const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
      const { createPrintLayout } = await import('/src/print-layout.ts');
      const { default: Zip } = await import('/node_modules/.vite/deps/jszip.js');
      const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      const n = (value: string) => value.replace(/\s/g, '');
      const source = new HwpReaderSession(), name = 'Authored page numbers.' + format;
      const states = [], documents = [];
      try {
        await source.open(new File([Uint8Array.from(bytes)], name));
        let doc = importEditableHwp(await source.editable(), name);
        for (let stage = 0; stage < (accepted || keepFooter ? 3 : 1); stage++) {
          const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
          try {
            const sheets = [...layout.element.querySelectorAll('.print-page')].map(paper => {
              const foot = paper.querySelector('.print-footer')!, number = foot.querySelector('[data-page-number]');
              const text = paper.querySelector('.print-content')!.textContent || '';
              const css = number && getComputedStyle(number);
              return { number: number?.textContent ?? null, label: foot.textContent?.trim(), header: paper.querySelector('.print-header')?.textContent?.trim(),
                body: n(text), color: css?.color, size: css?.fontSize, family: css?.fontFamily };
            });
            states.push({ pages: layout.pages, sheets, numbers: doc.page.pageNumbers, firstPage: doc.page.differentFirstPage, oddEven: doc.page.differentOddEven });
          } finally { layout.dispose(); }
          if (stage === 2 || (!accepted && !keepFooter)) break;
          const blob = await exportDocx(doc), array = await blob.arrayBuffer(), zip = await Zip.loadAsync(array);
          const stories = [];
          for (const file of zip.file(/^word\/footer\d+\.xml$/)) {
            const dom = new DOMParser().parseFromString(await file.async('string'), 'application/xml');
            const tags = (tag: string) => [...dom.getElementsByTagNameNS(ns, tag)];
            stories.push({ fields: tags('instrText').map(el => el.textContent?.trim()),
              literals: tags('t').map(el => el.textContent), alignment: tags('jc').map(el => el.getAttributeNS(ns, 'val')),
              sizes: tags('sz').map(el => el.getAttributeNS(ns, 'val')) });
          }
          const docPart = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
          documents.push({ bytes: [...new Uint8Array(array)], stories,
            footerTypes: [...docPart.getElementsByTagNameNS(ns, 'footerReference')].map(el => el.getAttributeNS(ns, 'type')) });
          const next = await importFile(new File([blob], 'Again.docx'));
          doc = { ...next.document, content: parseDocumentHtml(next.html!) };
        }
        return { states, documents };
      } finally { source.close(); }
    }, { bytes, format, accepted, keepFooter: variant === 'footer' });
    for (const [i, saved] of result.documents.entries())
      await fs.writeFile(info.outputPath(`textbox-save-${format}-${variant}-${i + 1}.docx`), Buffer.from(saved.bytes));
    for (const state of result.states) {
      expect(state.pages).toBe(3);
      expect(state.sheets.map(s => s.body)).toEqual(['FIRSTBODY', 'SECONDBODY', 'THIRDBODY']);
      expect(state.sheets.map(s => s.number)).toEqual(accepted ? ['1', '2', '3'] : [null, null, null]);
      if (accepted) {
        expect(state.sheets.map(s => s.label)).toEqual(['- 1 -', '- 2 -', '- 3 -']);
        expect(state.sheets.every(s => (s.family ?? '').includes('함초롬돋움') && s.size === '13.3333px' && s.color === 'rgb(0, 0, 0)')).toBe(true);
        if (variant === 'even-header') expect(state.sheets.map(s => s.header)).toEqual(['REGULAR HEADER', 'EVEN HEADER', 'REGULAR HEADER']);
      } else if (variant === 'footer') expect(state.sheets.every(s => s.label?.includes('OWN FOOTER'))).toBe(true);
    }
    for (const save of result.documents) {
      const active = save.stories.filter(s => s.fields.includes('PAGE'));
      expect(active.length).toBe(!accepted ? 0 : variant === 'even-header' ? 2 : 1);
      expect(save.footerTypes.includes('first')).toBe(true);
      if (accepted) expect(save.footerTypes.includes('even')).toBe(variant === 'even-header');
      for (const s of active) {
        expect(s.fields).toEqual(['PAGE']);
        expect(s.literals).toEqual(['- ', ' -']);
        expect(s.alignment).toEqual(['center']);
        expect(s.sizes.length).toBeGreaterThan(0);
        expect(s.sizes.every(size => size === '20')).toBe(true);
      }
    }
  });

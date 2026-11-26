import { expect, test } from '@playwright/test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';

const req = createRequire(import.meta.url);

async function source(format: 'hwp' | 'hwpx', kind: 'blank' | 'hidden' | 'break') {
  const core = path.dirname(req.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  const doc = HwpDocument.createEmpty();
  try {
    doc.createBlankDocument();
    doc.pasteHtml(0, 0, 0, '<p><span style="font-size:9pt">BEFORE HEADER</span></p><p>HEADER OWNER</p><p>AFTER HEADER</p>');
    doc.applyParaFormat(0, 2, JSON.stringify({ pageBreakBefore: true }));
    const files = unzipSync(doc.exportHwpx()), section = 'Contents/section0.xml';
    let xml = strFromU8(files[section]);
    const shape = xml.match(/<hp:run charPrIDRef="(\d+)"><hp:t>BEFORE HEADER<\/hp:t><\/hp:run>/)?.[1];
    if (!shape) throw new Error('Expected 9pt source character shape');
    const para = (id: number, body: string) => `<hp:p id="${id}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="${shape}">${body}</hp:run></hp:p>`;
    const field = (type: string) => `<hp:ctrl><hp:autoNum num="4" numType="${type}"><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar="" supScript="0"/></hp:autoNum></hp:ctrl>`;
    const story = kind === 'break' ? para(1000501, '<hp:t>LINE A<hp:lineBreak/>LINE B</hp:t>' + field('PAGE') + field('TOTAL_PAGE')) :
      para(1000501, '<hp:t/>') + para(1000502, '<hp:t/>');
    const header = `<hp:ctrl><hp:header id="1000500" applyPageType="BOTH"><hp:subList id="0" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">${story}</hp:subList></hp:header></hp:ctrl>`;
    xml = xml.replace('<hp:t>HEADER OWNER</hp:t>', header + '<hp:t>HEADER OWNER</hp:t>');
    if (kind === 'hidden') xml = xml.replace('<hp:t>BEFORE HEADER</hp:t>', '<hp:ctrl><hp:pageHiding hideHeader="1" hideFooter="0" hideMasterPage="0" hideBorder="0" hidePageNum="0"/></hp:ctrl><hp:t>BEFORE HEADER</hp:t>');
    files[section] = strToU8(xml);

    const hwpx = zipSync(files), authored = new HwpDocument(hwpx);
    try {
      const bytes = format === 'hwp' ? authored.exportHwp() : hwpx;
      const original = new HwpDocument(bytes);
      try {
        const { paraIndex, paraCount, text } = JSON.parse(original.getHeaderFooter(0, true, 0));
        expect(paraIndex).toBe(1);
        expect(paraCount).toBe(kind === 'break' ? 1 : 2);
        if (kind === 'break') expect(text).toContain('LINE A\nLINE B');
        else {
          expect(text).toBe('\n');
          expect([0, 1].map(i => JSON.parse(original.getHeaderFooterParaInfo(0, true, 0, i)).charCount)).toEqual([0, 0]);
        }
        const findHeader = (node: any): boolean => node?.groupKind?.kind === 'header' ||
          !!node?.children?.some(findHeader) || !!(node?.child && findHeader(node.child));
        expect([0, 1].map(number => findHeader(JSON.parse(original.getPageLayerTreeWithProfile(number, 'print', true, true)).root)))
          .toEqual([kind !== 'hidden', true]);
        return [...bytes];
      } finally { original.free(); }
    } finally { authored.free(); }
  } finally { doc.free(); }
}

for (const format of ['hwp', 'hwpx'] as const) for (const kind of ['blank', 'hidden', 'break'] as const)
  test(`${format}, ${kind}: initial rich header preserves source paragraphs through two DOCX saves`, async ({ page }, info) => {
    const bytes = await source(format, kind);
    await page.goto('/');
    const result = await page.evaluate(async ({ bytes, format }) => {
      const { HwpReaderSession } = await import('/src/hwp/reader.ts');
      const { importEditableHwp } = await import('/src/hwp/editing.ts');
      const { exportDocx, importFile } = await import('/src/document-io.ts');
      const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
      const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
      const { createPrintLayout } = await import('/src/print-layout.ts');
      const { default: Zip } = await import('/node_modules/.vite/deps/jszip.js');
      const session = new HwpReaderSession(), filename = 'Header.' + format;
      await session.open(new File([Uint8Array.from(bytes)], filename));
      try {
        let doc = importEditableHwp(await session.editable(), filename);
        const before = { plain: doc.page.header, rich: !!doc.page.headerHtml };
        const states = [], saves = [];
        const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        for (let stage = 0; stage < 3; stage++) {
          const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
          try { states.push({ pages: layout.pages, headers: [...layout.element.querySelectorAll('.print-header')].map(header => ({
            p: header.querySelectorAll('p').length, br: header.querySelectorAll('br').length,
            fields: [...header.querySelectorAll('[data-page-number]')].map(e => e.getAttribute('data-page-number')),
            text: header.textContent,
          })) }); } finally { layout.dispose(); }
          if (stage === 2) break;
          const blob = await exportDocx(doc);
          const packed = await blob.arrayBuffer();
          const zip = await Zip.loadAsync(packed);
          const headers = [];
          for (const file of zip.file(/^word\/header\d+\.xml$/)) {
            const story = new DOMParser().parseFromString(await file.async('string'), 'application/xml');
            const tags = (tag: string) => [...story.getElementsByTagNameNS(ns, tag)];
            headers.push({ p: tags('p').length, br: tags('br').length, text: tags('t').map(t => t.textContent).join(''),
              fonts: tags('sz').map(e => e.getAttributeNS(ns, 'val')),
              lines: tags('spacing').map(e => e.getAttributeNS(ns, 'line')),
              fields: tags('instrText').map(e => e.textContent?.trim()) });
          }
          saves.push({ headers, bytes: [...new Uint8Array(packed)] });
          const imported = await importFile(new File([blob], 'Again.docx'));
          doc = { ...imported.document, content: parseDocumentHtml(imported.html!) };
        }
        return { before, states, saves };
      } finally { session.close(); }
    }, { bytes, format });

    for (const [index, save] of result.saves.entries())
      await fs.writeFile(info.outputPath(`textbox-save-${format}-${kind}-${index + 1}.docx`), Buffer.from(save.bytes));
    expect(result.states.map(s => s.pages)).toEqual([2, 2, 2]);
    if (kind === 'blank') {
      expect(result.before).toEqual({ plain: '', rich: true });
      for (const { headers } of result.saves)
        expect(headers[0]).toEqual({ p: 2, br: 0, text: '', fonts: ['18', '18'], lines: ['288', '288'], fields: [] });
    } else if (kind === 'hidden') {
      expect(result.before).toEqual({ plain: '\n', rich: false });
      for (const { headers } of result.states) expect(headers[0].text.trim()).toBe('');
    } else {
      expect(result.before).toEqual({ plain: 'LINE A\nLINE B  ', rich: false });
      for (const { headers } of result.saves)
        expect(headers.some(h => h.p === 1 && h.br === 1 && h.fields.includes('PAGE') && h.fields.includes('NUMPAGES'))).toBe(true);
    }
  });

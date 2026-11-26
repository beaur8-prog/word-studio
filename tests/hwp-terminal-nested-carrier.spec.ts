import { expect, test } from '@playwright/test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';

const req = createRequire(import.meta.url);
const cases = ['nested-only', 'authored-blank', 'plain-cell'] as const;

async function source(format: 'hwp' | 'hwpx', kind: typeof cases[number]) {
  const root = path.dirname(req.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(root, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(root, 'rhwp_bg.wasm')) });
  const doc = HwpDocument.createEmpty();
  try {
    doc.createBlankDocument();
    const nested = '<table style="width:465px"><tr><td><p>INNER LEFT</p></td><td><p>INNER RIGHT</p></td></tr><tr><td colspan="2"><p>INNER LAST</p></td></tr></table>';
    doc.pasteHtml(0, 0, 0, `<p>BEFORE</p><table style="width:530px"><tr><td style="padding:3px">${kind === 'plain-cell' ? '<p>OUTER ONLY</p>' : nested}</td></tr></table><p>AFTER</p>`);
    const files = unzipSync(doc.exportHwpx()), section = 'Contents/section0.xml';
    if (kind === 'authored-blank') {
      const xml = strFromU8(files[section]), child = xml.indexOf('</hp:tbl>', xml.indexOf('INNER LAST'));
      const end = xml.indexOf('</hp:p>', child) + '</hp:p>'.length;
      expect(child).toBeGreaterThan(0);
      expect(end).toBeGreaterThan(child);
      const blank = '<hp:p id="891" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:t/></hp:run></hp:p>';
      files[section] = strToU8(xml.slice(0, end) + blank + xml.slice(end));
    }
    const hwpx = zipSync(files), authored = new HwpDocument(hwpx);
    try {
      const bytes = format === 'hwp' ? authored.exportHwp() : hwpx;
      const original = new HwpDocument(bytes);
      try {
        const { para, controlIndex } = JSON.parse(original.getControls()).find(c => c.list === 0 && c.ctrlId === 'tbl');
        expect(original.getCellParagraphCount(0, para, controlIndex, 0)).toBe(kind === 'authored-blank' ? 2 : 1);
        const metrics = JSON.parse(original.getParagraphLineMetrics(0, para, JSON.stringify([{ controlIndex, cellIndex: 0, cellParaIndex: 0 }])));
        expect(metrics.tables.length).toBe(kind === 'plain-cell' ? 0 : 1);
        return [...bytes];
      } finally { original.free(); }
    } finally { authored.free(); }
  } finally { doc.free(); }
}

for (const format of ['hwp', 'hwpx'] as const) for (const kind of cases)
  test(`${format} ${kind}: only the source table's line is occupied`, async ({ page }, info) => {
    await page.goto('/');
    const input = await source(format, kind);
    const result = await page.evaluate(async ({ input, format }) => {
      const { HwpReaderSession } = await import('/src/hwp/reader.ts');
      const { importEditableHwp } = await import('/src/hwp/editing.ts');
      const { exportDocx, importFile } = await import('/src/document-io.ts');
      const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
      const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
      const { default: Zip } = await import('/node_modules/.vite/deps/jszip.js');
      const { createPrintLayout } = await import('/src/print-layout.ts');
      const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      const session = new HwpReaderSession(), file = new File([Uint8Array.from(input)], 'Nested.' + format);
      const stages = [], saves = [];
      try {
        await session.open(file);
        let doc = importEditableHwp(await session.editable(), file.name);
        for (let pass = 0; pass < 3; pass++) {
          const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
          try {
            const body = layout.element.querySelector('.print-content')!;
            const cell = body.querySelector('td')!;
            stages.push({ pages: layout.pages, cellNodes: [...cell.children].map(e => e.tagName),
              text: body.textContent?.replace(/\s/g, ''), blankHeights: [...cell.querySelectorAll(':scope > p:empty')].map(p => p.getBoundingClientRect().height) });
          } finally { layout.dispose(); }
          if (pass === 2) break;
          const saved = await exportDocx(doc), bytes = await saved.arrayBuffer();
          const archive = await Zip.loadAsync(bytes);
          const xml = new DOMParser().parseFromString(await archive.file('word/document.xml')!.async('string'), 'application/xml');
          const tc = xml.getElementsByTagNameNS(ns, 'tc')[0];
          const last = tc.lastElementChild!;
          const pStyle = last.getElementsByTagNameNS(ns, 'pStyle')[0]?.getAttributeNS(ns, 'val');
          saves.push({ bytes: [...new Uint8Array(bytes)], terminal: last.localName, pStyle });
          const imported = await importFile(new File([saved], 'Reopened.docx'));
          doc = { ...imported.document, content: parseDocumentHtml(imported.html!) };
        }
      } finally { session.close(); }
      return { stages, saves };
    }, { input, format });
    for (const [i, save] of result.saves.entries())
      await fs.writeFile(info.outputPath(`textbox-save-carrier-${format}-${kind}-${i + 1}.docx`), Buffer.from(save.bytes));
    const expected = kind === 'nested-only' ? ['TABLE'] : kind === 'authored-blank' ? ['TABLE', 'P'] : ['P'];
    for (const stage of result.stages) {
      expect(stage.pages).toBe(1);
      expect(stage.cellNodes).toEqual(expected);
      expect(stage.text).toBe(kind === 'plain-cell' ? 'BEFOREOUTERONLYAFTER' : 'BEFOREINNERLEFTINNERRIGHTINNERLASTAFTER');
      if (kind === 'authored-blank') expect(stage.blankHeights[0]).toBeGreaterThan(10);
    }
    for (const save of result.saves) {
      expect(save.terminal).toBe('p');
      if (kind === 'nested-only') expect(save.pStyle).toBe('WordStudioTableGap');
      else expect(save.pStyle).not.toBe('WordStudioTableGap');
    }
  });

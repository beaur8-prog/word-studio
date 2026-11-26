import { readFile, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { expect, test } from '@playwright/test';

for (const format of ['hwp', 'hwpx'] as const) {
  test(`${format}: authored blank paragraph sizes and tab stops survive two native saves`, async ({ page }) => {
    await page.goto('/');
    const saves = await page.evaluate(async format => {
      const { exportEditableHwp, importEditableHwp } = await import('/src/hwp/editing.ts');
      const { HwpReaderSession } = await import('/src/hwp/reader.ts');
      const { parseDocumentHtml } = await import('/src/extensions.ts');
      const { defaultPage } = await import('/src/types.ts');
      let document = {
        format: 'word-studio' as const, version: 1 as const, title: 'Authored blanks', id: 'authored-blanks',
        comments: [], updatedAt: new Date().toISOString(), page: defaultPage,
        content: parseDocumentHtml(
          '<p>First witness</p><p style="font-size:6pt;line-height:12.8px"></p>' +
          `<p>Second witness</p><p style="font-size:14pt;font-family:Arial;color:#113355;line-height:29.8667px" data-tab-stops='[{"position":120,"type":"right"}]'></p>` +
          '<p>Third witness</p><div data-page-break="true"></div>' +
          '<p style="font-size:9pt;line-height:19.2px"></p><p>Last witness</p>'),
      };
      const output = [];
      for (let stage = 0; stage < 2; stage++) {
        const saved = await exportEditableHwp(document, format);
        const bytes = new Uint8Array(await saved.blob.arrayBuffer());
        output.push([...bytes]);
        const reader = new HwpReaderSession();
        try {
          const name = `authored-blanks.${format}`;
          await reader.open(new File([bytes], name));
          document = importEditableHwp(await reader.editable(), name);
        } finally { reader.close(); }
      }
      return output;
    }, format);
    const require = createRequire(import.meta.url), core = path.dirname(require.resolve('@rhwp/core'));
    const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
    await init({ module_or_path: await readFile(path.join(core, 'rhwp_bg.wasm')) });
    for (const saved of saves) {
      const doc = new HwpDocument(Uint8Array.from(saved));
      try {
        expect(doc.pageCount()).toBe(2);
        const text = Array.from({ length: doc.getParagraphCount(0) }, (_, index) => doc.getTextRange(0, index, 0, 100));
        expect(text).toEqual(['First witness', '', 'Second witness', '', 'Third witness', '', 'Last witness']);
        const blanks = [1, 3, 5].map(i => {
          const style = JSON.parse(doc.getCharPropertiesAt(0, i, 0));
          const paragraph = JSON.parse(doc.getParaPropertiesAt(0, i));
          return [style.fontSize, paragraph.lineSpacing, paragraph.lineSpacingType];
        });
        expect(blanks).toEqual([[600, 160, 'Percent'], [1400, 160, 'Percent'], [900, 160, 'Percent']]);
        const stops = JSON.parse(doc.getParaPropertiesAt(0, 3)).tabStops;
        expect(stops).toHaveLength(1);
        expect(stops[0].position).toBe(120 * 150);
        const style = JSON.parse(doc.getCharPropertiesAt(0, 3, 0));
        expect(style.fontFamily).toBe('Arial');
        expect(style.textColor).toBe('#113355');
        const second = JSON.parse(doc.getPageTextLayout(1)).runs.find(run => run.text === 'Last witness');
        const paper = JSON.parse(doc.getPageInfo(1));
        expect(second.y - paper.marginTop - paper.marginHeader).toBeCloseTo(19.2, 0);
      } finally { doc.free(); }
    }
  });
}

test('a styled blank between cell paragraphs retains its own height through DOCX and Hangul saves', async ({ page }, info) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { exportEditableHwp } = await import('/src/hwp/editing.ts');
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { defaultPage } = await import('/src/types.ts');
    const cell = (first: string, last: string, width: number, font: number, pitch: number) =>
      `<td colwidth="${width}"><p style="font-size:10pt;line-height:17.333px">${first}</p>` +
      `<p style="font-size:${font}pt;line-height:${pitch}px"></p>` +
      `<p style="font-size:10pt;line-height:17.333px">${last}</p></td>`;
    const anchor = JSON.stringify({ x: 0, y: 1, bottom: 1, reference: 'paragraph', horizontal: 'margin' }).replaceAll('"', '&quot;');
    let document = {
      format: 'word-studio' as const, version: 1 as const, title: 'Table paragraph styles',
      comments: [], updatedAt: new Date().toISOString(),
      page: { ...defaultPage, size: 'Custom' as const, width: 210, height: 100,
        margins: { top: 10, bottom: 10, left: 10, right: 10 } },
      content: parseDocumentHtml('<p>Before\u2060 table</p><table data-table-anchor="' + anchor + '">' +
        '<tr data-row-split="avoid" style="height:250px"><td colwidth="80"><p>Start</p></td>' +
        '<td colwidth="80"><p>Keep</p></td><td colwidth="480"><p>First page</p></td></tr>' +
        '<tr data-row-split="avoid">' + cell('A s', 'B s', 80, 3, 5.2) + cell('A l', 'B l', 80, 8, 13.8667) +
        '<td colwidth="480"><p>Continued page</p></td></tr></table><p>After table</p>'),
    };
    const docx = [];
    for (let stage = 0; stage < 2; stage++) {
      const saved = await exportDocx(document);
      docx.push([...new Uint8Array(await saved.arrayBuffer())]);
      const reopened = await importFile(new File([saved], 'blank-table.docx'));
      document = { ...reopened.document, content: parseDocumentHtml(reopened.html!) };
    }
    const native = [];
    for (const format of ['hwp', 'hwpx'] as const) {
      const saved = await exportEditableHwp(document, format);
      native.push([...new Uint8Array(await saved.blob.arrayBuffer())]);
    }
    return { docx, native };
  });

  for (const [index, bytes] of result.docx.entries())
    await writeFile(info.outputPath(`textbox-save-empty-cell-${index + 1}.docx`), Buffer.from(bytes));

  const require = createRequire(import.meta.url), core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await readFile(path.join(core, 'rhwp_bg.wasm')) });
  for (const bytes of result.native) {
    const doc = new HwpDocument(Uint8Array.from(bytes));
    try {
      expect(doc.pageCount()).toBe(2);
      const tables = JSON.parse(doc.getPageControlLayout(1)).controls.filter(control => control.type === 'table');
      expect(tables).toHaveLength(1);
      expect(doc.getTextRange(0, 0, 0, 100)).toBe('Before\u2060 table');
      const [{ secIdx, paraIdx, controlIdx, colCount, rowCount, w }] = tables;
      expect(w).toBeCloseTo(640, 0);
      expect([rowCount, colCount]).toEqual([1, 3]);
      for (const [cell, label, size] of [[0, 's', 300], [1, 'l', 800]] as const) {
        const address = [secIdx, paraIdx, controlIdx, cell] as const;
        const model = Array.from({ length: doc.getCellParagraphCount(...address) }, (_, p) =>
          doc.getTextInCell(...address, p, 0, 200));
        expect(model).toEqual([`A ${label}`, '', `B ${label}`]);
        const blank = JSON.parse(doc.getCellCharPropertiesAt(...address, 1, 0));
        const spacing = JSON.parse(doc.getCellParaPropertiesAt(...address, 1));
        expect([blank.fontSize, spacing.lineSpacing, spacing.lineSpacingType]).toEqual([size, 130, 'Percent']);
        const runs = JSON.parse(doc.getPageTextLayout(1)).runs.filter(run =>
          run.cellIdx === cell && run.parentParaIdx === paraIdx && run.text.trim());
        expect(runs.at(-1).y - runs[0].y).toBeCloseTo((1000 + size) * 1.3 / 75, 0);
      }
    } finally { doc.free(); }
  }
});

test('preserving explicit tabs in an empty column paragraph adds no character run', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { stageHwpTextNavigation } = await import('/src/hwp/text-navigation.ts');
    const { default:init, HwpDocument } = await import('/node_modules/.vite/deps/@rhwp_core.js');
    const { default:wasm } = await import('/node_modules/@rhwp/core/rhwp_bg.wasm?url');
    await init({ module_or_path:wasm });
    const before = HwpDocument.createEmpty(), paste = HwpDocument.createEmpty();
    let restored;
    const html = '<p>Opening witness</p><p style="line-height:160%"><span style="font-size:11pt"></span></p><p>Following witness</p>';
    const annotated = html.replace('line-height:160%">', 'line-height:160%" data-tab-stops=\'[{"position":120,"type":"right"}]\'>');
    try {
      before.createBlankDocument(); before.pasteHtml(0,0,0,html);
      paste.createBlankDocument();
      const navigation = stageHwpTextNavigation(annotated);
      paste.pasteHtml(0,0,0,navigation.html);
      restored = navigation.restore(HwpDocument,paste);
      const blank = Array.from({length:restored.getParagraphCount(0)},(_,i)=>i).find(i=>
        restored.getTextRange(0,i,0,100)==='' && JSON.parse(restored.getParaPropertiesAt(0,i)).tabStops?.length);
      if (blank === undefined) throw Error('Missing native blank tab stop');
      const paint = doc => JSON.parse(doc.getPageTextLayout(0)).runs.filter(run=>run.text.trim()).map(({text,x,y,w,h})=>[text,x,y,w,h]);
      const props = JSON.parse(restored.getParaPropertiesAt(0,blank));
      return { pageSame:restored.pageCount()===before.pageCount(), paintSame:JSON.stringify(paint(restored))===JSON.stringify(paint(before)),
        tab:props.tabStops[0].position, count:props.tabStops.length };
    } finally { if(restored && restored!==paste) restored.free();paste.free();before.free(); }
  });
  expect(result).toEqual({pageSame:true,paintSame:true,tab:120*150,count:1});
});

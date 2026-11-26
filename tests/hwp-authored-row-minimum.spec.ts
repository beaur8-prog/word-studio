import { expect, test } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';

const req = createRequire(import.meta.url);

async function sourceRows(format: 'hwp' | 'hwpx') {
  const core = path.dirname(req.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  const doc = HwpDocument.createEmpty();
  try {
    doc.createBlankDocument();
    doc.setPageDef(0, JSON.stringify({ ...JSON.parse(doc.getPageDef(0)), width: 45000, height: 37500,
      marginTop: 3000, marginBottom: 3000, marginLeft: 3000, marginRight: 3000, marginHeader: 0, marginFooter: 0 }));
    const para = (text: string) => `<p style="font-size:10pt;line-height:130%;margin:0">${text}</p>`;
    const rows = Array.from({ length: 9 }, (_, i) => `<tr><td style="width:120px">${para('LEFT ' + i)}</td><td style="width:120px">${para('RIGHT ' + i)}</td></tr>`).join('');
    doc.pasteHtml(0, 0, 0, '<table style="width:240px">' + rows + '</table>' + para('AFTER'));
    const table = JSON.parse(doc.getControls()).find(c => c.list === 0 && c.ctrlId === 'tbl');
    const address = [0, table.para, table.controlIndex] as const;
    doc.setTableProperties(...address, JSON.stringify({ pageBreak: 2, repeatHeader: true,
      treatAsChar: false, textWrap: 'TopAndBottom', vertRelTo: 'Para', vertAlign: 'Top',
      vertOffset: -316, outerTop: 283, restrictInPage: true }));
    for (let cell = 0; cell < 18; cell++) {
      const height = cell < 2 ? 2400 : cell === 2 ? 1850 : cell === 3 ? 2982 : 4200;
      doc.setCellProperties(...address, cell, JSON.stringify({ height, isHeader: cell < 2, verticalAlign: 0,
        paddingTop: 141, paddingBottom: 141, applyInnerMargin: true }));
    }
    doc.reflowLinesegs();
    const bytes = format === 'hwp' ? doc.exportHwp() : doc.exportHwpx();
    const reopened = new HwpDocument(bytes);
    try {
      const cells = [2, 3].map(i => JSON.parse(reopened.getCellProperties(...address, i)).height);
      const boxes = JSON.parse(reopened.getTableCellBboxes(...address));
      expect(reopened.pageCount()).toBe(2);
      expect(cells).toEqual([1850, 2982]);
      expect(boxes.filter(box => box.row === 1).map(box => [box.pageIndex, box.h])).toEqual([[0, 39.8], [0, 39.8]]);
    } finally { reopened.free(); }
    return [...bytes] as number[];
  } finally { doc.free(); }
}

for (const format of ['hwp', 'hwpx'] as const) test(`${format}: a complete repeated-table row retains the larger authored cell minimum through two DOCX saves`, async ({ page }, info) => {
  const bytes = await sourceRows(format);
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { default: Zip } = await import('/node_modules/.vite/deps/jszip.js');
    const session = new HwpReaderSession(), filename = 'Authored minimum.' + format;
    try {
      await session.open(new File([new Uint8Array(bytes)], filename));
      let doc = importEditableHwp(await session.editable(), filename);
      const states = [], saves: number[][] = [];
      for (let stage = 0; stage < 3; stage++) {
        const table = doc.content.content?.find(node => node.type === 'table');
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const rows = [...layout.element.querySelectorAll('tr')];
          const target = rows.find(row => row.cells[0]?.textContent?.trim() === 'LEFT 1');
          states.push({
            pages: layout.pages, modelMinimum: table?.content?.[1].attrs?.rowHeight,
            modelGrid: table?.content?.map(row => row.content?.map(cell => [cell.attrs?.colspan, cell.attrs?.rowspan])),
            actualMinimum: target?.getBoundingClientRect().height,
            rowContents: rows.filter(row => [...row.cells].every(cell => cell.tagName === 'TD'))
              .map(row => [...row.cells].map(cell => cell.textContent?.trim() ?? '')).filter(row => row.some(Boolean)),
          });
        } finally { layout.dispose(); }
        if (stage === 2) break;
        const blob = await exportDocx(doc);
        const packed = await blob.arrayBuffer();
        saves.push([...new Uint8Array(packed)]);
        const file = await Zip.loadAsync(packed);
        const xml = await file.file('word/document.xml').async('string');
        const markup = new DOMParser().parseFromString(xml, 'application/xml');
        const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        const height = markup.getElementsByTagNameNS(ns, 'tr')[1]?.getElementsByTagNameNS(ns, 'trHeight')[0];
        if (height?.getAttributeNS(ns, 'hRule') !== 'atLeast' || Number(height.getAttributeNS(ns, 'val')) <= 370)
          throw new Error('The standard DOCX minimum was not exported');
        const reopened = await importFile(new File([blob], 'Again.docx'));
        doc = { ...reopened.document, content: parseDocumentHtml(reopened.html!) };
      }
      return { states, saves };
    } finally { session.close(); }
  }, { bytes, format });
  for (const [save, bytes] of result.saves.entries()) await fs.writeFile(info.outputPath(`textbox-save-${format}-authored-minimum-${save + 1}.docx`), Buffer.from(bytes));
  for (const state of result.states) {
    expect(state.pages).toBe(2);
    expect(state.modelMinimum).toBeCloseTo(39.76, 0);
    expect(state.actualMinimum).toBeCloseTo(39.76, 0);
    expect(state.modelGrid).toEqual(Array.from({ length: 9 }, () => [[1, 1], [1, 1]]));
    expect(state.rowContents).toEqual(Array.from({ length: 8 }, (_, i) => ['LEFT ' + (i + 1), 'RIGHT ' + (i + 1)]));
  }
});

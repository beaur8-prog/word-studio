import { readFile, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { expect, test } from '@playwright/test';

for (const format of ['hwp', 'hwpx'] as const) {
  test(`${format}: a complete merged grid keeps every cell inside its physical paper`, async ({ page }) => {
    await page.goto('/');
    const outputs = await page.evaluate(async format => {
      const { HwpReaderSession } = await import('/src/hwp/reader.ts');
      const { defaultPage } = await import('/src/types.ts');
      const sized = (text, widths, columns = 1, rows = 1) => `<td colspan="${columns}" rowspan="${rows}" colwidth="${widths.join(',')}" style="width:${widths.reduce((a, b) => a + b, 0)}px;height:${rows * 32}px"><p>${text}</p></td>`;
      const row = html => `<tr style="height:32px">${html}</tr>`;
      const body = inconsistent => '<p>Before</p><table>' +
        row(sized('First', [80, 60], 2) + sized('Second', inconsistent ? [18, 21] : [18, 22], 2)) +
        row(sized('Third', [80]) + sized('Fourth', [60]) + sized('Fifth', [18, 22], 2, 2)) +
        row(sized('Sixth', [80]) + sized('Last', [60])) +
        '</table><p>After</p>';
      const files = [];
      const reader = new HwpReaderSession();
      try {
        for (const inconsistent of [false, true]) {
          const saved = await reader.export(body(inconsistent),
            { ...defaultPage, size: 'Custom', width: 110, height: 160, margin: 15 }, format);
          files.push([...saved.bytes]);
        }
      } finally { reader.close(); }
      return files;
    }, format);
    const require = createRequire(import.meta.url), core = path.dirname(require.resolve('@rhwp/core'));
    const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
    await init({ module_or_path: await readFile(path.join(core, 'rhwp_bg.wasm')) });
    for (const [index, bytes] of outputs.entries()) {
      const document = new HwpDocument(Uint8Array.from(bytes));
      try {
        expect(document.pageCount()).toBe(1);
        const paper = JSON.parse(document.getPageInfo(0));
        const controls = JSON.parse(document.getPageControlLayout(0)).controls;
        const table = controls.find(c => c.type === 'table');
        expect(table.rowCount).toBe(3);
        expect(table.colCount).toBe(4);
        const cells = JSON.parse(document.getTableCellBboxes(table.secIdx, table.paraIdx, table.controlIdx));
        expect(cells).toHaveLength(7);
        const text = new Map<string, string>();
        for (const run of JSON.parse(document.getPageTextLayout(0)).runs) {
          const key = `${run.paraIdx}/${run.parentParaIdx}/${run.cellIdx}`;
          text.set(key, (text.get(key) ?? '') + run.text);
        }
        const words = [...text.values()];
        for (const label of ['Before', 'First', 'Second', 'Third', 'Fourth', 'Fifth', 'Sixth', 'Last', 'After'])
          expect(words.filter(value => value === label), label + ' in ' + JSON.stringify(words)).toHaveLength(1);
        if (index === 0) {
          expect(table.w).toBeCloseTo(180, 0);
          expect(table.x + table.w).toBeLessThan(paper.width);
          expect(cells.map(c => c.w)).toEqual([140, 40, 80, 60, 40, 80, 60]);
        } else {
          // An inconsistent authored column cannot establish a verified grid.
          expect(table.x + table.w).toBeGreaterThan(paper.width);
        }
      } finally { document.free(); }
    }
  });
}

test('source cell widths survive two DOCX saves and still yield to a column resize', async ({ page }, info) => {
  const require = createRequire(import.meta.url), core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await readFile(path.join(core, 'rhwp_bg.wasm')) });
  const JSZip = require('jszip');
  const nativeWidths = [6002, 4506, 1358, 1664];
  const nativeWidth = nativeWidths.reduce((sum, width) => sum + width, 0);
  const fixture = HwpDocument.createEmpty();
  let bytes: Uint8Array;
  try {
    fixture.createBlankDocument();
    fixture.setPageDef(0, JSON.stringify({ width: 31181, height: 45354, marginLeft: 4252, marginRight: 4252 }));
    fixture.pasteHtml(0, 0, 0, '<p>Before</p><table>' +
      '<tr><td colspan="2" style="width:140px;height:32px"><p>First</p></td><td colspan="2" style="width:40px;height:32px"><p>Second</p></td></tr>' +
      '<tr><td style="width:80px;height:32px"><p>Third</p></td><td style="width:60px;height:32px"><p>Fourth</p></td>' +
      '<td colspan="2" rowspan="2" style="width:40px;height:64px"><p>Fifth</p></td></tr>' +
      '<tr><td style="width:80px;height:32px"><p>Sixth</p></td><td style="width:60px;height:32px"><p>Last</p></td></tr>' +
      '</table><p>After</p>');
    const zip = await JSZip.loadAsync(fixture.exportHwpx());
    const file = 'Contents/section0.xml';
    const xml = (await zip.file(file).async('string')).replace(/(<hp:tbl\b[^>]*>\s*<hp:sz width=")\d+/, (_: string, prefix: string) => prefix + nativeWidth)
      .replace(/(<hp:cellAddr colAddr="(\d+)"[^>]*\/>\s*<hp:cellSpan colSpan="(\d+)"[^>]*\/>\s*<hp:cellSz width=")\d+/g,
        (_: string, prefix: string, at: string, span: string) => prefix + nativeWidths
          .slice(Number(at), Number(at) + Number(span)).reduce((sum, width) => sum + width, 0));
    zip.file(file, xml);
    bytes = await zip.generateAsync({ type: 'uint8array' });
  } finally { fixture.free(); }

  await page.goto('/');
  const result = await page.evaluate(async bytes => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp, exportEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const native = new HwpReaderSession();
    let doc;
    try {
      const file = new File([Uint8Array.from(bytes)], 'merged-source.hwpx');
      await native.open(file);
      doc = importEditableHwp(await native.editable(), file.name);
    } finally { native.close(); }
    const docx: number[][] = [];
    for (let i = 0; i < 2; i++) {
      const saved = await exportDocx(doc);
      docx.push([...new Uint8Array(await saved.arrayBuffer())]);
      const opened = await importFile(new File([saved], 'roundtrip.docx'));
      doc = { ...opened.document, content: parseDocumentHtml(opened.html!) };
    }
    const outputs = [];
    for (const format of ['hwp', 'hwpx'] as const) {
      const file = await exportEditableHwp(doc, format);
      outputs.push({ resized: false, bytes: [...new Uint8Array(await file.blob.arrayBuffer())] });
    }
    const resize = (node: import('@tiptap/core').JSONContent) => {
      const widths = node.attrs?.colwidth;
      if (widths?.[0] === 80) widths[0] = 82;
      node.content?.forEach(resize);
    };
    resize(doc.content);
    const changed = await exportEditableHwp(doc, 'hwpx');
    outputs.push({ resized: true, bytes: [...new Uint8Array(await changed.blob.arrayBuffer())] });
    return { docx, outputs };
  }, [...bytes]);
  for (const [i, bytes] of result.docx.entries()) await writeFile(info.outputPath(`textbox-save-native-grid-${i + 1}.docx`), Buffer.from(bytes));
  for (const output of result.outputs) {
    const doc = new HwpDocument(Uint8Array.from(output.bytes));
    try {
      const paper = JSON.parse(doc.getPageInfo(0));
      const table = JSON.parse(doc.getPageControlLayout(0)).controls.find(c => c.type === 'table');
      expect(table).toMatchObject({ rowCount: 3, colCount: 4 });
      expect(table.x + table.w).toBeLessThan(paper.width);
      const geometry = JSON.parse(doc.getTableProperties(table.secIdx, table.paraIdx, table.controlIdx));
      expect(geometry.tableWidth).toBe(output.resized ? 182 * 75 : nativeWidth);
      const cells = Array.from({ length: 7 }, (_, cell) =>
        ({ ...JSON.parse(doc.getCellInfo(table.secIdx, table.paraIdx, table.controlIdx, cell)),
          ...JSON.parse(doc.getCellProperties(table.secIdx, table.paraIdx, table.controlIdx, cell)) }));
      const widths = output.resized ? [82, 60, 18, 22].map(px => px * 75) : nativeWidths;
      for (const cell of cells) expect(cell.width).toBe(widths.slice(cell.col, cell.col + cell.colSpan).reduce((sum, width) => sum + width, 0));
      const text = JSON.parse(doc.getPageTextLayout(0)).runs.map(run => run.text).join('').replace(/\s/gu, '');
      expect(text).toBe('BeforeFirstSecondThirdFourthFifthSixthLastAfter');
    } finally { doc.free(); }
  }
});

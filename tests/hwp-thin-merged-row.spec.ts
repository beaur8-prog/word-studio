import { expect, test } from '@playwright/test';
import { createRequire } from 'node:module';

const createFixture = createRequire(import.meta.url)('./hwp-thin-merged-row-fixture.cjs');
let fixtures: Record<'hwp' | 'hwpx', Uint8Array>;
test.beforeAll(async () => { fixtures = await createFixture(); });

for (const format of ['hwp', 'hwpx'] as const) {
  test(`${format}: a short native grid row in a vertical merge keeps its height through two DOCX saves`, async ({ page }) => {
    await page.goto('/');
    const proof = await page.evaluate(async ({ bytes, format }) => {
      const { HwpReaderSession } = await import('/src/hwp/reader.ts');
      const { importEditableHwp } = await import('/src/hwp/editing.ts');
      const { exportDocx, importFile } = await import('/src/document-io.ts');
      const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
      const { createPrintLayout } = await import('/src/print-layout.ts');
      const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
      const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
      const { hwpGridRowHeightNamespace } = await import('/src/docx-table-cell.ts');
      const Word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      const source = new HwpReaderSession();
      try {
        const name = 'crossing-merge.' + format;
        await source.open(new File([new Uint8Array(bytes)], name));
        const stages = [], saved = [];
        let model = importEditableHwp(await source.editable(), name);
        async function inspect() {
          const view = await createPrintLayout(generateHTML(model.content, editorExtensions), model.page);
          try {
            const table = view.element.querySelector<HTMLTableElement>('.print-content table')!;
            return {
              pages: view.pages,
              minima: Array.from(table.rows, row => parseFloat(row.style.height)),
              middleHeight: table.rows[1].getBoundingClientRect().height,
              spans: Array.from(table.rows, row => Array.from(row.cells, cell => cell.rowSpan)),
              text: Array.from(table.querySelectorAll('p'), p => p.textContent),
            };
          } finally { view.dispose(); }
        }
        async function reopen(blob: Blob) {
          const loaded = await importFile(new File([blob], 'crossing-merge.docx'));
          model = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
        }
        let afterStandardEdit;
        for (let stage = 0; stage < 3; stage++) {
          stages.push(await inspect());
          if (stage === 2) break;
          const blob = await exportDocx(model);
          const zip = await JSZip.loadAsync(await blob.arrayBuffer());
          const xml = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'text/xml');
          const row = xml.getElementsByTagNameNS(Word, 'tr')[1];
          const height = row.getElementsByTagNameNS(Word, 'trHeight')[0];
          saved.push({
            native: Number(row.getAttributeNS(hwpGridRowHeightNamespace, 'height')),
            value: height.getAttributeNS(Word, 'val'),
            rule: height.getAttributeNS(Word, 'hRule'),
          });
          if (stage === 1) {
            height.setAttributeNS(Word, 'w:val', '240');
            zip.file('word/document.xml', new XMLSerializer().serializeToString(xml));
            afterStandardEdit = await zip.generateAsync({ type: 'blob' });
          }
          await reopen(blob);
        }
        await reopen(afterStandardEdit!);
        return { stages, saved, external: await inspect() };
      } finally { source.close(); }
    }, { bytes: [...fixtures[format]], format });
    await test.info().attach('row-height', { body: JSON.stringify(proof, null, 2), contentType: 'application/json' });
    for (const stage of proof.stages) {
      expect(stage.pages).toBe(1);
      expect(stage.minima[1]).toBeCloseTo(2, 1);
      expect(stage.middleHeight).toBeCloseTo(2, 1);
      expect(stage.spans).toEqual([[2, 1], [2], [1]]);
      expect(stage.text).toEqual(['UPPER OWNER', 'UPPER RIGHT', 'LOWER OWNER', 'LOWER LEFT']);
    }
    expect(proof.saved).toEqual([{ native: 2, value: '1', rule: 'atLeast' }, { native: 2, value: '1', rule: 'atLeast' }]);
    expect(proof.external.minima[1]).toBeGreaterThan(16);
    expect(proof.external.text).toEqual(proof.stages[0].text);
  });
}

test('a short native row beside a new padded merge keeps its own minimum', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { hwpGridRowHeightNamespace } = await import('/src/docx-table-cell.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const html = '<table><tbody><tr data-hwp-grid-row="true" style="height:6.67px">'
      + '<td rowspan="2" style="padding:7.5px 0"><p>SPANNING TEXT</p></td>'
      + '<td style="padding:0"><p data-word-line-rule="exact" style="line-height:6.67px"></p></td>'
      + '</tr><tr style="height:53px"><td>FOLLOWING CELL</td></tr></tbody></table>';
    const createModel = (html: string) => ({ format: 'word-studio' as const, version: 1 as const, title: 'Native merge',
      id: crypto.randomUUID(), updatedAt: new Date().toISOString(), comments: [], page: defaultPage,
      content: parseDocumentHtml(html) });
    let doc = createModel(html);
    const inspect = (doc: ReturnType<typeof createModel>) => {
      const row = doc.content.content![0].content![0];
      return { height: row.attrs?.rowHeight, grid: row.attrs?.hwpGridRow, rowspan: row.content![0].attrs?.rowspan };
    };
    const reopen = async (file: Blob) => {
      const data = await importFile(new File([file], 'native-merge.docx'));
      return { ...doc, content: parseDocumentHtml(data.html!) };
    };
    const stages = [inspect(doc)], pack = [];
    let changedHeight, changedNeighbour;
    for (let stage = 0; stage < 2; stage++) {
      const file = await exportDocx(doc);
      const zip = await JSZip.loadAsync(await file.arrayBuffer());
      const savedXml = await zip.file('word/document.xml')!.async('string');
      const parse = () => new DOMParser().parseFromString(savedXml, 'text/xml');
      const normal = parse();
      const row = normal.getElementsByTagNameNS(word, 'tr')[0];
      const height = row.getElementsByTagNameNS(word, 'trHeight')[0];
      pack.push([row.getAttributeNS(hwpGridRowHeightNamespace, 'height'), height.getAttributeNS(word, 'val'),
        height.getAttributeNS(word, 'hRule')]);
      if (stage === 1) {
        height.setAttributeNS(word, 'w:val', '240');
        zip.file('word/document.xml', new XMLSerializer().serializeToString(normal));
        changedHeight = inspect(await reopen(await zip.generateAsync({ type: 'blob' })));
        const edit = parse();
        const paragraph = edit.getElementsByTagNameNS(word, 'tc')[1].getElementsByTagNameNS(word, 'p')[0];
        const run = edit.createElementNS(word, 'w:r'), text = edit.createElementNS(word, 'w:t');
        text.textContent = 'ADDED TEXT';
        run.append(text); paragraph.append(run);
        zip.file('word/document.xml', new XMLSerializer().serializeToString(edit));
        changedNeighbour = inspect(await reopen(await zip.generateAsync({ type: 'blob' })));
      }
      doc = await reopen(file);
      stages.push(inspect(doc));
    }
    const ordinary = inspect(await reopen(await exportDocx(createModel(html.replace(' data-hwp-grid-row="true"', '')))));
    return { stages, pack, changedHeight, changedNeighbour, ordinary };
  });
  for (const stage of result.stages) {
    expect(stage.height).toBeCloseTo(6.67, 2);
    expect(stage.grid).toBe(true);
    expect(stage.rowspan).toBe(2);
  }
  expect(result.pack).toEqual([['6.67', '1', 'atLeast'], ['6.67', '1', 'atLeast']]);
  expect(result.changedHeight?.height).toBeGreaterThan(20);
  expect(result.changedNeighbour?.height).toBeGreaterThan(14);
  expect(result.ordinary.height).toBeGreaterThan(14);
  expect(result.changedNeighbour?.grid).not.toBe(true);
  expect(result.ordinary.grid).not.toBe(true);
});

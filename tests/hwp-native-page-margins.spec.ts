import { readFile, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { expect, test } from '@playwright/test';

test('Hancom horizontal margins keep their native units across DOCX and yield to an edited side', async ({ page }, info) => {
  const require = createRequire(import.meta.url);
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await readFile(path.join(core, 'rhwp_bg.wasm')) });
  const source = HwpDocument.createEmpty();
  let original: number[];
  try {
    source.createBlankDocument();
    source.setPageDef(0, JSON.stringify({ width: 59528, height: 84188, marginLeft: 2834, marginRight: 2834,
      marginTop: 2834, marginHeader: 2834, marginBottom: 2834, marginFooter: 1417 }));
    source.pasteHtml(0, 0, 0, '<p>Horizontal space belongs to the paper.</p>');
    original = [...source.exportHwp()];
  } finally { source.free(); }

  await page.goto('/');
  const result = await page.evaluate(async original => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp, exportEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { basePageMargins } = await import('/src/types.ts');
    const input = new File([Uint8Array.from(original)], 'authored-margins.hwp');
    const reader = new HwpReaderSession();
    let doc;
    try {
      await reader.open(input);
      doc = importEditableHwp(await reader.editable(), input.name);
    } finally { reader.close(); }
    const save = async () => {
      const blob = await exportDocx(doc);
      const file = await importFile(new File([blob], 'roundtrip.docx'));
      doc = { ...file.document, content: parseDocumentHtml(file.html!) };
      return [...new Uint8Array(await blob.arrayBuffer())];
    };
    const docx = [await save(), await save()];
    const observed = [doc.page.hwpHorizontalMargins];
    const outputs = [];
    for (const format of ['hwp', 'hwpx'] as const) {
      const out = await exportEditableHwp(doc, format);
      outputs.push([...new Uint8Array(await out.blob.arrayBuffer())]);
    }
    // This edit is within the same Word twip; stale provenance must still drop.
    doc.page.margins = { ...basePageMargins(doc.page), left: 10.006 };
    const inMemory = await exportEditableHwp(doc, 'hwpx');
    outputs.push([...new Uint8Array(await inMemory.blob.arrayBuffer())]);
    docx.push(await save());
    observed.push(doc.page.hwpHorizontalMargins);
    const edited = await exportEditableHwp(doc, 'hwpx');
    outputs.push([...new Uint8Array(await edited.blob.arrayBuffer())]);
    return { docx, outputs, observed };
  }, original);

  expect(result.observed).toEqual([[2834, 2834], [null, 2834]]);
  for (const [index, bytes] of result.docx.entries())
    await writeFile(info.outputPath(`textbox-save-native-margins-${index + 1}.docx`), Buffer.from(bytes));
  const expected = [[2834, 2834], [2834, 2834], [2836, 2834], [2835, 2834]];
  for (const [index, bytes] of result.outputs.entries()) {
    const reopened = new HwpDocument(Uint8Array.from(bytes));
    try {
      const { marginLeft, marginRight } = JSON.parse(reopened.getPageDef(0));
      expect([marginLeft, marginRight]).toEqual(expected[index]);
      expect(reopened.pageCount()).toBe(1);
      const text = JSON.parse(reopened.getPageTextLayout(0)).runs.map(run => run.text).join('');
      expect(text).toContain('Horizontal space belongs to the paper.');
    } finally { reopened.free(); }
  }
});

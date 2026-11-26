import { readFile, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { expect, test } from '@playwright/test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';

test('authored inline table top margins survive two DOCX saves and both Hancom formats', async ({ page }, info) => {
  test.setTimeout(90_000);
  const require = createRequire(import.meta.url);
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await readFile(path.join(core, 'rhwp_bg.wasm')) });

  const original = HwpDocument.createEmpty();
  let input: Uint8Array;
  try {
    original.createBlankDocument();
    const table = (text: string) => '<table><tr><td style="width:180px;height:32px"><p>' + text + '</p></td></tr></table>';
    original.pasteHtml(0, 0, 0, '<p>Before</p>' + table('First grid') + '<p>Between</p>' + table('Second grid') + '<p>After</p>');
    const parts = unzipSync(original.exportHwpx()), section = 'Contents/section0.xml';
    const margins = [0, 141];
    let index = 0;
    parts[section] = strToU8(strFromU8(parts[section]).replace(/<hp:outMargin\b[^>]*\/>/g,
      tag => tag.replace(/\btop="\d+"/, () => 'top="' + margins[index++] + '"')));
    expect(index).toBe(2);
    input = zipSync(parts);
  } finally { original.free(); }

  await page.goto('/');
  const result = await page.evaluate(async bytes => {
    const { importEditableHwp, exportEditableHwp } = await import('/src/hwp/editing.ts');
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const reader = new HwpReaderSession();
    let doc;
    try {
      const file = new File([Uint8Array.from(bytes)], 'inline-margins.hwpx');
      await reader.open(file);
      doc = importEditableHwp(await reader.editable(), file.name);
    } finally { reader.close(); }
    const docx = [];
    for (let i = 0; i < 2; i++) {
      const save = await exportDocx(doc);
      docx.push([...new Uint8Array(await save.arrayBuffer())]);
      const reopened = await importFile(new File([save], 'again.docx'));
      doc = { ...reopened.document, content: parseDocumentHtml(reopened.html!) };
    }
    const native = [];
    for (const format of ['hwp', 'hwpx'] as const) {
      const out = await exportEditableHwp(doc, format);
      native.push([...new Uint8Array(await out.blob.arrayBuffer())]);
    }
    return { docx, native };
  }, [...input]);

  for (const [index, bytes] of result.docx.entries()) {
    await writeFile(info.outputPath(`textbox-save-inline-margin-${index + 1}.docx`), Buffer.from(bytes));
    const xml = strFromU8(unzipSync(Uint8Array.from(bytes))['word/document.xml']);
    expect([...xml.matchAll(/ws:inlineOuterTop="(\d+)"/g)].map(match => Number(match[1]))).toEqual([0, 141]);
  }
  for (const bytes of result.native) {
    const doc = new HwpDocument(Uint8Array.from(bytes));
    try {
      expect(doc.pageCount()).toBe(1);
      const { controls } = JSON.parse(doc.getPageControlLayout(0));
      expect(controls.filter(c => c.type === 'table').map(c =>
        JSON.parse(doc.getTableProperties(c.secIdx, c.paraIdx, c.controlIdx)).outerTop)).toEqual([0, 141]);
      const { runs } = JSON.parse(doc.getPageTextLayout(0));
      const text = runs.map(run => run.text).join('').replace(/\s+/gu, '');
      expect(text).toBe('BeforeFirstgridBetweenSecondgridAfter');
    } finally { doc.free(); }
  }
});

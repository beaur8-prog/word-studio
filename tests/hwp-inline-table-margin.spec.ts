import { readFile, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { expect, test } from '@playwright/test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';

test('authored inline table margins survive two DOCX saves and both Hancom formats', async ({ page }, info) => {
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
    const grid = '<table><tr><td><p>A</p></td><td><p>B</p></td></tr>' +
      '<tr><td><p>C</p></td><td><p>D</p></td></tr></table>';
    original.pasteHtml(0, 0, 0, '<p>Before</p>' + table('First grid') + '<p>Between</p>' + table('Second grid') + grid + '<p>After</p>');
    const sourceGrid = JSON.parse(original.getPageControlLayout(0)).controls.find(c => c.type === 'table' && c.rowCount === 2);
    expect(sourceGrid).toBeDefined();
    expect(JSON.parse(original.insertText(sourceGrid.secIdx, sourceGrid.paraIdx, 0, '   ')).ok).toBe(true);
    const parts = unzipSync(original.exportHwpx()), section = 'Contents/section0.xml';
    const margins = [0, 141, 283];
    let index = 0;
    let xml = strFromU8(parts[section]).replace(/<hp:outMargin\b[^>]*\/>/g, tag => {
        const top = margins[index++];
        const changed = tag.replace(/\btop="\d+"/, 'top="' + top + '"');
        return index === 3 ? changed.replace(/\bbottom="\d+"/, 'bottom="0"') : changed;
      });
    expect(index).toBe(3);
    // Record an authored line shared by the leading spaces and inline object.
    const tableEnd = xml.lastIndexOf('</hp:tbl>');
    const tail = xml.slice(tableEnd);
    const lines = /^<\/hp:tbl><\/hp:run><hp:linesegarray>(<hp:lineseg\b[^>]*\/>)[\s\S]*?<\/hp:linesegarray>/.exec(tail);
    expect(lines).toBeTruthy();
    xml = xml.slice(0, tableEnd) + tail.replace(lines![0], '</hp:tbl></hp:run><hp:linesegarray>' + lines![1] + '</hp:linesegarray>');
    parts[section] = strToU8(xml);
    input = zipSync(parts);
    const authored = new HwpDocument(input);
    try {
      expect(JSON.parse(authored.getLineInfo(sourceGrid.secIdx, sourceGrid.paraIdx, 3)).lineCount).toBe(1);
    } finally { authored.free(); }
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
    expect([...xml.matchAll(/ws:inlineOuterTop="(\d+)"/g)].map(match => Number(match[1]))).toEqual([0, 141, 283]);
    expect([...xml.matchAll(/ws:inlineOuterBottom="(\d+)"/g)].map(match => Number(match[1]))).toEqual([0]);
    expect(xml).toMatch(/<w:t xml:space="preserve"> {3}<\/w:t>/);
  }
  for (const bytes of result.native) {
    const doc = new HwpDocument(Uint8Array.from(bytes));
    try {
      expect(doc.pageCount()).toBe(1);
      const { controls } = JSON.parse(doc.getPageControlLayout(0));
      const tables = controls.filter(c => c.type === 'table').map(c =>
        JSON.parse(doc.getTableProperties(c.secIdx, c.paraIdx, c.controlIdx)));
      expect(tables.map(table => table.outerTop)).toEqual([0, 141, 283]);
      expect(tables.map(table => table.outerBottom)).toEqual([283, 283, 0]);
      const { runs } = JSON.parse(doc.getPageTextLayout(0));
      const text = runs.map(run => run.text).join('').replace(/\s+/gu, '');
      expect(text).toBe('BeforeFirstgridBetweenSecondgridABCDAfter');
    } finally { doc.free(); }
  }
});

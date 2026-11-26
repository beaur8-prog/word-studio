import { test, expect } from '@playwright/test';
import { createRequire } from 'node:module';
import { readFile, writeFile } from 'node:fs/promises';
import { dirname } from 'node:path';
import { pathToFileURL } from 'node:url';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';

const require = createRequire(import.meta.url);
const section = 'Contents/section0.xml';

async function boxedDigitFixture(format: 'hwp' | 'hwpx') {
  const core = dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(core + '/rhwp.js').href);
  await init({ module_or_path: await readFile(core + '/rhwp_bg.wasm') });
  const document = HwpDocument.createEmpty();
  try {
    document.createBlankDocument();
    const paragraph = (text: string) => '<p><span style="font-family:Arial;font-size:15pt">' + text + '</span></p>';
    const nested = '<table><tr><td><table><tr><td>' + paragraph('NESTEDEND') + '</td></tr></table></td></tr></table>';
    document.pasteHtml(0, 0, 0, paragraph('BODYMIDEND') + nested + paragraph('GUARDEND'));
    const parts = unzipSync(document.exportHwpx());
    const source = strFromU8(parts[section]);
    const digit = (value: string, shape = 'SHAPE_RECTANGLE') =>
      '<hp:compose circleType="' + shape + '" charSz="-3" composeType="SPREAD" charPrCnt="1" composeText="' + value +
      '"><hp:charPr prIDRef="0"/></hp:compose>';
    const updated = source
      .replace('BODYMIDEND', 'BODY</hp:t>' + digit('1') + '<hp:t>MID</hp:t>' + digit('2') + '<hp:t>END')
      .replace('NESTEDEND', 'NESTED</hp:t>' + digit('3') + '<hp:t>END')
      .replace('GUARDEND', 'GUARD</hp:t>' + digit('4', 'SHAPE_CIRCLE') + '<hp:t>END');
    if (updated === source) throw Error('Synthetic source controls missing');
    parts[section] = strToU8(updated);
    const hwpx = zipSync(parts);
    if (format === 'hwpx') return hwpx;
    const converted = new HwpDocument(hwpx);
    try { return converted.exportHwp(); } finally { converted.free(); }
  } finally { document.free(); }
}

for (const format of ['hwp', 'hwpx'] as const) test(format + ' square digits survive nested cells and two DOCX saves', async ({ page }, info) => {
  const bytes = await boxedDigitFixture(format);
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const reader = new HwpReaderSession(), states = [], packages = [];
    try {
      const name = 'boxed.' + format;
      await reader.open(new File([new Uint8Array(bytes)], name));
      const editable = await reader.editable();
      let doc = importEditableHwp(editable, name);
      for (let save = 0; save <= 2; save++) {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const boxes = [...layout.element.querySelectorAll<HTMLElement>('[data-hwp-boxed-digit]')];
          states.push(boxes.map(span => {
            const rect = span.getBoundingClientRect(), digit = span.firstElementChild!;
            const ink = digit.getBoundingClientRect();
            return { value: span.textContent, parent: span.closest('p')!.textContent!.trim(),
              width: rect.width, height: rect.height, ink: ink.width, border: getComputedStyle(span).boxShadow };
          }));
        } finally { layout.dispose(); }
        if (save === 2) break;
        const file = await exportDocx(doc);
        packages.push([...new Uint8Array(await file.arrayBuffer())]);
        const opened = await importFile(new File([file], 'boxed.docx'));
        doc = { ...opened.document, content: parseDocumentHtml(opened.html!) };
      }
      return { states, packages };
    } finally { reader.close(); }
  }, { bytes: [...bytes], format });
  for (const [index, content] of result.packages.entries())
    await writeFile(info.outputPath('textbox-save-boxed-' + format + '-' + (index + 1) + '.docx'), Buffer.from(content));
  for (const boxes of result.states) {
    expect(boxes.map(item => [item.value, item.parent])).toEqual([
      ['1', 'BODY1MID2END'], ['2', 'BODY1MID2END'], ['3', 'NESTED3END'],
    ]);
    for (const box of boxes) {
      expect(box.width).toBeCloseTo(20);
      expect(box.height).toBeCloseTo(20);
      expect(box.ink).toBeGreaterThan(0);
      expect(box.border).not.toBe('none');
    }
  }
});

import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';

test('spaces sharing an HWP inline table line survive two DOCX saves without another text line', async ({ page }) => {
  const requireProject = createRequire(import.meta.url);
  const library = dirname(requireProject.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(join(library, 'rhwp.js')).href);
  await init({ module_or_path: await readFile(join(library, 'rhwp_bg.wasm')) });
  const doc = HwpDocument.createEmpty();
  let source: Uint8Array;
  try {
    doc.createBlankDocument();
    doc.pasteHtml(0, 0, 0, '<p>BODY BEFORE</p><table><tr><td><p>INLINE TABLE</p></td></tr></table><p>BODY AFTER</p>');
    const control = JSON.parse(doc.getControls()).find(item => item.ctrlId === 'tbl' && item.list === 0);
    doc.setTableProperties(0, control.para, control.controlIndex, JSON.stringify({ treatAsChar: true }));
    const files = unzipSync(doc.exportHwpx()), name = 'Contents/section0.xml';
    files[name] = strToU8(strFromU8(files[name]).replace('<hp:tbl ', '<hp:t>   </hp:t><hp:tbl '));
    source = zipSync(files);
    const reopen = new HwpDocument(source);
    try {
      expect(reopen.getTextRange(0, control.para, 0, 0xffffffff)).toBe('   ');
      expect(JSON.parse(reopen.getLineInfo(0, control.para, 0)).lineCount).toBe(1);
    } finally { reopen.free(); }
  } finally { doc.free(); }

  await page.goto('/');
  const result = await page.evaluate(async bytes => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const reader = new HwpReaderSession();
    let document;
    try {
      await reader.open(new File([new Uint8Array(bytes)], 'inline-spaces.hwpx'));
      document = importEditableHwp(await reader.editable(), 'inline-spaces.hwpx');
    } finally { reader.close(); }
    const first = document, stages = [];
    for (let i = 0; i < 3; i++) {
      const print = await createPrintLayout(generateHTML(document.content, editorExtensions), document.page);
      try {
        const paragraphs = [...print.element.querySelectorAll<HTMLElement>('.print-content > p')];
        const spaces = paragraphs.find(p => p.textContent === '   ');
        stages.push({ text: spaces?.textContent, table: spaces?.nextElementSibling?.textContent,
          height: spaces?.getBoundingClientRect().height, firstLayout: spaces?.hasAttribute('data-hwp-inline-table-spaces-layout') });
      } finally { print.dispose(); }
      if (i < 2) {
        const opened = await importFile(new File([await exportDocx(document)], 'inline-spaces.docx'));
        document = { ...opened.document, content: parseDocumentHtml(opened.html!) };
      }
    }
    return { first, stages };
  }, [...source]);
  for (const stage of result.stages) {
    expect(stage.firstLayout).toBe(true);
    expect(stage.text).toBe('   ');
    expect(stage.table).toContain('INLINE TABLE');
    expect(stage.height).toBeLessThan(.3);
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'inline-spaces.wordstudio', mimeType: 'application/json',
    buffer: Buffer.from(JSON.stringify(result.first)) });
  const carrier = page.getByRole('textbox', { name: '문서 편집 영역' }).locator('p[data-hwp-inline-table-spaces="true"]');
  await expect(carrier).toHaveAttribute('data-hwp-inline-table-spaces-layout', 'true');
  await carrier.fill('ADDED TEXT');
  await expect(carrier).toHaveText('ADDED TEXT');
  await expect(carrier).not.toHaveAttribute('data-hwp-inline-table-spaces-layout', 'true');
});

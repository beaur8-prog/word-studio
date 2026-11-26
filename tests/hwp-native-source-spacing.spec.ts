import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { expect, test } from '@playwright/test';

const bodyText = '한글 모양과 공백 확인';
const cellText = '한글 영어와 문서 확인';

test('native HWP saves retain authored tracking while preview uses fallback advances', async ({ page }) => {
  await page.goto('/');
  const installed = await page.evaluate(async () => {
    try {
      await new FontFace('Native HWP tracking witness', 'local("한양중고딕")').load();
      return true;
    } catch { return false; }
  });
  test.skip(installed, 'An installed source face does not use the fallback projection');
  const result = await page.evaluate(async ({ bodyText, cellText }) => {
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { exportEditableHwp } = await import('/src/hwp/editing.ts');
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { defaultPage } = await import('/src/types.ts');
    const paragraph = text => `<p data-hwp-source-spaces="true" style='font-family:"한양중고딕";font-size:12.5pt;line-height:30px'><span style='font-family:"한양중고딕";font-size:12.5pt;letter-spacing:-1.83333px'>${text}</span></p>`;
    const model = {
      format: 'word-studio' as const, version: 1 as const, title: 'Source tracking', id: 'source-tracking',
      updatedAt: new Date().toISOString(), comments: [], page: defaultPage,
      content: parseDocumentHtml(paragraph(bodyText) + `<table><tr><td colwidth="540">${paragraph(cellText)}</td></tr></table>`),
    };
    const layout = await createPrintLayout(generateHTML(model.content, editorExtensions), model.page, undefined, undefined,
      { paintHwpCellLines: false, hwpNativeExport: true });
    let projection;
    try {
      projection = [...layout.element.querySelectorAll('.print-content p')].map(p => ({
        text: p.textContent,
        runs: p.querySelectorAll('[data-hwp-source-advance-run]').length,
      }));
    } finally { layout.dispose(); }
    const outputs = [];
    for (const kind of ['hwp', 'hwpx'] as const) {
      const saved = await exportEditableHwp(model, kind);
      outputs.push([...new Uint8Array(await saved.blob.arrayBuffer())]);
    }
    return { projection, outputs };
  }, { bodyText, cellText });
  expect(result.projection.map(p => p.text)).toEqual([bodyText, cellText]);
  expect(result.projection.every(p => p.runs > 0)).toBe(true);
  const require = createRequire(import.meta.url), core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await readFile(path.join(core, 'rhwp_bg.wasm')) });
  for (const output of result.outputs) {
    const doc = new HwpDocument(Uint8Array.from(output));
    try {
      expect(doc.pageCount()).toBe(1);
      expect(doc.getTextRange(0, 0, 0, 100)).toBe(bodyText);
      expect(JSON.parse(doc.getTableDimensions(0,1,0)).cellCount).toBe(1);
      const body = Array.from(bodyText, (_, i) => JSON.parse(doc.getCharPropertiesAt(0,0,i)).spacings[0]);
      const cell = Array.from(cellText, (_, i) => JSON.parse(doc.getCellCharPropertiesAt(0,1,0,0,0,i)).spacings[0]);
      expect(body).toEqual(Array(bodyText.length).fill(-11));
      expect(cell).toEqual(Array(cellText.length).fill(-11));
    } finally { doc.free(); }
  }
});

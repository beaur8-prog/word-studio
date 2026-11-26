import { readFile } from 'node:fs/promises';
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

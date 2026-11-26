import { expect, test } from '@playwright/test';
import { readFile, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import JSZip from 'jszip';

const fixture = createRequire(import.meta.url)('./hwp-paragraph-border-box-fixture.cjs');
const text = ['Lead', 'First bordered line', 'Last bordered line', 'Plain break', 'Disconnected control', 'Incomplete begins', 'Incomplete ends'];
const layoutAttribute = 'data-paragraph-border-box-layout';
const startAttribute = 'data-paragraph-border-box-start';
const endAttribute = 'data-paragraph-border-box-end';

for (const format of ['hwp', 'hwpx'] as const) {
  test(format + ' connected paragraph borders survive editing and two standard DOCX saves', async ({ page }, info) => {
    const bytes = (await fixture())[format];
    await page.goto('/');
    const result = await page.evaluate(async ({ bytes, format, layoutAttribute, startAttribute, endAttribute }) => {
      const { HwpReaderSession } = await import('/src/hwp/reader.ts');
      const { importEditableHwp } = await import('/src/hwp/editing.ts');
      const { exportDocx, importFile } = await import('/src/document-io.ts');
      const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
      const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
      const { createPrintLayout } = await import('/src/print-layout.ts');
      const session = new HwpReaderSession();
      try {
        await session.open(new File([new Uint8Array(bytes)], 'border-box.' + format));
        let doc = importEditableHwp(await session.editable(), 'border-box.' + format);
        const first = structuredClone(doc), stages = [], saves = [];
        for (let i = 0; i < 3; i++) {
          const markup = generateHTML(doc.content, editorExtensions);
          const print = await createPrintLayout(markup, doc.page);
          const states = (root: HTMLElement) => [...root.querySelectorAll<HTMLElement>('.print-content p')].filter(p => p.textContent).map(p => {
            const box = p.closest('.print-page')!.getBoundingClientRect();
            const range = document.createRange(); range.selectNodeContents(p);
            const rect = range.getBoundingClientRect();
            return { text: p.textContent, layout: p.hasAttribute(layoutAttribute), start: p.hasAttribute(startAttribute), end: p.hasAttribute(endAttribute),
              bg: getComputedStyle(p).backgroundImage, glyph: [rect.left - box.left, rect.top - box.top, rect.width, rect.height] };
          });
          try {
            const noBorder = markup.replace(/ data-paragraph-border-box="[^"]*"/g, '');
            const control = await createPrintLayout(noBorder, doc.page);
            try { stages.push({ pages: print.pages, withBox: states(print.element), withoutBox: states(control.element) }); }
            finally { control.dispose(); }
          } finally { print.dispose(); }
          if (i < 2) {
            const packageBytes = await exportDocx(doc);
            saves.push([...new Uint8Array(await packageBytes.arrayBuffer())]);
            const loaded = await importFile(new File([packageBytes], 'border-save.docx'));
            doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
          }
        }
        return { first, stages, saves };
      } finally { session.close(); }
    }, { bytes: [...bytes], format, layoutAttribute, startAttribute, endAttribute });

    for (const stage of result.stages) {
      expect(stage.pages).toBe(1);
      expect(stage.withBox.map(p => p.text)).toEqual(text);
      expect(stage.withBox.map(p => [p.layout, p.start, p.end])).toEqual([
        [false, false, false], [true, true, false], [true, false, true], [false, false, false], [false, false, false],
        [false, false, false], [false, false, false],
      ]);
      expect(stage.withBox[1]?.bg).toContain('rgb(0, 112, 85)');
      expect(stage.withBox.map(p => p.glyph)).toEqual(stage.withoutBox.map(p => p.glyph));
    }
    for (const [i, bytes] of result.saves.entries()) {
      const zip = await JSZip.loadAsync(Uint8Array.from(bytes));
      const xml = await zip.file('word/document.xml')!.async('string');
      const boxes = [...xml.matchAll(/<w:pBdr>(.*?)<\/w:pBdr>/g)].map(m => m[1]);
      expect(boxes).toHaveLength(2);
      expect(boxes[0]).toBe(boxes[1]);
      for (const edge of ['top', 'left', 'right', 'bottom']) expect(boxes[0]).toMatch(new RegExp('<w:' + edge + '\\b[^>]*w:val="single"[^>]*/>'));
      expect(boxes[0]).toContain('<w:between w:val="nil"/>');
      await writeFile(info.outputPath('connected-box-' + format + '-' + (i + 1) + '.docx'), Uint8Array.from(bytes));
    }

    await page.locator('input[type=file]').first().setInputFiles({
      name: 'box.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.first)),
    });
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
    const leading = editor.locator('p[' + startAttribute + ']');
    const trailing = editor.locator('p[' + endAttribute + ']');
    await expect(leading).toHaveText(text[1]!);
    await expect(trailing).toHaveText(text[2]!);
    await leading.click();
    await page.keyboard.press('End');
    await page.keyboard.type(' edited');
    await expect(leading).toContainText(' edited');
    const mod = process.platform === 'darwin' ? 'Meta' : 'Control';
    await page.keyboard.press(mod + '+z');
    await expect(leading).toHaveText(text[1]!);
    await expect(trailing).toHaveText(text[2]!);
    await page.keyboard.press(mod + '+Shift+z');
    await expect(leading).toContainText(' edited');
  });
}

test('preserves a standard border wider than the available side margins without painting over text', async ({ page }, info) => {
  const { BorderStyle, Document, Packer, Paragraph } = await import('docx');
  const edge = { style: BorderStyle.SINGLE, size: 24, color: '007055', space: 0 };
  const document = new Document({ sections: [{ children: [new Paragraph({
    text: 'Wide border with insufficient side margins',
    indent: { left: 20, right: 20 },
    border: { top: edge, bottom: edge, left: { ...edge, space: 1 }, right: { ...edge, space: 1 } },
  })] }] });
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'wide-border.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: await Packer.toBuffer(document),
  });
  const paragraph = page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).locator('p').filter({ hasText: /^Wide border with insufficient side margins$/ });
  await expect(paragraph).toHaveAttribute('data-paragraph-border-box', /"size":24/);
  await expect(paragraph).not.toHaveAttribute(layoutAttribute);
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  const preview = page.locator('.print-layout.preview');
  await expect(preview.getByText('Wide border with insufficient side margins')).toBeVisible();
  await expect(preview.locator('[' + layoutAttribute + ']')).toHaveCount(0);
  await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
  await page.getByRole('button', { name: '파일', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Word 문서로 내보내기 DOCX', exact: true }).click();
  const file = info.outputPath('textbox-save-wide-border.docx');
  await (await pending).saveAs(file);
  const saved = await JSZip.loadAsync(await readFile(file));
  const xml = await saved.file('word/document.xml')!.async('string');
  expect(xml.match(/<w:pBdr>(.*?)<\/w:pBdr>/)?.[1]).toContain('w:sz="24"');
});

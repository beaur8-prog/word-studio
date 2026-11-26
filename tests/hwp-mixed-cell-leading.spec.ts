import { expect, test } from '@playwright/test';
import fs from 'node:fs/promises';
import { createRequire } from 'node:module';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const request = createRequire(import.meta.url);
const samples = new Map<string, { bytes: Uint8Array; height: number }>();
const text = 'WIDE small line wraps around natural word boundaries with different heights.';

test.beforeAll(async () => {
  const core = path.dirname(request.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  const doc = HwpDocument.createEmpty();
  try {
    doc.createBlankDocument();
    doc.pasteHtml(0, 0, 0, '<p>BEFORE</p><table style="width:320px"><tr>' +
      '<td style="width:320px;height:1px;padding:0;border:none"><p>TOP</p><p style="font-size:12pt">' +
      '<span style="font-size:12pt">WIDE </span><span style="font-size:11pt">' + text.slice(5) +
      '</span></p><p>END</p></td></tr></table><p>AFTER</p>');
    const table = JSON.parse(doc.getControls()).find((control: any) => control.ctrlId === 'tbl');
    doc.setTableProperties(0, table.para, table.controlIndex, JSON.stringify({ treatAsChar: true }));
    doc.applyParaFormatInCell(0, table.para, table.controlIndex, 0, 1, JSON.stringify({ lineSpacingType: 'Percent', lineSpacing: 160 }));
    for (const format of ['hwp', 'hwpx']) {
      const bytes = format === 'hwp' ? doc.exportHwp() : doc.exportHwpx();
      const reopened = new HwpDocument(bytes);
      try {
        const entry = JSON.parse(reopened.getControls()).find((control: any) => control.ctrlId === 'tbl');
        const reference = JSON.stringify([{ controlIndex: entry.controlIndex, cellIndex: 0, cellParaIndex: 1 }]);
        const { stored, lines } = JSON.parse(reopened.getParagraphLineMetrics(0, entry.para, reference));
        expect(stored).toBe(true);
        expect(lines).toHaveLength(2);
        expect(lines[0].height).toBeCloseTo(16);
        expect(lines[1].height).toBeCloseTo(11 * 4 / 3);
        const height = lines.reduce((sum: number, line: any) => sum + line.height + line.spacing, 0);
        samples.set(format, { bytes, height });
      } finally { reopened.free(); }
    }
  } finally { doc.free(); }
});

for (const format of ['hwp', 'hwpx']) {
  test(`${format}: an inline cell uses each line's own font for percentage spacing`, async ({ page }) => {
    const sample = samples.get(format)!;
    await page.goto('/');
    const result = await page.evaluate(async ({ format, bytes, text }) => {
      const { HwpReaderSession } = await import('/src/hwp/reader.ts');
      const { importEditableHwp } = await import('/src/hwp/editing.ts');
      const { exportDocx, importFile } = await import('/src/document-io.ts');
      const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
      const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
      const { createPrintLayout } = await import('/src/print-layout.ts');
      const session = new HwpReaderSession();
      const states = [];
      try {
        await session.open(new File([new Uint8Array(bytes)], 'cell.' + format));
        let doc = importEditableHwp(await session.editable(), 'cell.' + format);
        const first = doc;
        for (let save = 0; save < 3; save++) {
          const render = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
          try {
            const p = [...render.element.querySelectorAll<HTMLParagraphElement>('.print-content td p')]
              .find(paragraph => paragraph.textContent === text)!;
            states.push({ height: p.getBoundingClientRect().height,
              mode: p.dataset.hwpBodyLayout, text: p.textContent, previous: p.previousElementSibling?.textContent,
              next: p.nextElementSibling?.textContent });
          } finally { render.dispose(); }
          if (save < 2) {
            const next = await importFile(new File([await exportDocx(doc)], 'again.docx'));
            doc = { ...next.document, content: parseDocumentHtml(next.html!) };
          }
        }
        return { first, states };
      } finally { session.close(); }
    }, { format, bytes: [...sample.bytes], text });
    for (const state of result.states) {
      expect(state).toMatchObject({ mode: 'true', text, previous: 'TOP', next: 'END' });
      expect(state.height).toBeCloseTo(sample.height, 1);
    }
    await page.locator('input[type=file]').first().setInputFiles({
      name: 'cell.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.first)),
    });
    const paragraph = page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).locator('td p').filter({ hasText: 'WIDE' });
    await expect(paragraph).toHaveAttribute('data-hwp-body-layout', 'true');
    await expect.poll(() => paragraph.evaluate(p => p.getBoundingClientRect().height)).toBeCloseTo(sample.height, 1);
    await paragraph.click();
    await page.keyboard.press('End');
    await page.keyboard.insertText(' edit');
    await expect(paragraph).toContainText(' edit');
    await page.keyboard.press('ControlOrMeta+z');
    await expect(paragraph).toHaveText(text);
    await expect.poll(() => paragraph.evaluate(p => p.getBoundingClientRect().height)).toBeCloseTo(sample.height, 1);
  });
}

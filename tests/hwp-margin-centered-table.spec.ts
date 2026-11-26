import { expect, test } from '@playwright/test';
import { writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import JSZip from 'jszip';

const fixture = createRequire(import.meta.url)('./hwp-margin-centered-table-fixture.cjs');

for (const format of ['hwp', 'hwpx'] as const) test(format + ' first table keeps its authored center and paper top through two DOCX saves', async ({ page }, info) => {
  const generated = await fixture();
  const original = generated[format + ':Center'];
  const negative = generated[format + ':Left'];
  await page.goto('/');
  const audit = await page.evaluate(async ({ format, source, control }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const hwp = async (bytes: number[]) => {
      const session = new HwpReaderSession();
      const label = 'margin-centered.' + format;
      try {
        await session.open(new File([Uint8Array.from(bytes)], label));
        return importEditableHwp(await session.editable(), label);
      } finally { session.close(); }
    };
    let doc = await hwp(source);
    const before = structuredClone(doc), wrong = await hwp(control);
    const negative = generateHTML(wrong.content, editorExtensions).includes('horizontalAlign');
    const states = [], files = [];
    for (let stage = 0; stage < 3; stage++) {
      const markup = generateHTML(doc.content, editorExtensions);
      const snapshot = async (settings: typeof doc.page, html = markup) => {
        const print = await createPrintLayout(html, settings);
        try {
          const paper = print.element.querySelector<HTMLElement>('.print-page')!;
          const table = paper.querySelector<HTMLTableElement>('.print-content table')!;
          const p = paper.getBoundingClientRect(), r = table.getBoundingClientRect();
          return { pages: print.pages, x: r.x - p.x, y: r.y - p.y, width: r.width, center: r.x + r.width / 2 - p.x, text: paper.querySelector('.print-content')!.textContent,
            anchor: JSON.parse(table.dataset.tableAnchor || 'null') };
        } finally { print.dispose(); }
      };
      const normal = await snapshot(doc.page);
      // Moving both margins in the same direction translates the live center
      // without changing the available width or the absolute paper top.
      const delta = 24 * 25.4 / 96;
      const moved = await snapshot({ ...doc.page, margins: { ...doc.page.margins!, left: doc.page.margins!.left + delta, right: doc.page.margins!.right - delta } });
      const resizedDoc = structuredClone(doc);
      const firstTable = resizedDoc.content.content!.find((node: any) => node.type === 'table');
      firstTable!.content![0].content![0].attrs!.colwidth = [430];
      const resized = await snapshot(doc.page, generateHTML(resizedDoc.content, editorExtensions));
      states.push({ normal, moved, resized });
      if (stage < 2) {
        const bytes = await exportDocx(doc);
        files.push([...new Uint8Array(await bytes.arrayBuffer())]);
        const loaded = await importFile(new File([bytes], 'margin-centered.docx'));
        doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
      }
    }
    return { before, states, files, negative };
  }, { format, source: [...original.bytes], control: [...negative.bytes] });

  expect(audit.negative).toBe(false);
  for (const { normal, moved, resized } of audit.states) {
    expect(normal.pages).toBe(1);
    expect(normal.anchor).toMatchObject({ x: 0, horizontal: 'margin', horizontalAlign: 'center' });
    // Chromium's bounding box includes the collapsed outer border; use
    // the actual midpoint for centering and separately check native placement.
    expect(normal.center).toBeCloseTo((75 + 600 - 45) / 2, 1);
    expect(Math.abs(normal.x - original.place.x)).toBeLessThan(1);
    expect(normal.y).toBeCloseTo(original.place.y, 0);
    expect(normal.text).toBe('Math table titleBody follows this.');
    expect(moved.x - normal.x).toBeCloseTo(24, 0);
    expect(moved.y).toBeCloseTo(normal.y, 1);
    expect(moved.text).toBe(normal.text);
    expect(resized.width - normal.width).toBeCloseTo(20, 0);
    expect(resized.center).toBeCloseTo(normal.center, 1);
    expect(resized.y).toBeCloseTo(normal.y, 1);
    expect(resized.text).toBe(normal.text);
  }
  for (const [i, bytes] of audit.files.entries()) {
    const zip = await JSZip.loadAsync(Uint8Array.from(bytes));
    const xml = await zip.file('word/document.xml')!.async('string');
    const anchor = xml.match(/<w:tblpPr\b[^>]*\/>/)?.[0] ?? '';
    expect(anchor).toContain('w:horzAnchor="margin"');
    expect(anchor).toContain('w:tblpXSpec="center"');
    expect(anchor).toContain('w:vertAnchor="page"');
    expect(anchor).not.toMatch(/\bw:tblpX=/);
    await writeFile(info.outputPath('textbox-save-center-' + format + '-' + (i + 1) + '.docx'), Uint8Array.from(bytes));
  }
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'center.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(audit.before)),
  });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  const title = editor.locator('td p');
  await expect(title).toHaveText('Math table title');
  await expect.poll(() => title.evaluate(p => {
    const table = p.closest('table')!, paper = table.closest('.document-stack')!;
    const r = table.getBoundingClientRect(), sheet = paper.getBoundingClientRect();
    return (r.left + r.width / 2 - sheet.left) / (sheet.width / 600);
  })).toBeCloseTo((75 + 600 - 45) / 2, 1);
  await title.fill('Editable heading');
  await expect(title).toHaveText('Editable heading');
  await page.keyboard.press('ControlOrMeta+z');
  await expect(title).toHaveText('Math table title');
});

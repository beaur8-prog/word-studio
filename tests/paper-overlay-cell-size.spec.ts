import { expect, test, type Page } from '@playwright/test';
import { createRequire } from 'node:module';
import { writeFile } from 'node:fs/promises';
import { Document, HorizontalPositionRelativeFrom, ImageRun, Packer, Paragraph, Table, TableCell, TableLayoutType,
  TableRow, TextWrappingType, VerticalPositionRelativeFrom, WidthType } from 'docx';
import JSZip from 'jszip';

const createHwpFixture = createRequire(import.meta.url)('./hwp-paper-overlay-cell-fixture.cjs');
const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==', 'base64');
const controls = [
  { name: 'paper-foreground', layoutInCell: false, page: true, wrap: TextWrappingType.NONE, behindDocument: false },
  { name: 'cell-foreground', layoutInCell: true, page: true, wrap: TextWrappingType.NONE, behindDocument: false },
  { name: 'paper-behind', layoutInCell: false, page: true, wrap: TextWrappingType.NONE, behindDocument: true },
  { name: 'paper-wrapped', layoutInCell: false, page: true, wrap: TextWrappingType.TOP_AND_BOTTOM, behindDocument: false },
  { name: 'margin-foreground', layoutInCell: false, page: false, wrap: TextWrappingType.NONE, behindDocument: false },
] as const;

async function wordFixture(control: typeof controls[number]) {
  const picture = new ImageRun({ type: 'png', data: png, transformation: { width: 96, height: 80 },
    altText: { name: 'stamp', title: 'stamp', description: 'stamp' }, floating: {
      horizontalPosition: { relative: control.page ? HorizontalPositionRelativeFrom.PAGE : HorizontalPositionRelativeFrom.MARGIN,
        offset: 200 * 9525 },
      verticalPosition: { relative: VerticalPositionRelativeFrom.PAGE, offset: 350 * 9525 },
      layoutInCell: control.layoutInCell, behindDocument: control.behindDocument, wrap: { type: control.wrap },
    } });
  const cell = new TableCell({ width: { type: WidthType.DXA, size: 720 }, children: [new Paragraph({ children: [picture] })] });
  const table = new Table({ layout: TableLayoutType.FIXED, width: { type: WidthType.DXA, size: 720 }, columnWidths: [720],
    rows: [new TableRow({ children: [cell] })] });
  return await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph('BEFORE TABLE'), table, new Paragraph('AFTER TABLE')] }] }));
}

async function inspect(page: Page, name: string, bytes: Uint8Array, alt: string) {
  await page.goto('/');
  return await page.evaluate(async ({ name, bytes, alt }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const reopen = async (file: File) => {
      const source = await importFile(file);
      return { ...source.document, content: parseDocumentHtml(source.html!) };
    };
    const packet = new File([Uint8Array.from(bytes)], name);
    const entry = async () => {
      if (name.endsWith('.docx')) return reopen(packet);
      const reader = new HwpReaderSession();
      try { await reader.open(packet); return importEditableHwp(await reader.editable(), name); }
      finally { reader.close(); }
    };
    let doc = await entry();
    const stages = [], files = [];
    for (let index = 0; index < 3; index++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try {
        const image = [...layout.element.querySelectorAll<HTMLImageElement>('img')].find(img => img.alt === alt)!;
        const paper = image.closest('.print-page')!, cell = image.closest('td,th')!;
        const origin = paper.getBoundingClientRect(), box = image.getBoundingClientRect(), css = getComputedStyle(image);
        stages.push({ doc, pages: layout.pages, text: paper.querySelector('.print-content')!.textContent,
          x: box.x - origin.x, y: box.y - origin.y, width: box.width, height: box.height,
          cellWidth: cell.getBoundingClientRect().width, anchorWidth: image.offsetParent!.getBoundingClientRect().width,
          maxWidth: css.maxWidth, declared: [Number(image.getAttribute('width')), Number(image.getAttribute('height'))],
          anchor: JSON.parse(image.dataset.imageFloating!) });
      } finally { layout.dispose(); }
      if (index < 2) {
        const blob = await exportDocx(doc);
        files.push([...new Uint8Array(await blob.arrayBuffer())]);
        doc = await reopen(new File([blob], 'paper-overlay.docx'));
      }
    }
    return { stages, files };
  }, { name, bytes: [...bytes], alt });
}

async function inspectEditor(page: Page, doc: object, alt: string, fullWidth: boolean) {
  await page.locator('input[type=file]').first().setInputFiles({ name: 'paper-overlay.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(doc)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  const image = editor.getByAltText(alt, { exact: true });
  await expect(image).toBeVisible();
  await expect.poll(() => image.evaluate(img => getComputedStyle(img).maxWidth)).toBe(fullWidth ? 'none' : '100%');
  await expect.poll(() => image.evaluate((img, full) => {
    const cell = img.closest('td,th')!, rendered = img.getBoundingClientRect().width;
    const cellWidth = cell.getBoundingClientRect().width;
    return full ? rendered > cellWidth + 15 : rendered <= cellWidth + 1;
  }, fullWidth)).toBe(true);
}

async function checkSaves(files: number[][], label: string, layoutInCell: boolean, wrap: number) {
  for (const [index, file] of files.entries()) {
    const bytes = Buffer.from(file);
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file('word/document.xml')!.async('string');
    expect(xml).toMatch(/<w:tc(?=[ >])[\s\S]*?<wp:anchor\b/);
    expect(xml).toContain(`layoutInCell="${layoutInCell ? '1' : '0'}"`);
    expect(xml).toMatch(wrap === TextWrappingType.NONE ? /<wp:wrapNone\b/ : /<wp:wrapTopAndBottom\b/);
    if (!layoutInCell && label !== 'margin-foreground') expect(xml.match(/relativeFrom="page"/g)).toHaveLength(2);
    await writeFile(test.info().outputPath(`textbox-save-overlay-${label}-${index + 1}.docx`), bytes);
  }
}

for (const format of ['hwp', 'hwpx'] as const) {
  test(`${format} paper overlay keeps its authored bounds outside its narrow anchor cell`, async ({ page }) => {
    const native = (await createHwpFixture())[format];
    const loaded = await inspect(page, 'overlay.' + format, native.bytes, '타원');
    for (const stage of loaded.stages) {
      expect(stage.pages).toBe(1);
      expect(stage.anchor).toMatchObject({ layoutInCell: false, behindDocument: false, wrap: { type: 0 },
        horizontalPosition: { relative: 'page' }, verticalPosition: { relative: 'page' } });
      expect(stage.maxWidth).toBe('none');
      expect(stage.width).toBeCloseTo(stage.declared[0], 1);
      expect(stage.height).toBeCloseTo(stage.declared[1], 1);
      expect(stage.x + (stage.width - native.bounds.width) / 2).toBeCloseTo(native.bounds.x, 1);
      expect(stage.y + (stage.height - native.bounds.height) / 2).toBeCloseTo(native.bounds.y, 1);
      expect(stage.width).toBeGreaterThan(stage.cellWidth + 15);
      expect(stage.text).toBe('TABLE CELLTAILAFTER TABLE');
    }
    await checkSaves(loaded.files, format, false, TextWrappingType.NONE);
    await inspectEditor(page, loaded.stages[2].doc, '타원', true);
  });
}

for (const control of controls) {
  test(`standard DOCX ${control.name} follows its own cell width policy through two saves`, async ({ page }) => {
    const loaded = await inspect(page, `${control.name}.docx`, await wordFixture(control), 'stamp');
    const fullWidth = control.name === 'paper-foreground';
    for (const stage of loaded.stages) {
      expect(stage.pages).toBe(1);
      expect(stage.maxWidth).toBe(fullWidth ? 'none' : '100%');
      expect(stage.width).toBeCloseTo(fullWidth ? 96 : stage.anchorWidth, 1);
      expect(stage.width / stage.height).toBeCloseTo(96 / 80, 2);
      expect(stage.text).toBe('BEFORE TABLEAFTER TABLE');
      if (fullWidth) { expect(stage.x).toBeCloseTo(200, 1); expect(stage.y).toBeCloseTo(350, 1); }
    }
    await checkSaves(loaded.files, control.name, control.layoutInCell, control.wrap);
    await inspectEditor(page, loaded.stages[2].doc, 'stamp', fullWidth);
  });
}

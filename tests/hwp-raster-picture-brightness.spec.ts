import { expect, test } from '@playwright/test';
import { Document, ImageRun, Packer, Paragraph, Table, TableCell, TableRow } from 'docx';

test('an externally moved HWP raster fill keeps its luminance on both DOCX saves', async ({ page }) => {
  const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==', 'base64');
  const picture = new ImageRun({
    type: 'png', data: png, transformation: { width: 48, height: 24 },
    floating: {
      horizontalPosition: { relative: 'column', offset: 190500 },
      verticalPosition: { relative: 'paragraph', offset: 0 },
      wrap: { type: 0 }, layoutInCell: true, behindDocument: false,
    },
  });
  const table = new Table({ rows: [new TableRow({ children: [
    new TableCell({ children: [new Paragraph({ children: [picture] }), new Paragraph('Editable cell text')] }),
  ] })] });
  const bytes = [...await Packer.toBuffer(new Document({ sections: [{ children: [table] }] }))];
  await page.goto('/');
  const { states, control } = await page.evaluate(async bytes => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { hwpCellRasterFillNamespace: sourceNamespace } = await import('/src/hwp-cell-raster-fill.ts');
    const { unzipSync, zipSync } = await import('/node_modules/fflate/esm/browser.js');
    const drawingNamespace = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const anchorNamespace = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
    const file = 'word/document.xml';
    const packageXml = (raw: Uint8Array) => {
      const files = unzipSync(raw);
      const xml = new DOMParser().parseFromString(new TextDecoder().decode(files[file]), 'application/xml');
      return { files, xml, anchor: xml.getElementsByTagNameNS(anchorNamespace, 'anchor')[0] };
    };
    const original = packageXml(new Uint8Array(bytes));
    const blip = original.anchor.getElementsByTagNameNS(drawingNamespace, 'blip')[0];
    const lum = original.xml.createElementNS(drawingNamespace, 'a:lum');
    lum.setAttribute('bright', '8000');
    lum.setAttribute('contrast', '0');
    blip.append(lum);
    const packed = () => {
      original.files[file] = new TextEncoder().encode(new XMLSerializer().serializeToString(original.xml));
      return zipSync(original.files);
    };
    const read = async (raw: Uint8Array) => {
      const loaded = await importFile(new File([raw], 'moved-raster.docx'));
      const doc = { ...loaded.document, content: parseDocumentHtml(loaded.html) };
      const root = document.createElement('div');
      root.innerHTML = generateHTML(doc.content, editorExtensions);
      const image = root.querySelector<HTMLImageElement>('td img[data-image-floating]')!;
      return { doc, image, text: root.textContent };
    };
    const { image: unmarked } = await read(packed());
    original.anchor.setAttributeNS('http://www.w3.org/2000/xmlns/', 'xmlns:wscf', sourceNamespace);
    original.anchor.setAttributeNS(sourceNamespace, 'wscf:fill', JSON.stringify('457200:228600:0:0'));
    let raw = packed();
    const states = [];
    for (const stage of [0, 1, 2]) {
      const { doc, image, text } = await read(raw);
      const { anchor } = packageXml(raw);
      const floating = JSON.parse(image.dataset.imageFloating!);
      states.push({
        text, brightness: floating.hwpRasterPictureBrightness,
        offset: floating.horizontalPosition.offset, width: image.width, height: image.height,
        effect: image.style.filter.startsWith('url('),
        xmlBrightness: anchor.getElementsByTagNameNS(drawingNamespace, 'lum')[0]?.getAttribute('bright'),
        ordinaryPicture: anchor.getAttributeNS(sourceNamespace, 'picture'),
      });
      if (stage < 2) raw = new Uint8Array(await (await exportDocx(doc)).arrayBuffer());
    }
    return { states, control: unmarked.style.filter };
  }, bytes);
  expect(control).toBe('');
  expect(states).toEqual([null, 'true', 'true'].map(ordinaryPicture => ({
    text: 'Editable cell text', brightness: 8, offset: 190500, width: 48, height: 24,
    effect: true, xmlBrightness: '8000', ordinaryPicture,
  })));
});

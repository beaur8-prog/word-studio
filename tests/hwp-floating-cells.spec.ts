import { test, expect } from '@playwright/test';

for (const format of ['hwp', 'hwpx'] as const) for (const location of ['cell', 'nested']) test(`${format}, ${location}: printed floating picture positions survive native export and two reopenings`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async ({format, location}) => {
    const { defaultPage } = await import('/src/types.ts'), { importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const { exportEditableHwp, importEditableHwp } = await import('/src/hwp/editing.ts'), { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';
    const picture = (i: number, type: number) => {
      const geometry = { horizontalPosition: { relative: 'column', offset: 20 * 9525 }, verticalPosition: { relative: 'paragraph', offset: 20 * 9525 }, wrap: { type }, behindDocument: !type,
        margins: { left: 3 * 9525, right: 4 * 9525, top: 5 * 9525, bottom: 6 * 9525 } };
      return `<img data-inline-image="true" data-image-floating="${JSON.stringify(geometry).replaceAll('"', '&quot;')}" src="data:image/png;base64,${png}" width="80" height="40" alt="Picture ${i}">`;
    };
    const cell = (body: string) => `<table style="width:400px"><tr style="height:150px"><td>${location === 'nested' ? `<table style="width:300px"><tr><td>${body}</td></tr></table>` : body}</td></tr></table>`;
    const loaded = await importFile(new File([`${cell(`<p style="font-size:10pt;line-height:20px;margin:0">FIRST${picture(0, 0)}</p><p>Body one</p>`)}<div data-page-break="true"></div>${cell(`<p style="font-size:10pt;line-height:20px;margin:0">SECOND${picture(1, 1)}</p><p>Body two</p>`)}`], 'float.html'));
    let doc = { ...loaded.document, page: { ...defaultPage, pageNumbers: false }, content: parseDocumentHtml(loaded.html!) };
    const original = doc, states = [];
    for (let save = 0; save < 3; save++) {
      const print = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try { states.push({ pages: print.pages, text: [...print.element.querySelectorAll('.print-content')].map(p => p.textContent!.replace(/\s/g, '')), pictures: [...print.element.querySelectorAll<HTMLImageElement>('.print-content img')].map(img => {
        const box = img.getBoundingClientRect(), paper = img.closest<HTMLElement>('.print-page')!, origin = paper.getBoundingClientRect();
        return { page: paper.dataset.page, x: box.left - origin.left, y: box.top - origin.top, width: box.width, height: box.height, alt: img.alt, geometry: JSON.parse(img.dataset.imageFloating || 'null') };
      }) }); } finally { print.dispose(); }
      if (save < 2) {
        const out = await exportEditableHwp(doc, format), session = new HwpReaderSession();
        try { await session.open(new File([out.blob], 'saved.' + format)); doc = importEditableHwp(await session.editable(), 'saved.' + format); }
        finally { session.close(); }
      }
    }
    return { states, original };
  }, {format, location});
  const { states } = result;
  for (const state of states) {
    expect(state.pages).toBe(2); expect(state.text).toEqual(states[0].text); expect(state.pictures).toHaveLength(2);
    state.pictures.forEach((img, i) => {
      const original = states[0].pictures[i];
      expect(img.page).toBe(original.page); expect(img.alt).toBe(original.alt);
      for (const property of ['x', 'y', 'width', 'height'] as const) expect(Math.abs(img[property] - original[property]), `${property} of ${img.alt}`).toBeLessThan(.2);
      expect(img.geometry.wrap.type).toBe(original.geometry.wrap.type); expect(img.geometry.margins).toEqual(original.geometry.margins);
    });
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'floating-cells.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.original)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => editor.evaluate((root, {zoom, expected}) => {
      const sheet = root.closest<HTMLElement>('.document-stack')!, origin = sheet.getBoundingClientRect(), scale = zoom / 100;
      const stride = Number(sheet.dataset.pageHeight) + 24, images = [...root.querySelectorAll<HTMLImageElement>('img[data-image-floating]')];
      if (images.length !== expected.length) return Infinity;
      return Math.max(...images.map((img, i) => {
        const box = img.getBoundingClientRect(), source = expected[i];
        return Math.max(Math.abs((box.left - origin.left) / scale - source.x),
          Math.abs((box.top - origin.top) / scale - (Number(source.page) - 1) * stride - source.y));
      }));
    }, {zoom, expected: states[0].pictures})).toBeLessThan(.5);
  }

});

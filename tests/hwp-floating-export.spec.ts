import { test, expect } from '@playwright/test';

for (const format of ['hwp', 'hwpx'] as const) test(`${format}: printed floating picture positions survive native export and two reopenings`, async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async format => {
    const { defaultPage } = await import('/src/types.ts'), { importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const { exportEditableHwp, importEditableHwp } = await import('/src/hwp/editing.ts'), { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';
    const picture = (i: number, type: number) => {
      const geometry = { horizontalPosition: { relative: 'margin', align: 'center' }, verticalPosition: { relative: 'paragraph', offset: 20 * 9525 }, wrap: { type }, behindDocument: !type,
        margins: { left: 3 * 9525, right: 4 * 9525, top: 5 * 9525, bottom: 6 * 9525 } };
      return `<img data-inline-image="true" data-image-floating="${JSON.stringify(geometry).replaceAll('"', '&quot;')}" src="data:image/png;base64,${png}" width="80" height="40" alt="Picture ${i}">`;
    };
    const loaded = await importFile(new File([`<p style="font-size:10pt;line-height:20px;margin:0">FIRST${picture(0, 0)}</p><p>Body one</p><div data-page-break="true"></div><p style="font-size:10pt;line-height:20px;margin:0">SECOND${picture(1, 1)}</p><p>Body two</p>`], 'float.html'));
    let doc = { ...loaded.document, page: { ...defaultPage, pageNumbers: false }, content: parseDocumentHtml(loaded.html!) };
    const states = [];
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
    return states;
  }, format);
  for (const state of states) {
    expect(state.pages).toBe(2); expect(state.text).toEqual(states[0].text); expect(state.pictures).toHaveLength(2);
    state.pictures.forEach((img, i) => {
      const original = states[0].pictures[i];
      expect(img.page).toBe(original.page); expect(img.alt).toBe(original.alt);
      for (const property of ['x', 'y', 'width', 'height'] as const) expect(Math.abs(img[property] - original[property])).toBeLessThan(.2);
      expect(img.geometry.wrap.type).toBe(original.geometry.wrap.type); expect(img.geometry.margins).toEqual(original.geometry.margins);
    });
  }
});

for (const format of ['hwp', 'hwpx'] as const) test(`${format} retains the anchor of a picture outside its paper`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async format => {
    const { defaultPage } = await import('/src/types.ts'), { importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { exportEditableHwp } = await import('/src/hwp/editing.ts'), { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';
    const pos = {
      horizontalPosition: { relative: 'page', offset: 120 * 9525 },
      verticalPosition: { relative: 'page', offset: 1500 * 9525 },
      wrap: { type: 1 }, layoutInCell: false,
    };
    const loaded = await importFile(new File([`<p>Body before<img data-inline-image="true" src="data:image/png;base64,${png}"
      width="80" height="40" data-image-floating="${JSON.stringify(pos).replaceAll('"', '&quot;')}" alt="Off-paper original"/></p><p>Body after</p>`], 'off-paper.html'));
    const doc = { ...loaded.document, page: { ...defaultPage, pageNumbers: false }, content: parseDocumentHtml(loaded.html!) };
    const exported = await exportEditableHwp(doc, format), session = new HwpReaderSession();
    try {
      const file = new File([exported.blob], 'off-paper.' + format);
      const info = await session.open(file);
      const { html } = await session.editable();
      const root = new DOMParser().parseFromString(html, 'text/html');
      const picture = root.querySelector('img');
      return { pages: info.count, text: root.body.textContent!.replace(/\s+/g, ' ').trim(),
        size: [Number(picture?.getAttribute('width')), Number(picture?.getAttribute('height'))],
        hasPicture: !!picture?.getAttribute('src'),
        geometry: JSON.parse(picture?.getAttribute('data-image-floating') || 'null') };
    } finally { session.close(); }
  }, format);
  expect(result.pages).toBe(1);
  expect(result.text).toBe('Body before Body after');
  expect(result.hasPicture).toBe(true);
  expect(result.size).toEqual([80, 40]);
  expect(result.geometry.horizontalPosition).toEqual({ relative: 'page', offset: 120 * 9525 });
  expect(result.geometry.verticalPosition).toEqual({ relative: 'page', offset: 1500 * 9525 });
});

test('HWPX picture staging binds grouped controls by paragraph and exact raster bytes', async ({ page }) => {
  await page.goto('/');
  const results = await page.evaluate(async () => {
    const { default: init, HwpDocument } = await import('/node_modules/.vite/deps/@rhwp_core.js');
    const { default: wasm } = await import('/node_modules/@rhwp/core/rhwp_bg.wasm?url');
    const { strFromU8, strToU8, unzipSync, zipSync } = await import('/node_modules/.vite/deps/fflate.js');
    const { stageHwpPictures } = await import('/src/hwp/export-pictures.ts');
    await init({ module_or_path: fetch(wasm).then(response => response.arrayBuffer()) });

    const canvas = document.createElement('canvas');
    canvas.width = canvas.height = 6;
    const paint = canvas.getContext('2d')!;
    const colors = ['red', 'green', 'blue', 'purple'];
    const payloads = colors.map(color => {
      paint.fillStyle = color;
      paint.fillRect(0, 0, 6, 6);
      return canvas.toDataURL().split(',')[1];
    });
    const pictures = colors.map((color, i) =>
      '<img src="data:image/png;base64,' + payloads[i] + '" width="18" height="18" alt="' + color + '">');
    const staged = stageHwpPictures('<p>Body' + pictures.slice(0, 2).join('') +
      '</p><table><tbody><tr><td><p>Cell' + pictures.slice(2).join('') + '</p></td></tr></tbody></table>');
    const seed = HwpDocument.createEmpty();
    let parts: Record<string, Uint8Array>;
    try {
      seed.createBlankDocument();
      seed.pasteHtml(0, 0, 0, '<p></p>' + staged.html + '<p></p>');
      parts = unzipSync(seed.exportHwpx());
    } finally { seed.free(); }
    const name = 'Contents/section0.xml', parse = (xml: string) => new DOMParser().parseFromString(xml, 'application/xml');
    const section = parse(strFromU8(parts[name]));
    const tags = (root: Document | Element, name: string) => Array.from(root.getElementsByTagName(name));
    const tokens = /WORDSTUDIO-PICTURE-[\da-f-]+-\d+-END/g;
    let groups = 0;
    for (const paragraph of tags(section, 'hp:p')) {
      const runs = [...paragraph.children].filter(child => child.tagName === 'hp:run');
      const text = runs.flatMap(run => [...run.children].filter(el => el.tagName === 'hp:t'));
      const markers = text.flatMap(el => [...(el.textContent ?? '').matchAll(tokens)].map(match => match[0]));
      if (markers.length !== 2) continue;
      for (const el of text) el.textContent = el.textContent!.replace(tokens, '');
      const first = runs.find(run => [...run.children].some(el => el.tagName === 'hp:pic'))!;
      const before = section.createElementNS(first.namespaceURI, 'hp:run');
      before.setAttribute('charPrIDRef', first.getAttribute('charPrIDRef')!);
      const group = section.createElementNS(first.namespaceURI, 'hp:t');
      group.textContent = markers.join(' ');
      before.append(group);
      first.before(before);
      groups++;
    }
    const save = () => {
      parts[name] = strToU8(new XMLSerializer().serializeToString(section));
      return staged.restore(HwpDocument, { exportHwpx: () => zipSync(parts) } as InstanceType<typeof HwpDocument>);
    };
    const restored = save();
    let associations;
    try {
      const archive = unzipSync(restored.exportHwpx());
      const manifest = parse(strFromU8(archive['Contents/content.hpf']));
      const assets = new Map(tags(manifest, 'opf:item').map(item => [item.getAttribute('id'), item.getAttribute('href')]));
      associations = tags(parse(strFromU8(archive[name])), 'hp:pic').map(pic => {
        const ref = tags(pic, 'hc:img')[0].getAttribute('binaryItemIDRef');
        const bytes = archive[assets.get(ref)!];
        return { source: payloads.indexOf(btoa(String.fromCharCode(...bytes))),
          description: tags(pic, 'hp:shapeComment')[0]?.textContent };
      });
    } finally { restored.free(); }
    const [first, second] = tags(section, 'hc:img');
    const id = first.getAttribute('binaryItemIDRef')!;
    first.setAttribute('binaryItemIDRef', second.getAttribute('binaryItemIDRef')!);
    second.setAttribute('binaryItemIDRef', id);
    let rejectedSwap = false;
    try { save().free(); }
    catch (error) { rejectedSwap = String(error).includes('그림을 한글로 저장하지 못했습니다'); }
    return { groups, associations, rejectedSwap };
  });
  expect(results).toEqual({ groups: 2, associations: [
    { source: 0, description: 'red' }, { source: 1, description: 'green' },
    { source: 2, description: 'blue' }, { source: 3, description: 'purple' },
  ], rejectedSwap: true });
});

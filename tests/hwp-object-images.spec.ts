import { test, expect } from '@playwright/test';
import { createRequire } from 'node:module';
import fs from 'node:fs/promises';
const fixture = createRequire(import.meta.url)('./hwp-object-image-fixture.cjs');
let pending: Promise<any>, pure: Promise<any>;
const samples = (edge: boolean) => edge ? pure ??= fixture({ pureImage: true, pageHide: true }) : pending ??= fixture();
for (const format of ['hwp', 'hwpx']) for (const edge of [false, true]) test(`${format}: ${edge ? 'an image-only group at the page edge' : 'a picture and dotted TOC inside one group'} survives the worker and two DOCX saves`, async ({ page }, info) => {
  const sample = (await samples(edge))[format];
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format, edge, groupHeight }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const reader = new HwpReaderSession(), states = [], files = [];
    try {
      await reader.open(new File([new Uint8Array(bytes)], 'object.' + format));
      const source = await reader.editable(); let doc = importEditableHwp(source, 'object.' + format), controlPages = 0;
      if (edge) {
        const margins = doc.page.margins ?? { top: doc.page.margin, bottom: doc.page.margin };
        const available = (297 - margins.top - margins.bottom) * 96 / 25.4;
        const before = Number.parseFloat(doc.content.content[0].attrs.lineSpacing);
        if (!(groupHeight > 0 && before > 0)) throw Error('Missing source line dimensions');
        doc.content.content.unshift({ type: 'paragraph', attrs: { lineSpacing: `${available - groupHeight - before - 1.5}px`, spaceBefore: '0px', spaceAfter: '0px', paragraphSize: '10pt' }, content: [{ type: 'text', text: 'PAGE EDGE RESERVE' }] });
        const control = structuredClone(doc.content);
        for (const p of control.content) if (p.attrs?.wordTextAlignment === 'top') p.attrs.wordTextAlignment = 'baseline';
        const layout = await createPrintLayout(generateHTML(control, editorExtensions), doc.page);
        try { controlPages = layout.pages; } finally { layout.dispose(); }
      }
      for (let n = 0; n < 3; n++) {
        const html = generateHTML(doc.content, editorExtensions), root = document.createElement('div'); root.innerHTML = html;
        const images = await Promise.all([...root.querySelectorAll('img')].map(async img => {
          await img.decode(); const canvas = document.createElement('canvas'); canvas.width = img.naturalWidth; canvas.height = img.naturalHeight;
          const ctx = canvas.getContext('2d')!; ctx.drawImage(img, 0, 0); const pixels = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
          let red = 0, blue = 0;
          for (let i = 0; i < pixels.length; i += 4) {
            if (pixels[i] > 240 && pixels[i + 1] < 10 && pixels[i + 2] < 10 && pixels[i + 3] > 240) red++;
            if (pixels[i + 2] > 240 && pixels[i] < 10 && pixels[i + 1] < 10 && pixels[i + 3] > 240) blue++;
          }
          return { src: img.src, alt: img.alt, red, blue, width: img.width, height: img.height };
        }));
        const layout = await createPrintLayout(html, doc.page);
        try { states.push({ text: root.textContent, images, pages: layout.pages, printedImages: layout.element.querySelectorAll('img').length, warnings: layout.warnings }); }
        finally { layout.dispose(); }
        if (n < 2) {
          const blob = await exportDocx(doc), data = new Uint8Array(await blob.arrayBuffer()); let binary = '';
          for (let i = 0; i < data.length; i += 8192) binary += String.fromCharCode(...data.subarray(i, i + 8192)); files.push(btoa(binary));
          const loaded = await importFile(new File([blob], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
        }
      }
      return { states, files, controlPages, rawRemains: source.html.includes('data-hwp-object=') };
    } finally { reader.close(); }
  }, { bytes: [...sample.bytes], format, edge, groupHeight: sample.groupHeight });
  expect(result.rawRemains).toBe(false);
  if (edge) expect(result.controlPages).toBe(2);
  for (const state of result.states) {
    expect(state.text).toContain('BODY OUTSIDE GROUP'); expect(state.text).not.toContain('원본 보기');
    expect(state.images).toHaveLength(1); expect(state.printedImages).toBe(1); expect(state.pages).toBe(1); expect(state.warnings).toEqual([]);
    if (!edge) expect(state.images[0].alt).toContain('TITLE 한글'); expect(state.images[0].red).toBeGreaterThan(200); expect(state.images[0].blue).toBe(0);
    expect(state.images[0].src).toBe(result.states[0].images[0].src);
  }
  for (let i = 0; i < result.files.length; i++) await fs.writeFile(info.outputPath(`hwp-gov-object-image-${format}-${edge ? 'edge' : 'toc'}-save${i + 1}.docx`), Buffer.from(result.files[i], 'base64'));
});

test('object pictures use source crop endpoints and preserve transforms and clip bounds', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { rasterizeHwpObject } = await import('/src/hwp/object-preview.ts');
    const source = document.createElement('canvas'); source.width = 20; source.height = 10;
    const c = source.getContext('2d')!; c.fillStyle = '#ff0000'; c.fillRect(0, 0, 10, 10); c.fillStyle = '#0000ff'; c.fillRect(10, 0, 10, 10);
    const bounds = { x: 100, y: 200, width: 40, height: 20 }, base64 = source.toDataURL().split(',')[1];
    const op = { type: 'image', bbox: bounds, mime: 'image/png', base64, crop: { left: 0, top: 0, right: 50, bottom: 100 }, originalSizeHu: [100, 100], effect: 'realPic', brightness: 0, contrast: 0, transform: { horzFlip: true, vertFlip: true, rotation: 180 } };
    const node = { kind: 'clipRect', bounds, clip: { x: 100, y: 200, width: 20, height: 20 }, child: { kind: 'leaf', bounds, ops: [op] } };
    const img = new Image(); img.src = await rasterizeHwpObject(JSON.stringify(node), 44, 24); await img.decode();
    const canvas = document.createElement('canvas'); canvas.width = img.naturalWidth; canvas.height = img.naturalHeight; const ctx = canvas.getContext('2d')!; ctx.drawImage(img, 0, 0);
    return { left: [...ctx.getImageData(20, 20, 1, 1).data], right: [...ctx.getImageData(60, 20, 1, 1).data] };
  });
  expect(result.left).toEqual([255, 0, 0, 255]); expect(result.right[3]).toBe(0);
});

test('dotted leaders use coordinates relative to their own text bounds', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { rasterizeHwpObject } = await import('/src/hwp/object-preview.ts');
    const bounds = { x: 100, y: 200, width: 80, height: 20 }, op = { type: 'tabLeader', bbox: bounds, fontSize: 20, color: '#000000', rotation: 0, isVertical: false, leadersComplete: true, leaders: [{ startX: 10, endX: 70, fillType: 3 }] };
    const image = new Image(); image.src = await rasterizeHwpObject(JSON.stringify({ kind: 'leaf', bounds, ops: [op] }), 84, 24); await image.decode();
    const canvas = document.createElement('canvas'); canvas.width = image.naturalWidth; canvas.height = image.naturalHeight; const ctx = canvas.getContext('2d')!; ctx.drawImage(image, 0, 0);
    const row = ctx.getImageData(0, 24, canvas.width, 1).data; let ink = 0, gaps = 0, transitions = 0, previous = false;
    for (let x = 24; x < 144; x++) { const on = row[x * 4 + 3] > 100; ink += Number(on); gaps += Number(!on); if (on !== previous) transitions++; previous = on; }
    return { ink, gaps, transitions, before: ctx.getImageData(10, 24, 1, 1).data[3], after: ctx.getImageData(155, 24, 1, 1).data[3] };
  });
  expect(result.ink).toBeGreaterThan(10); expect(result.gaps).toBeGreaterThan(10); expect(result.transitions).toBeGreaterThan(10); expect(result.before).toBe(0); expect(result.after).toBe(0);
});

test('unknown images, incomplete crops, page backgrounds and unsupported leaders retain descriptions without external requests', async ({ page }) => {
  await page.goto('/'); const requests: string[] = [];
  page.on('request', request => { if (request.url().startsWith('https://object.invalid')) requests.push(request.url()); });
  const results = await page.evaluate(async () => {
    const { prepareEditableHwp } = await import('/src/hwp/svg-images.ts');
    const { defaultPage } = await import('/src/types.ts');
    const bounds = { x: 10, y: 20, width: 40, height: 20 }, result = [];
    const canvas = document.createElement('canvas'); canvas.width = canvas.height = 1; const base64 = canvas.toDataURL().split(',')[1];
    for (const op of [
      { type: 'image', bbox: bounds, mime: 'image/png', src: 'https://object.invalid/x.png' },
      { type: 'image', bbox: bounds, mime: 'image/svg+xml', base64: btoa('<svg/>') },
      { type: 'image', bbox: bounds, mime: 'image/png', base64, crop: { left: 0, top: 0, right: 100, bottom: 100 } },
      { type: 'image', bbox: bounds, mime: 'image/png', base64, effect: 'grayPic' },
      { type: 'pageBackground', bbox: bounds },
      { type: 'tabLeader', bbox: bounds, fontSize: 12, color: '#000000', leadersComplete: false, leaders: [] },
      { type: 'tabLeader', bbox: bounds, fontSize: 12, color: '#000000', leadersComplete: true, leaders: [{ startX: 0, endX: 20, fillType: 9 }] },
    ]) {
      const image = document.createElement('img'); image.width = 44; image.height = 24; image.alt = 'KEEP ORIGINAL OBJECT DESCRIPTION';
      image.dataset.hwpObject = JSON.stringify({ kind: 'leaf', bounds, ops: [op] });
      result.push(await prepareEditableHwp({ html: image.outerHTML, page: defaultPage, warnings: [] }));
    }
    return result;
  });
  for (const result of results) { expect(result.html).toContain('KEEP ORIGINAL OBJECT DESCRIPTION'); expect(result.html).not.toContain('<img'); expect(result.warnings).toHaveLength(1); }
  expect(requests).toEqual([]);
});

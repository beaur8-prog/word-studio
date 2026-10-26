import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createRequire } from 'node:module';
const createFixture = createRequire(import.meta.url)('./hwp-shape-fixture.cjs');
let directory: string, files: Record<string, string>, grouped: Record<string, string>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-studio-hwp-shapes-')); files = await createFixture(directory); grouped = await createFixture.grouped(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

for (const format of ['hwp', 'hwpx']) test(`${format}: a grouped drawing remains one inline image through two DOCX saves`, async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'group.' + format));
      let doc = importEditableHwp(await session.editable(), 'group.' + format);
      const states = [];
      for (let save = 0; save < 3; save++) {
        const root = document.createElement('div'); root.innerHTML = generateHTML(doc.content, editorExtensions);
        const images = [...root.querySelectorAll('img')];
        states.push({ text: root.textContent, images: await Promise.all(images.map(async img => {
          await img.decode(); const canvas = document.createElement('canvas'); canvas.width = img.naturalWidth; canvas.height = img.naturalHeight;
          const ctx = canvas.getContext('2d')!; ctx.drawImage(img, 0, 0); const pixels = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
          let painted = 0; for (let i = 3; i < pixels.length; i += 4) if (pixels[i]) painted++;
          return { src: img.src, alt: img.alt, floating: img.hasAttribute('data-image-floating'), painted };
        })) });
        if (save < 2) { const reopened = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...reopened.document, content: parseDocumentHtml(reopened.html!) }; }
      }
      return states;
    } finally { session.close(); }
  }, { bytes: [...await fs.readFile(grouped[format])], format });
  for (const state of states) {
    expect(state.text).toContain('BODY TEXT'); expect(state.text).not.toContain('원본 보기'); expect(state.images).toHaveLength(1);
    expect(state.images[0].alt).toContain('TITLE 한글'); expect(state.images[0].floating).toBe(false); expect(state.images[0].painted).toBeGreaterThan(100);
    expect(state.images[0].src).toBe(states[0].images[0].src);
  }
});

for (const format of ['hwp', 'hwpx']) test(`${format}: floating textbox and ellipse keep their pixels, page anchors and descriptions through DOCX`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { editorExtensions } = await import('/src/extensions.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'shapes.' + format));
      const source = await session.editable();
      const doc = importEditableHwp(source, 'shapes.' + format), html = generateHTML(doc.content, editorExtensions);
      const layout = await createPrintLayout(html, doc.page);
      const pictureData = async (html: string) => {
        const root = document.createElement('div'); root.innerHTML = html;
        return await Promise.all([...root.querySelectorAll('img')].map(async image => {
          await image.decode(); const canvas = document.createElement('canvas'); canvas.width = image.naturalWidth; canvas.height = image.naturalHeight;
          const ctx = canvas.getContext('2d')!; ctx.drawImage(image, 0, 0);
          const pixels = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
          let painted = 0; for (let i = 3; i < pixels.length; i += 4) if (pixels[i]) painted++;
          return { alt: image.alt, geometry: JSON.parse(image.dataset.imageFloating!), painted, src: image.src };
        }));
      };
      const original = await pictureData(html);
      const positions = [...layout.element.querySelectorAll('img')].map(image => {
        const paper = image.closest('.print-page')!.getBoundingClientRect(), b = image.getBoundingClientRect();
        return { alt: image.alt, x: b.x - paper.x, y: b.y - paper.y };
      });
      const pages = layout.pages; layout.dispose();
      const saved = await exportDocx(doc), again = await importFile(new File([saved], 'again.docx'));
      return { pages, positions, original, again: await pictureData(again.html!), warnings: source.warnings, body: html, rawPayloadRemains: source.html.includes('data-hwp-object=') };
    } finally { session.close(); }
  }, { bytes: [...await fs.readFile(files[format])], format });
  expect(result.pages).toBe(1); expect(result.original).toHaveLength(2); expect(result.again).toHaveLength(2);
  expect(result.rawPayloadRemains).toBe(false); expect(result.body).toContain('BODY TEXT');
  for (const image of result.original) {
    expect(image.painted).toBeGreaterThan(100);
    const again = result.again.find(p => p.alt === image.alt)!;
    expect(again.src).toBe(image.src);
    expect(again.geometry).toMatchObject({ horizontalPosition: image.geometry.horizontalPosition,
      verticalPosition: image.geometry.verticalPosition, wrap: image.geometry.wrap, behindDocument: image.geometry.behindDocument });
  }
  const title = result.positions.find(p => p.alt.includes('TITLE 한글'))!;
  expect(title.x).toBeCloseTo(118, 0); expect(title.y).toBeCloseTo(94, 0);
  const ellipse = result.positions.find(p => p.alt === '타원')!;
  expect(ellipse.x).toBeCloseTo(398, 0); expect(ellipse.y).toBeCloseTo(598, 0);
  expect(result.warnings.some(w => w.includes('개체 내부는 편집할 수 없습니다'))).toBe(true);
});

test('a nonwrapping page decoration extending into the bottom margin does not create another page', async ({ page }) => {
  await page.goto('/');
  const count = await page.evaluate(async () => {
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { defaultPage } = await import('/src/types.ts');
    const root = document.createElement('div'), p = document.createElement('p'), image = document.createElement('img');
    p.textContent = 'BODY'; image.src = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==';
    image.width = 2; image.height = 1000;
    image.dataset.imageFloating = JSON.stringify({ horizontalPosition: { relative: 'page', offset: 400 * 9525 }, verticalPosition: { relative: 'page', offset: 100 * 9525 }, wrap: { type: 0 }, behindDocument: true });
    p.append(image); root.append(p); const layout = await createPrintLayout(root.innerHTML, defaultPage);
    try { return layout.pages; } finally { layout.dispose(); }
  });
  expect(count).toBe(1);
});

test('object preparation draws only its own clipped paths and rejects unsupported or oversized content', async ({ page }) => {
  await page.goto('/'); const requests: string[] = [];
  page.on('request', request => { if (request.url().startsWith('https://object.invalid')) requests.push(request.url()); });
  const result = await page.evaluate(async () => {
    const { rasterizeHwpObject } = await import('/src/hwp/object-preview.ts');
    const { prepareEditableHwp } = await import('/src/hwp/svg-images.ts');
    const { defaultPage } = await import('/src/types.ts');
    const bounds = { x: 100, y: 200, width: 20, height: 20 };
    const leaf = { kind: 'leaf', bounds, ops: [{ type: 'path', bbox: bounds, commands: [{ type: 'moveTo', x: 100, y: 200 }, { type: 'lineTo', x: 120, y: 200 }, { type: 'lineTo', x: 120, y: 220 }, { type: 'lineTo', x: 100, y: 220 }, { type: 'closePath' }], style: { fillColor: '#ff0000' } }] };
    const node = { kind: 'clipRect', bounds, clip: { x: 100, y: 200, width: 10, height: 20 }, child: leaf };
    const png = await rasterizeHwpObject(JSON.stringify(node), 24, 24);
    const img = new Image(); img.src = png; await img.decode();
    const canvas = document.createElement('canvas'); canvas.width = canvas.height = 48; const ctx = canvas.getContext('2d')!; ctx.drawImage(img, 0, 0);
    const left = [...ctx.getImageData(10, 20, 1, 1).data], right = [...ctx.getImageData(35, 20, 1, 1).data];
    const root = document.createElement('div'), bad = document.createElement('img'); bad.width = bad.height = 24; bad.alt = 'KEEP OBJECT TEXT';
    bad.dataset.hwpObject = JSON.stringify({ ...leaf, ops: [{ type: 'image', src: 'https://object.invalid/a.png' }] }); root.append(bad);
    const fallback = await prepareEditableHwp({ html: root.innerHTML, page: defaultPage, warnings: [] });
    let oversized = false, aborted = false;
    try { await rasterizeHwpObject(JSON.stringify(node), 10001, 24); } catch { oversized = true; }
    const controller = new AbortController();
    bad.dataset.hwpObject = JSON.stringify(node);
    try { await prepareEditableHwp({ html: root.innerHTML + root.innerHTML, page: defaultPage, warnings: [] }, () => controller.abort(), controller.signal); } catch { aborted = controller.signal.aborted; }
    return { left, right, fallback, oversized, aborted };
  });
  expect(result.left).toEqual([255, 0, 0, 255]); expect(result.right[3]).toBe(0);
  expect(result.fallback.html).toContain('KEEP OBJECT TEXT'); expect(result.fallback.html).not.toContain('<img');
  expect(result.fallback.warnings).toHaveLength(1); expect(result.oversized).toBe(true); expect(result.aborted).toBe(true); expect(requests).toEqual([]);
});

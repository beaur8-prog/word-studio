import { test, expect, type Page } from '@playwright/test';
import { writeFile } from 'node:fs/promises';

const editor = (page: Page) => page.getByRole('textbox', { name: '문서 편집 영역', exact: true });

async function load(page: Page, doc: unknown) {
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'floating-page-segments.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(doc)),
  });
  await expect(page.locator('.busy-indicator')).toBeHidden();
  await page.evaluate(() => document.fonts.ready);
}

async function observe(page: Page, duration = 3000) {
  return page.evaluate(async duration => {
    const { documentScale } = await import('/src/document-scale.ts');
    const samples = [];
    const start = performance.now();
    do {
      const root = document.querySelector<HTMLElement>('.tiptap[contenteditable=true]')!;
      const sheet = root.closest<HTMLElement>('.document-stack')!, scale = documentScale(root), origin = root.getBoundingClientRect();
      const before = [...root.querySelectorAll('p')].find(p => p.textContent?.startsWith('BEFORE '))!;
      samples.push({ pages: Number(sheet.dataset.pages), beforeHeight: before.getBoundingClientRect().height / scale,
        images: [...root.querySelectorAll<HTMLImageElement>('img[data-image-floating]')].map(img => {
          const box = img.getBoundingClientRect(), h = Number(sheet.dataset.pageHeight), top = Number(sheet.dataset.margin);
          const page = Math.max(0, Math.floor(((box.top - origin.top) / scale - top + .5) / (h + 24)));
          return { alt: img.alt, page: page + 1, x: (box.left - origin.left) / scale,
            y: (box.top - origin.top) / scale - page * (h + 24), width: box.width / scale, height: box.height / scale, style: img.style.cssText, anchor: img.closest('p')?.getBoundingClientRect().top, rootTop: origin.top, scale };
        }) });
      await new Promise(resolve => setTimeout(resolve, 100));
    } while (performance.now() - start < duration);
    return samples;
  }, duration);
}

for (const lead of [8, 20]) test(`later picture cannot wrap across a source page break (${lead} leading lines), including two DOCX saves`, async ({ page }) => {
  test.setTimeout(150000);
  const errors: string[] = []; page.on('pageerror', error => errors.push(String(error)));
  await page.goto('/'); await expect(editor(page)).toBeVisible();
  const result = await page.evaluate(async lead => {
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
    const { defaultPage } = await import('/src/types.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { withoutEditorReflow } = await import('/src/editor-layout-task.ts');
    const raster = (color: string) => {
      const canvas = document.createElement('canvas'); canvas.width = canvas.height = 4;
      const context = canvas.getContext('2d')!; context.fillStyle = color; context.fillRect(0, 0, 4, 4); return canvas.toDataURL();
    };
    const paragraph = (text: string) => `<p style="font-family:Arial;font-size:12pt;line-height:20px;margin:0">${text}</p>`;
    const picture = (alt: string, y: number, height: number, src: string) => {
      const geometry = { horizontalPosition: { relative: 'page', offset: 96 * 9525 },
        verticalPosition: { relative: 'page', offset: y * 9525 }, wrap: { type: 1, side: 'bothSides' }, layoutInCell: true, allowOverlap: true };
      return `<img data-inline-image="true" src="${src}" width="600" height="${height}" alt="${alt}" data-image-floating="${JSON.stringify(geometry).replaceAll('"', '&quot;')}">`;
    };
    const html = paragraph(picture('EARLY DRAWING', 96, 120, raster('red')) + 'EARLY ANCHOR') +
      Array.from({ length: lead }, (_, i) => paragraph('LEADING ' + i)).join('') +
      paragraph('BEFORE ' + 'Ordinary text words for a later anchored picture. '.repeat(18)) +
      '<div data-page-break="true"></div>' + paragraph(picture('LATER DRAWING', 174, 650, raster('blue')) + 'ANCHOR') +
      paragraph('AFTER ' + 'Following text must remain visible and in order. '.repeat(8));
    let doc: any = { format: 'word-studio', version: 1, id: 'float-page-break', title: 'Floating page break',
      updatedAt: new Date().toISOString(), comments: [], page: { ...defaultPage, pageNumbers: false }, content: parseDocumentHtml(html) };
    const states = [], saves = [];
    for (let pass = 0; pass < 3; pass++) {
      const source = document.createElement('div'); source.innerHTML = generateHTML(doc.content, editorExtensions);
      const images = [...source.querySelectorAll<HTMLImageElement>('img')].map(img => ({ src: img.src, alt: img.alt,
        width: Number(img.getAttribute('width')), height: Number(img.getAttribute('height')), geometry: JSON.parse(img.dataset.imageFloating!) }));
      const printed = await withoutEditorReflow(async () => {
        const layout = await createPrintLayout(source.innerHTML, doc.page);
        try {
          const before = [...layout.element.querySelectorAll('p')].find(p => p.textContent?.startsWith('BEFORE '))!;
          return { pages: layout.pages, beforeHeight: before.getBoundingClientRect().height,
            images: [...layout.element.querySelectorAll<HTMLImageElement>('.print-content img')].map(img => {
              const box = img.getBoundingClientRect(), sheet = img.closest<HTMLElement>('.print-page')!, paper = sheet.getBoundingClientRect();
              return { alt: img.alt, page: Number(sheet.dataset.page), x: box.left - paper.left, y: box.top - paper.top, width: box.width, height: box.height };
            }) };
        } finally { layout.dispose(); }
      });
      states.push({ doc, text: source.textContent, breaks: source.querySelectorAll('[data-page-break]').length, images, printed });
      if (pass < 2) {
        const blob = await exportDocx(doc); saves.push([...new Uint8Array(await blob.arrayBuffer())]);
        const opened = await importFile(new File([blob], 'floating-page-break.docx'));
        doc = { ...opened.document, content: parseDocumentHtml(opened.html!) };
      }
    }
    return { states, saves };
  }, lead);
  for (const [index, bytes] of result.saves.entries()) await writeFile(test.info().outputPath(`textbox-save-floating-page-break-${lead}-${index + 1}.docx`), Buffer.from(bytes));
  const observations = [];
  for (const state of result.states) {
    expect(state.text).toBe(result.states[0].text); expect(state.breaks).toBe(1);
    expect(state.images.map(({ src, alt, width, height }) => ({ src, alt, width, height }))).toEqual(result.states[0].images.map(({ src, alt, width, height }) => ({ src, alt, width, height })));
    expect(state.printed.pages).toBe(2); expect(state.printed.beforeHeight).toBeCloseTo(220, 0);
    await load(page, state.doc); await expect(editor(page)).toContainText('AFTER ');
    await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe('2');
    await page.waitForTimeout(400);
    const samples = await observe(page, Number(process.env.FLOAT_STABILITY_MS || 3000)); observations.push(samples);
    for (const sample of samples) {
      expect(sample.pages).toBe(2); expect(sample.beforeHeight).toBeCloseTo(220, 0);
      expect(sample.images.length).toBe(2);
      for (const [i, image] of sample.images.entries()) {
        const print = state.printed.images[i]; expect(image.page).toBe(print.page);
        for (const key of ['x', 'y', 'width', 'height'] as const) expect(Math.abs(image[key] - print[key]), JSON.stringify({key,image,print})).toBeLessThan(.75);
      }
    }
    for (const zoom of [50, 125, 100]) {
      const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
      await page.waitForTimeout(350);
      const sample = (await observe(page, 200)).at(-1)!; expect(sample.pages).toBe(2); expect(sample.beforeHeight).toBeCloseTo(220, 0);
      for (const [i, image] of sample.images.entries()) for (const key of ['x', 'y', 'width', 'height'] as const)
        expect(Math.abs(image[key] - state.printed.images[i][key]), `${zoom}% ${key}`).toBeLessThan(.75);
    }
    const original = await editor(page).evaluate((root: any) => root.editor.getJSON());
    await editor(page).evaluate((root: any) => {
      let end = 0; root.editor.state.doc.descendants((node: any, pos: number) => { if (node.isTextblock && node.textContent.startsWith('BEFORE ')) end = pos + 1 + node.content.size; });
      root.editor.commands.setTextSelection(end); root.editor.commands.focus();
    });
    await page.keyboard.type(' EDITED'); await expect(editor(page)).toContainText(' EDITED');
    await page.keyboard.press('ControlOrMeta+z'); await expect.poll(() => editor(page).evaluate((root: any) => root.editor.getJSON())).toEqual(original);
    await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(editor(page)).toContainText(' EDITED');
    await page.keyboard.press('ControlOrMeta+z'); await expect.poll(() => editor(page).evaluate((root: any) => root.editor.getJSON())).toEqual(original);
    for (const sample of await observe(page, 500)) expect(sample.pages).toBe(2);
  }
  await writeFile(test.info().outputPath('page-segment-observations.json'), JSON.stringify({ lead, printed: result.states.map(s => s.printed), observations, errors }, null, 2));
  expect(errors).toEqual([]);
});

test('an empty body anchor keeps its authored blank lines without reserving its own drawing twice', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
    const { defaultPage } = await import('/src/types.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';
    const geometry = { horizontalPosition: { relative: 'page', offset: 96 * 9525 }, verticalPosition: { relative: 'page', offset: 96 * 9525 }, wrap: { type: 3 }, layoutInCell: true };
    let doc: any = { format: 'word-studio', version: 1, id: 'empty-float', title: 'Empty anchor', comments: [], updatedAt: new Date().toISOString(), page: { ...defaultPage, pageNumbers: false },
      content: parseDocumentHtml(`<p style="font-family:Arial;font-size:11pt;line-height:.75pt;margin:0" data-word-line-rule="exact"><img data-inline-image="true" src="data:image/png;base64,${png}" width="600" height="900" alt="EMPTY ANCHOR" data-image-floating="${JSON.stringify(geometry).replaceAll('"', '&quot;')}"><br><br><br></p><p style="font-family:Arial;font-size:12pt;line-height:20px;margin:0">AFTER IMAGE</p>`) };
    const states = [];
    for (let pass = 0; pass < 3; pass++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      const anchor = layout.element.querySelector('.print-content p')!, image = anchor.querySelector('img')!, after = [...layout.element.querySelectorAll('p')].find(p => p.textContent === 'AFTER IMAGE')!;
      const picture = image.getBoundingClientRect(), range = document.createRange(), walker = document.createTreeWalker(after, NodeFilter.SHOW_TEXT);
      let text: Node | null = null; while (walker.nextNode()) if (walker.currentNode.textContent?.trim()) { text = walker.currentNode; break; }
      if (!text) throw Error('Missing AFTER IMAGE text'); range.selectNodeContents(text); const next = range.getBoundingClientRect();
      states.push({ doc, pages: layout.pages, anchorHeight: anchor.getBoundingClientRect().height, hardBreaks: anchor.querySelectorAll('br:not([data-flow-widget])').length,
        image: { width: picture.width, height: picture.height }, afterTop: next.top, imageBottom: picture.bottom }); layout.dispose();
      if (pass < 2) { const opened = await importFile(new File([await exportDocx(doc)], 'empty-anchor.docx')); doc = { ...opened.document, content: parseDocumentHtml(opened.html!) }; }
    }
    return states;
  });
  for (const state of result) {
    expect(state.pages).toBe(1); expect(state.anchorHeight).toBeLessThan(10); expect(state.hardBreaks).toBe(3);
    expect(state.image.width).toBeCloseTo(600, 0); expect(state.image.height).toBeCloseTo(900, 0);
    expect(state.afterTop).toBeGreaterThanOrEqual(state.imageBottom - .5);
    await load(page, state.doc); await expect(editor(page)).toContainText('AFTER IMAGE');
    await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe('1');
    await expect.poll(() => editor(page).locator('p').first().evaluate(p => p.getBoundingClientRect().height)).toBeLessThan(10);
  }
});

test('a later picture still wraps preceding text when both fit on the same source page', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
    const { defaultPage } = await import('/src/types.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';
    const geometry = { horizontalPosition: { relative: 'page', offset: 396 * 9525 }, verticalPosition: { relative: 'page', offset: 96 * 9525 }, wrap: { type: 1 }, layoutInCell: true };
    const paragraph = (text: string) => `<p style="font-family:Arial;font-size:12pt;line-height:20px;margin:0">${text}</p>`;
    const before = paragraph('BEFORE ' + 'Ordinary text must wrap around the later picture. '.repeat(8));
    const image = `<img data-inline-image="true" src="data:image/png;base64,${png}" width="180" height="120" alt="SAME PAGE" data-image-floating="${JSON.stringify(geometry).replaceAll('"', '&quot;')}">`;
    const settings = { ...defaultPage, pageNumbers: false };
    const control = await createPrintLayout(before + paragraph('ANCHOR'), settings);
    const controlHeight = control.element.querySelector('.print-content p')!.getBoundingClientRect().height; control.dispose();
    const content = parseDocumentHtml(before + paragraph(image + 'ANCHOR'));
    const layout = await createPrintLayout(generateHTML(content, editorExtensions), settings);
    try {
      const p = layout.element.querySelector('.print-content p')!, picture = layout.element.querySelector('img[alt="SAME PAGE"]')!.getBoundingClientRect();
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT), overlap = [];
      while (walker.nextNode()) {
        const n = walker.currentNode as Text; if (!n.data.trim() || n.parentElement?.closest('[data-flow-widget]')) continue;
        for (const part of new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(n.data)) {
          if (!part.segment.trim()) continue;
          const range = document.createRange(); range.setStart(n, part.index); range.setEnd(n, part.index + part.segment.length);
          const box = range.getBoundingClientRect();
          if (Math.min(box.right, picture.right) - Math.max(box.left, picture.left) > .5 && Math.min(box.bottom, picture.bottom) - Math.max(box.top, picture.top) > .5) overlap.push(part.segment);
        }
      }
      return { doc: { format: 'word-studio', version: 1, id: 'same-page', title: 'Same page', comments: [], updatedAt: new Date().toISOString(), page: settings, content },
        pages: layout.pages, controlHeight, height: p.getBoundingClientRect().height, overlap };
    } finally { layout.dispose(); }
  });
  expect(result.pages).toBe(1); expect(result.height).toBeGreaterThan(result.controlHeight); expect(result.overlap).toEqual([]);
  await load(page, result.doc); await expect(editor(page)).toContainText('BEFORE ');
  await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe('1');
  await expect.poll(() => editor(page).locator('p').first().evaluate(p => p.getBoundingClientRect().height)).toBeCloseTo(result.height, 0);
});

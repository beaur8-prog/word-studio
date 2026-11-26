import { expect, test } from '@playwright/test';

test('a nonwrapping second-column image stays at its paper coordinates in editor and print through two DOCX saves', async ({ page }) => {
  await page.goto('/');
  const { stages, documents } = await page.evaluate(async () => {
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { defaultPage } = await import('/src/types.ts');
    const canvas = document.createElement('canvas');
    canvas.width = 4; canvas.height = 4;
    const context = canvas.getContext('2d')!;
    context.fillStyle = '#ce1a48';
    context.fillRect(0, 0, canvas.width, canvas.height);
    const source = canvas.toDataURL();
    const anchor = JSON.stringify({
      horizontalPosition: { relative: 'page', offset: 630 * 9525 },
      verticalPosition: { relative: 'page', offset: 75 * 9525 },
      wrap: { type: 0 }, layoutInCell: false
    }).replaceAll('"', '&quot;');
    const html = '<div data-column-count="2" data-column-space="24" data-column-balance="false">' +
      '<p>LEFT COLUMN</p><div data-column-break="true"></div><p>RIGHT COLUMN' +
      '<img data-inline-image="true" src="' + source + '" width="40" height="36" alt="Paper-relative mark" data-image-floating="' + anchor + '">' +
      '</p><p>RIGHT END</p></div><p>AFTER COLUMNS</p>';
    let doc = {
      format: 'word-studio', version: 1, id: 'paper-column-float', title: 'Paper-column float',
      updatedAt: new Date().toISOString(), comments: [],
      page: { ...defaultPage, size: 'Letter', orientation: 'landscape', margin: 25, pageNumbers: false },
      content: parseDocumentHtml(html)
    };
    const stages = [];
    const documents = [];
    for (let step = 0; step < 3; step++) {
      documents.push(doc);
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try {
        const mark = layout.element.querySelector<HTMLImageElement>('img[alt="Paper-relative mark"]')!;
        const paper = mark.closest('.print-page')!.getBoundingClientRect();
        const box = mark.getBoundingClientRect();
        const owner = mark.closest('p')!;
        const column = owner.closest('[data-column-count]')!.getBoundingClientRect();
        stages.push({
          pages: layout.pages,
          ownerInSecondColumn: owner.getBoundingClientRect().left > column.left + column.width / 2,
          x: box.left - paper.left, y: box.top - paper.top, width: box.width, height: box.height,
          anchorText: owner.textContent, imageCandidates: layout.element.querySelectorAll('img[alt="Paper-relative mark"]').length,
          imageSource: mark.src
        });
      } finally {
        layout.dispose();
      }
      if (step < 2) {
        const exported = await exportDocx(doc);
        const opened = await importFile(new File([exported], 'paper-column-float.docx'));
        doc = { ...opened.document, content: parseDocumentHtml(opened.html!) } as typeof doc;
      }
    }
    return { stages, documents };
  });
  for (const stage of stages) {
    expect(stage.pages).toBe(1);
    expect(stage.ownerInSecondColumn).toBe(true);
    expect(stage.x).toBeCloseTo(630, 0);
    expect(stage.y).toBeCloseTo(75, 0);
    expect(stage.width).toBeCloseTo(40, 0);
    expect(stage.height).toBeCloseTo(36, 0);
    expect(stage.anchorText).toBe('RIGHT COLUMN');
    expect(stage.imageCandidates).toBe(1);
    expect(stage.imageSource).toBe(stages[0].imageSource);
  }
  const requests = await page.evaluateHandle(() => {
    let frame = 0;
    const schedule = () => {
      frame = requestAnimationFrame(schedule);
      document.dispatchEvent(new Event('studio-editor-layout-resume'));
    };
    schedule();
    return { stop: () => cancelAnimationFrame(frame) };
  });
  try {
    for (const doc of [documents[0], documents[2]]) {
      await page.locator('input[type=file]').first().setInputFiles({
        name: 'paper-column.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(doc)),
      });
      const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
      const picture = editor.locator('img[alt="Paper-relative mark"]');
      for (const zoom of [100, 150, 75, 100]) {
        const slider = page.getByRole('slider', { name: '확대 축소' });
        await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
        await expect.poll(() => picture.evaluate(image => {
          const paper = image.closest<HTMLElement>('[data-page-height]')!;
          const sheet = paper.getBoundingClientRect(), rect = image.getBoundingClientRect();
          const scale = sheet.width / parseFloat(getComputedStyle(paper).width);
          const x = (rect.left - sheet.left) / scale, y = (rect.top - sheet.top) / scale;
          return Math.max(Math.abs(x - 630), Math.abs(y - 75));
        })).toBeLessThan(1);
      }
      await expect(editor.getByText('RIGHT COLUMN', { exact: true })).toHaveCount(1);
    }
  } finally {
    await requests.evaluate(requests => requests.stop());
    await requests.dispose();
  }
});

test('an attested paper cover retains its extent and artwork through two DOCX saves', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { defaultPage } = await import('/src/types.ts');
    const width = 210 * 96 / 25.4, height = 297 * 96 / 25.4;
    const raster = document.createElement('canvas');
    raster.width = Math.round(width); raster.height = Math.round(height);
    const paint = raster.getContext('2d')!;
    paint.fillStyle = '#d58b42';
    paint.fillRect(0, 0, raster.width, raster.height);
    const source = raster.toDataURL('image/png');
    const fixture = (mode: 'paper' | 'no-evidence' | 'moved' | 'resized') => {
      const anchor = JSON.stringify({
        horizontalPosition: { relative: 'page', offset: mode === 'moved' ? 9525 : 0 },
        verticalPosition: { relative: 'page', offset: 0 },
        behindDocument: true, wrap: { type: 0 }, layoutInCell: false,
        ...(mode !== 'no-evidence' && { hwpPaperDrawing: { width, height } }),
      }).replaceAll('"', '&quot;');
      return {
        format: 'word-studio', version: 1, id: 'paper-cover', title: 'Paper cover',
        updatedAt: new Date().toISOString(), comments: [],
        page: { ...defaultPage, size: 'A4', orientation: 'portrait', margin: 25, pageNumbers: false },
        content: parseDocumentHtml(`<p><img data-inline-image="true" data-image-floating="${anchor}" src="${source}"
          width="${width}" height="${mode === 'resized' ? height - 10 : height}" alt="Paper cover"></p><p>Readable text above cover</p>`),
      };
    };
    const measure = async (doc: ReturnType<typeof fixture>) => {
      const print = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try {
        const image = print.element.querySelector<HTMLImageElement>('img[alt="Paper cover"]')!;
        const p = image.closest('.print-page')!, bounds = p.getBoundingClientRect();
        const region = image.getBoundingClientRect(), style = getComputedStyle(image);
        return { pages: print.pages, size: [bounds.width, bounds.height],
          coordinates: [region.x - bounds.x, region.y - bounds.y, region.width, region.height],
          marker: image.dataset.hwpFullWidthPicture ?? null, bytes: image.src,
          text: p.querySelector('.print-content')!.textContent, zIndex: style.zIndex };
      } finally { print.dispose(); }
    };
    let doc = fixture('paper');
    const stages = [];
    const documents = [];
    for (let stage = 0; stage < 3; stage++) {
      documents.push(doc);
      stages.push(await measure(doc));
      if (stage < 2) {
        const bytes = await exportDocx(doc);
        const opened = await importFile(new File([bytes], 'paper-cover.docx'));
        doc = { ...opened.document, content: parseDocumentHtml(opened.html!) } as typeof doc;
      }
    }
    const controls = await Promise.all((['no-evidence', 'moved', 'resized'] as const).map(mode => measure(fixture(mode))));
    return { source, stages, controls, documents };
  });
  for (const stage of result.stages) {
    expect(stage).toMatchObject({ pages: 1, marker: 'true', bytes: result.source, text: 'Readable text above cover', zIndex: '-1' });
    expect(stage.coordinates.slice(0, 2)).toEqual([0, 0]);
    expect(Math.abs(stage.coordinates[2] - stage.size[0])).toBeLessThan(1);
    expect(Math.abs(stage.coordinates[3] - stage.size[1])).toBeLessThan(1);
  }
  for (const control of result.controls) {
    expect(control).toMatchObject({ pages: 1, marker: null, bytes: result.source, text: 'Readable text above cover' });
    expect(control.coordinates[2]).toBeLessThan(control.size[0] - 10);
  }

  for (const doc of result.documents) {
    await page.locator('input[type=file]').first().setInputFiles({
      name: 'paper-cover.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(doc)),
    });
    const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
    const picture = editor.locator('img[alt="Paper cover"][data-image-floating]');
    await expect(picture).toBeVisible();
    await page.evaluate(() => document.dispatchEvent(new Event('studio-editor-layout-resume')));
    await expect(picture).toHaveCSS('max-width', 'none');
    for (const zoom of [100, 150, 75, 100]) {
      const slider = page.getByRole('slider', { name: '확대 축소' });
      await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
      await expect.poll(() => picture.evaluate(image => {
        const paper = image.closest<HTMLElement>('[data-page-height]')!;
        const sheet = paper.getBoundingClientRect(), rect = image.getBoundingClientRect();
        const paperWidth = parseFloat(getComputedStyle(paper).width);
        const scale = sheet.width / paperWidth;
        const expectedHeight = paperWidth * 297 / 210;
        return Math.max(Math.abs((rect.left - sheet.left) / scale), Math.abs((rect.top - sheet.top) / scale),
          Math.abs(rect.width / scale - paperWidth), Math.abs(rect.height / scale - expectedHeight));
      })).toBeLessThan(1);
    }
    await expect(editor.getByText('Readable text above cover', { exact: true })).toHaveCount(1);
  }

});

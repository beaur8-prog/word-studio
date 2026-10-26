import { test, expect } from '@playwright/test';

test('table headers stay editable with the first body group across a page boundary and save without view gaps', async ({ page }) => {
  await page.goto('/');
  const source = `<p style="line-height:20px;margin:0">OPEN${'<br>line'.repeat(20)}</p><table style="margin:0"><tr style="height:30px"><th><p>HEADER ONE</p></th><th><p>HEADER TWO</p></th></tr><tr style="height:20px"><th><p>SUB ONE</p></th><th><p>SUB TWO</p></th></tr><tr style="height:110px"><td rowspan="2"><p>BODY LEFT</p></td><td><p>BODY RIGHT ONE</p></td></tr><tr style="height:110px"><td><p>BODY RIGHT TWO</p></td></tr></table>`;
  const { defaultPage } = await page.evaluate(async () => await import('/src/types.ts'));
  const doc = await page.evaluate(async ({ source, paper }) => {
    const { importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const loaded = await importFile(new File([source], 'headers.html'));
    return { ...loaded.document, page: { ...paper, size: 'Custom', width: 150, height: 180, margin: 10 }, content: parseDocumentHtml(loaded.html!) };
  }, { source, paper: defaultPage });
  await page.locator('input[type=file]').first().setInputFiles({ name: 'headers.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(doc)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  await expect(editor.locator('th').first()).toHaveText('HEADER ONE');
  const positions = () => editor.locator('table tr:not(.pagination-gap)').evaluateAll(rows => {
    const root = rows[0].closest('.tiptap')!, origin = root.getBoundingClientRect().top, stack = root.closest<HTMLElement>('.document-stack')!;
    const height = Number(stack.dataset.pageHeight), top = Number(stack.dataset.margin);
    return rows.map(row => Math.floor((row.getBoundingClientRect().top - origin - top + .5) / (height + 24)) + 1);
  });
  await expect.poll(positions).toEqual([2, 2, 2, 2]);
  await editor.locator('th p').first().fill('EDITED HEADER');
  await page.keyboard.press('ControlOrMeta+z'); await expect(editor.locator('th').first()).toHaveText('HEADER ONE');
  await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(editor.locator('th').first()).toHaveText('EDITED HEADER');
  await expect.poll(positions).toEqual([2, 2, 2, 2]);
  const body = editor.locator('td p').filter({ hasText: 'BODY RIGHT ONE' });
  await body.fill('BODY EDIT '.repeat(40));
  await expect.poll(positions).toEqual([2, 2, 2, 2]);
  await page.keyboard.press('ControlOrMeta+z'); await expect(body).toHaveText('BODY RIGHT ONE');
  const download = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
  const saved = JSON.parse(await (await import('node:fs/promises')).readFile((await (await download).path())!, 'utf8'));
  expect(JSON.stringify(saved.content)).not.toContain('pagination-gap');
  expect(saved.content.content.find((n: any) => n.type === 'table').content).toHaveLength(4);
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  await expect(page.locator('.print-page')).toHaveCount(2);
  await expect(page.locator('.print-page').first()).not.toContainText('HEADER');
  await expect(page.locator('.print-page').nth(1)).toContainText('EDITED HEADER');
});

test('fractional paper widths keep long-document page starts and page counts at 50, 100, 125 and 200 percent zoom', async ({ page }) => {
  await page.goto('/');
  const doc = await page.evaluate(async () => {
    const { defaultPage } = await import('/src/types.ts');
    const { importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const html = Array.from({ length: 400 }, (_, i) => `<p style="line-height:20px;margin:0;font-size:10pt">LINE ${i}</p>`).join('');
    const loaded = await importFile(new File([html], 'long.html'));
    return { ...loaded.document, page: { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10 }, content: parseDocumentHtml(loaded.html!) };
  });
  await page.locator('input[type=file]').first().setInputFiles({ name: 'long.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(doc)) });
  const slider = page.getByRole('slider', { name: '확대 축소' });
  for (const zoom of [100, 50, 125, 200, 100]) {
    await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe('14');
    await expect.poll(async () => page.locator('.tiptap').evaluate((root, zoom) => {
      const stack = root.closest<HTMLElement>('.document-stack')!, height = Number(stack.dataset.pageHeight), margin = Number(stack.dataset.margin), bottom = Number(stack.dataset.bottomMargin);
      const origin = root.getBoundingClientRect().top, scale = zoom / 100;
      return Math.max(...[...root.querySelectorAll(':scope > p')].map(p => {
        const r = p.getBoundingClientRect(), y = (r.top - origin) / scale, page = Math.floor((y - margin + .5) / (height + 24));
        return Math.max(page * (height + 24) + margin - y, y + r.height / scale - (page * (height + 24) + height - bottom));
      }));
    }, zoom)).toBeLessThan(.6);
  }
});

test('document scale comes from nested transforms and CSS zoom rather than rounded element widths', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { documentScale } = await import('/src/document-scale.ts');
    const outer = document.createElement('div'), inner = document.createElement('div');
    outer.style.cssText = 'position:absolute;width:1122.52px;transform-origin:0 0'; inner.style.width = '100%'; outer.append(inner); document.body.append(outer);
    try {
      const at100 = documentScale(inner), roundedRatio = inner.getBoundingClientRect().width / inner.offsetWidth;
      outer.style.transform = 'scale(1.25)'; inner.style.transform = 'scale(2)';
      const nested = documentScale(inner); inner.style.zoom = '1.1';
      return { at100, roundedRatio, nested, cssZoom: documentScale(inner) };
    } finally { outer.remove(); }
  });
  expect(result.at100).toBe(1); expect(result.roundedRatio).not.toBe(1); expect(result.nested).toBe(2.5); expect(result.cssZoom).toBeCloseTo(2.75, 8);
});

test('continuing editable table cells keep ink out of page margins at every zoom without fragmenting saved rows', async ({ page }) => {
  await page.goto('/');
  const doc = await page.evaluate(async () => {
    const { defaultPage } = await import('/src/types.ts');
    const { importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const column = (name: string, count: number) => Array.from({ length: count }, (_, i) => `<p style="line-height:20px;margin:0">${name} ${i} 한글 paragraph</p>`).join('');
    const html = `<p>START</p><table><tr><th><p>LEFT HEADER</p></th><th><p>RIGHT HEADER</p></th></tr><tr data-row-split="allow"><td>${column('LEFT', 100)}</td><td>${column('RIGHT', 30)}<table><tr><td><p>NESTED A</p></td><td><p>NESTED B</p></td></tr><tr><td><p>NESTED C</p></td><td><p>NESTED D</p></td></tr></table>${column('RIGHT TAIL', 35)}</td></tr></table><p>END</p>`;
    const loaded = await importFile(new File([html], 'cells.html'));
    return { ...loaded.document, page: { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10 }, content: parseDocumentHtml(loaded.html!) };
  });
  await page.locator('input[type=file]').first().setInputFiles({ name: 'cells.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(doc)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  await expect(editor.locator('.pagination-cell-gap').first()).toBeAttached();
  let count = '';
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    const overflow = () => editor.evaluate((root, zoom) => {
      const stack = root.closest<HTMLElement>('.document-stack')!, height = Number(stack.dataset.pageHeight), top = Number(stack.dataset.margin), bottom = Number(stack.dataset.bottomMargin), stride = height + 24;
      const origin = root.getBoundingClientRect().top, scale = zoom / 100; let overflow = 0;
      for (const p of root.querySelectorAll('td p')) {
        const range = document.createRange(); range.selectNodeContents(p);
        for (const r of range.getClientRects()) {
          if (r.width < 1 || r.height < 1) continue;
          const y = (r.top - origin) / scale, end = (r.bottom - origin) / scale, page = Math.max(0, Math.floor((y - top + .5) / stride));
          overflow = Math.max(overflow, page * stride + top - y, end - (page * stride + height - bottom));
        }
      }
      return overflow;
    }, zoom);
    await expect.poll(overflow).toBeLessThan(1);
    const current = await page.locator('.document-stack').getAttribute('data-pages');
    if (!count) count = current!; else expect(current).toBe(count);
  }
  const last = editor.locator('td p').filter({ hasText: /^LEFT 99 / });
  await last.fill('EDITED LAST CELL'); await page.keyboard.press('ControlOrMeta+z'); await expect(last).toContainText('LEFT 99');
  const download = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
  const saved = JSON.parse(await (await import('node:fs/promises')).readFile((await (await download).path())!, 'utf8'));
  expect(JSON.stringify(saved.content)).not.toContain('pagination-cell-gap');
  const table = saved.content.content.find((n: any) => n.type === 'table'); expect(table.content).toHaveLength(2);
  expect(table.content[1].content.map((c: any) => c.content.length)).toEqual([100, 66]);
});

test('a single long cell paragraph can continue at line boundaries while remaining editable', async ({ page }) => {
  await page.goto('/');
  const doc = await page.evaluate(async () => {
    const { defaultPage } = await import('/src/types.ts'), { importFile } = await import('/src/document-io.ts'), { parseDocumentHtml } = await import('/src/extensions.ts');
    const html = `<table><tr data-row-split="allow"><td><p style="line-height:20px">${'한글 long paragraph '.repeat(280)}</p></td><td><p>SHORT</p></td></tr></table>`;
    const loaded = await importFile(new File([html], 'paragraph.html'));
    return { ...loaded.document, page: { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10 }, content: parseDocumentHtml(loaded.html!) };
  });
  await page.locator('input[type=file]').first().setInputFiles({ name: 'paragraph.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(doc)) });
  await expect(page.locator('.pagination-cell-gap').first()).toBeAttached();
  await expect.poll(() => page.locator('.tiptap').evaluate(root => {
    const stack = root.closest<HTMLElement>('.document-stack')!, h = Number(stack.dataset.pageHeight), top = Number(stack.dataset.margin), bottom = Number(stack.dataset.bottomMargin), stride = h + 24, origin = root.getBoundingClientRect().top;
    let overflow = 0; const walker = document.createTreeWalker(root.querySelector('td')!, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) { const r = document.createRange(); r.selectNodeContents(walker.currentNode); for (const b of r.getClientRects()) { const y = b.top - origin, p = Math.max(0, Math.floor((y - top + .5) / stride)); overflow = Math.max(overflow, p * stride + top - y, b.bottom - origin - (p * stride + h - bottom)); } }
    return overflow;
  })).toBeLessThan(1);
  expect(await page.locator('.tiptap td').first().textContent()).toBe('한글 long paragraph '.repeat(280));
});


test('editable page gutters preserve a continuing row minimum height and saved model', async ({ page }) => {
  await page.goto('/');
  const doc = await page.evaluate(async () => {
    const { defaultPage } = await import('/src/types.ts'), { importFile } = await import('/src/document-io.ts'), { parseDocumentHtml } = await import('/src/extensions.ts');
    const paragraphs = Array.from({ length: 42 }, (_, i) => `<p style="line-height:20px;margin:0">ROW ${i}</p>`).join('');
    const loaded = await importFile(new File([`<table><tr data-row-split="allow" style="height:1100px"><td>${paragraphs}</td><td><p>SHORT</p></td></tr></table><p>AFTER</p>`], 'reserve.html'));
    return { ...loaded.document, page: { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10 }, content: parseDocumentHtml(loaded.html!) };
  });
  await page.locator('input[type=file]').first().setInputFiles({ name: 'reserve.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(doc)) });
  await expect.poll(() => page.locator('.tiptap tr.pagination-split-row').evaluateAll(rows => {
    const row = rows[0]; if (!row) return 0;
    // Count the row's visible space on paper, including repeated cell padding.
    // A cell widget contains that padding as well as the physical paper gutter;
    // subtracting its entire height incorrectly removes part of the form.
    const root = row.closest<HTMLElement>('.tiptap')!, sheet = root.closest<HTMLElement>('.document-stack')!;
    const origin = root.getBoundingClientRect().top, box = row.getBoundingClientRect();
    const height = Number(sheet.dataset.pageHeight), top = Number(sheet.dataset.margin), bottom = Number(sheet.dataset.bottomMargin);
    let visible = 0;
    for (let page = 0; page < Number(sheet.dataset.pages); page++) {
      const start = origin + page * (height + 24);
      visible += Math.max(0, Math.min(box.bottom, start + height - bottom) - Math.max(box.top, start + top));
    }
    return visible;
  })).toBeGreaterThanOrEqual(1099.5);
  const download = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
  const saved = JSON.parse(await (await import('node:fs/promises')).readFile((await (await download).path())!, 'utf8'));
  expect(saved.content.content[0].content[0].attrs.rowHeight).toBe(1100);
  expect(JSON.stringify(saved.content)).not.toContain('pagination-row-height');
});

test('continuing vertical merges keep their grid editable and ink inside pages', async ({ page }) => {
  await page.goto('/');
  const doc = await page.evaluate(async () => {
    const { defaultPage } = await import('/src/types.ts'), { importFile } = await import('/src/document-io.ts'), { parseDocumentHtml } = await import('/src/extensions.ts');
    const rows = Array.from({ length: 6 }, (_, i) => `<tr style="height:200px">${i === 0 ? '<td rowspan="6" data-cell-split="allow" style="vertical-align:middle"><p>MERGED ONCE</p></td>' : ''}<td><p>ROW ${i}</p></td></tr>`).join('');
    const loaded = await importFile(new File([`<table>${rows}</table><p>AFTER</p>`], 'merge.html'));
    return { ...loaded.document, page: { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10 }, content: parseDocumentHtml(loaded.html!) };
  });
  await page.locator('input[type=file]').first().setInputFiles({ name: 'merge.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(doc)) });
  for (const zoom of [100, 200, 50, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => page.locator('.tiptap').evaluate((root, zoom) => {
      const stack = root.closest<HTMLElement>('.document-stack')!, h = Number(stack.dataset.pageHeight), top = Number(stack.dataset.margin), bottom = Number(stack.dataset.bottomMargin), stride = h + 24, origin = root.getBoundingClientRect().top;
      let overflow = 0;
      for (const p of root.querySelectorAll('td p')) { const r = document.createRange(); r.selectNodeContents(p); for (const b of r.getClientRects()) { const y = (b.top - origin) / (zoom / 100), page = Math.max(0, Math.floor((y - top + .5) / stride)); overflow = Math.max(overflow, page * stride + top - y, (b.bottom - origin) / (zoom / 100) - (page * stride + h - bottom)); } }
      return overflow;
    }, zoom)).toBeLessThan(1);
  }
  const last = page.locator('.tiptap p').filter({ hasText: /^ROW 5$/ }); await last.fill('EDITED'); await page.keyboard.press('ControlOrMeta+z'); await expect(last).toHaveText('ROW 5');
  const download = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
  const saved = JSON.parse(await (await import('node:fs/promises')).readFile((await (await download).path())!, 'utf8'));
  const table = saved.content.content[0]; expect(table.content).toHaveLength(6); expect(table.content[0].content[0].attrs).toMatchObject({ rowspan: 6, cellSplit: 'allow', verticalAlign: 'middle' });
  expect(table.content.map((r: any) => r.attrs.rowHeight)).toEqual(Array(6).fill(200));
  expect(JSON.stringify(saved.content)).not.toContain('pagination-merge');
});

test('fractional custom paper and consecutive column groups settle before page gaps at every zoom', async ({ page }) => {
  await page.goto('/');
  const doc = await page.evaluate(async () => {
    const { defaultPage } = await import('/src/types.ts'), { importFile } = await import('/src/document-io.ts'), { parseDocumentHtml } = await import('/src/extensions.ts');
    const group = (name: string, count: number) => `<div data-column-count="2" data-column-space="41.6" data-column-balance="false">${Array.from({ length: count }, (_, i) => `<p style="line-height:27px;margin:0 0 6.65px">${name} ${i} 한글 column words</p>`).join('')}</div>`;
    const source = '<div data-column-count="2" data-column-space="41.6"><p style="line-height:22.96px;margin:0"></p></div><div data-column-count="1"><table><tr style="height:160px"><td><p>HEADING</p></td></tr></table></div>' + group('FIRST', 155) + '<div data-page-break="true"></div>' + group('SECOND', 125);
    const loaded = await importFile(new File([source], 'columns.html'));
    return { ...loaded.document, page: { ...defaultPage, size: 'Custom', width: 297, height: 419.99, margin: 31, margins: { top: 56.5, right: 31, bottom: 43.5, left: 31 } }, content: parseDocumentHtml(loaded.html!) };
  });
  await page.locator('input[type=file]').first().setInputFiles({ name: 'columns.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(doc)) });
  let pages = '';
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => page.locator('.tiptap').evaluate((root, zoom) => {
      const stack = root.closest<HTMLElement>('.document-stack')!, h = Number(stack.dataset.pageHeight), top = Number(stack.dataset.margin), bottom = Number(stack.dataset.bottomMargin), origin = root.getBoundingClientRect().top, stride = h + 24;
      let bad = 0; const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) { const n = walker.currentNode; if (!n.textContent?.trim() || n.parentElement?.closest('.page-break')) continue; const range = document.createRange(); range.selectNodeContents(n); for (const b of range.getClientRects()) { const y = (b.top - origin) / (zoom / 100), p = Math.max(0, Math.floor((y - top + .5) / stride)); if (y < p * stride + top - 1 || (b.bottom - origin) / (zoom / 100) > p * stride + h - bottom + 1) bad++; } }
      return bad;
    }, zoom)).toBe(0);
    const current = await page.locator('.document-stack').getAttribute('data-pages'); if (!pages) pages = current!; else expect(current).toBe(pages);
  }
  await expect(page.locator('.tiptap')).toContainText('SECOND 124');
});

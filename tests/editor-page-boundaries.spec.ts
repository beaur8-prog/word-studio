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

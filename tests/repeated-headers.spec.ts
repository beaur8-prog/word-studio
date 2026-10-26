import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';

for (const kind of ['ordinary', 'continuing', 'reserved', 'nested', 'nested-form', 'nested-reserve']) test(`repeated headers share page starts with print for ${kind} rows, without changing saved content`, async ({ page }) => {
  await page.goto('/');
  const state = await page.evaluate(async kind => {
    const { importFile } = await import('/src/document-io.ts'), { defaultPage } = await import('/src/types.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts'), { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const paras = Array.from({ length: 100 }, (_, i) => `<p style="font-size:10pt;line-height:20px;margin:0">ROW ${i}</p>`);
    const rows = kind === 'reserved' ? `<tr data-row-split="allow" style="height:1300px"><td>${paras[0]}</td><td><p>ONE CELL</p></td></tr><tr><td>${paras[1]}</td><td><p>AFTER RESERVED SPACE</p></td></tr>`
      : kind === 'nested-form' || kind === 'nested-reserve' ? `<tr data-row-split="allow"${kind === 'nested-reserve' ? ' style="height:1400px"' : ''}><td>${paras.slice(0, 15).join('')}<table><tr><td>${paras.slice(15, 35).join('')}</td></tr></table>${paras.slice(35, 60).join('')}</td><td><p>ONE CELL</p></td></tr>`
      : kind === 'nested' ? `<tr data-row-split="allow"><td>${paras[0]}<table>${[0, 1, 2].map(i => `<tr><td>${paras.slice(1 + i * 20, 21 + i * 20).join('')}</td></tr>`).join('')}</table>${paras[61]}</td><td><p>ONE CELL</p></td></tr>`
      : kind === 'continuing' ? `<tr data-row-split="allow"><td>${paras.join('')}</td><td><p>ONE CELL</p></td></tr>` : paras.map(p => `<tr style="height:32px"><td>${p}</td><td><p>DATA</p></td></tr>`).join('');
    const html = `<table style="margin:0;width:400px"><colgroup><col style="width:200px"><col style="width:200px"></colgroup><tr style="height:40px"><th><p>REPEAT HEADER</p></th><th><p>OTHER HEADER</p></th></tr>${rows}</table><p></p>`;
    const imported = await importFile(new File([html], 'headers.html'));
    const doc = { ...imported.document, page: { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10 }, content: parseDocumentHtml(imported.html!) };
    const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
    try { return { doc, pages: layout.pages, first: [...layout.element.querySelectorAll('.print-content')].map(p => [...p.querySelectorAll('td p')].find(p => /^ROW /.test(p.textContent!))?.textContent) }; }
    finally { layout.dispose(); }
  }, kind);
  await page.locator('input[type=file]').first().setInputFiles({ name: 'repeat.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(state.doc)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe(String(state.pages));
    await expect.poll(() => page.locator('.pagination-headers table').count()).toBe(state.pages - 1);
    await expect.poll(() => editor.evaluate((root, zoom) => {
      const sheet = root.closest<HTMLElement>('.document-stack')!, origin = root.getBoundingClientRect().top;
      const stride = Number(sheet.dataset.pageHeight) + 24, margin = Number(sheet.dataset.margin), first: string[] = [];
      for (const p of root.querySelectorAll('td p')) {
        if (!/^ROW /.test(p.textContent!)) continue;
        const range = document.createRange(); range.selectNodeContents(p); const r = range.getBoundingClientRect();
        const index = Math.max(0, Math.floor(((r.top - origin) / (zoom / 100) - margin + .5) / stride));
        first[index] ??= p.textContent!;
      }
      return first;
    }, zoom)).toEqual(state.first);
  }
  if (kind === 'nested-form' || kind === 'nested-reserve') {
    // A pagination TR must not become a source row on the next measurement.
    // Observe many frames: a one-off assertion can pass halfway through a loop.
    const positions = await editor.evaluate(async root => {
      const snapshots = [];
      for (let frame = 0; frame < 24; frame++) {
        await new Promise(requestAnimationFrame);
        snapshots.push([...root.querySelectorAll('td p')].map(p => Math.round(p.getBoundingClientRect().top * 100)).join(','));
      }
      return snapshots;
    });
    expect(new Set(positions).size).toBe(1);
  }
  const header = editor.locator('th p').first(); await header.fill('EDITED HEADER');
  await expect(page.locator('.pagination-headers').getByText('EDITED HEADER', { exact: true })).toHaveCount(state.pages - 1);
  await page.keyboard.press('ControlOrMeta+z'); await expect(header).toHaveText('REPEAT HEADER');
  const downloaded = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
  const saved = JSON.parse(await fs.readFile((await (await downloaded).path())!, 'utf8'));
  expect(saved.content).toEqual(state.doc.content);
});

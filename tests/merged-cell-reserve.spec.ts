import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';

for (const nested of [false, true]) for (const alignment of ['top', 'middle', 'bottom']) test(`${alignment}${nested ? ' nested' : ''}: crossing merges grow a short continuation row for repeated padding without changing saved minimum heights`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ alignment, nested }) => {
    const { defaultPage } = await import('/src/types.ts'), { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts'), { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const p = (text: string) => `<p style="font-size:10pt;line-height:21px;margin:0">${text}</p>`;
    const cell = `style="padding:2px;vertical-align:${alignment}"`;
    const contents = nested ? p('BEFORE NESTED') + `<table style="width:170px;margin:0"><tr style="height:30px"><td style="padding:0">${p('NESTED ONE')}</td></tr><tr style="height:30px"><td style="padding:0">${p('NESTED TWO')}</td></tr></table>` + Array.from({ length: 6 }, (_, i) => p('MERGED ' + i)).join('')
      : Array.from({ length: 10 }, (_, i) => p('MERGED ' + i)).join('');
    const html = '<p style="font-size:10pt;line-height:420px;margin:0">LEAD</p>' +
      `<table style="width:400px;margin:0"><colgroup><col style="width:100px"><col style="width:180px"><col style="width:120px"></colgroup>` +
      `<tr style="height:70px"><td rowspan="4" data-cell-split="allow" ${cell}>${p('LABEL')}${p('TAIL')}</td><td rowspan="3" data-cell-split="allow" ${cell}>${contents}</td><td ${cell}>${p('ROW 0')}</td></tr>` +
      `<tr style="height:70px"><td ${cell}>${p('ROW 1')}</td></tr>` +
      `<tr style="height:75px"><td ${cell}>${p('ROW 2')}</td></tr>` +
      `<tr style="height:270px"><td ${cell}>${p('AFTER MERGE')}</td><td ${cell}>${p('ROW 3')}</td></tr></table>`;
    const loaded = await importFile(new File([html], 'reserve.html'));
    let doc = { ...loaded.document, page: { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10, pageNumbers: false }, content: parseDocumentHtml(loaded.html!) };
    const first = doc, states = [];
    for (let save = 0; save < 3; save++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try { states.push([...layout.element.querySelectorAll('.print-content')].map(p => p.textContent!.replace(/\s/g, ''))); }
      finally { layout.dispose(); }
      if (save < 2) { const again = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...again.document, content: parseDocumentHtml(again.html!) }; }
    }
    return { first, states };
  }, { alignment, nested });
  expect(result.states[0]).toHaveLength(2);
  expect(result.states[0][0]).toContain('ROW1'); expect(result.states[0][0]).not.toContain('ROW2');
  expect(result.states[1]).toEqual(result.states[0]); expect(result.states[2]).toEqual(result.states[0]);
  await page.locator('input[type=file]').first().setInputFiles({ name: 'reserve.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.first)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => editor.evaluate((root, zoom) => {
      const sheet = root.closest<HTMLElement>('.document-stack')!, origin = root.getBoundingClientRect().top, stride = Number(sheet.dataset.pageHeight) + 24, pages: string[] = [];
      for (const p of root.querySelectorAll('p')) {
        const at = Math.floor(((p.getBoundingClientRect().top - origin) / (zoom / 100) + .5) / stride);
        pages[at] = (pages[at] || '') + p.textContent!.replace(/\s/g, '');
      }
      return pages;
    }, zoom)).toEqual(result.states[0]);
  }
  await editor.locator(':scope > p').first().fill('EDIT'); await page.keyboard.press('ControlOrMeta+z'); await expect(editor.locator(':scope > p').first()).toHaveText('LEAD');
  const download = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
  const saved = JSON.parse(await fs.readFile((await (await download).path())!, 'utf8'));
  expect(saved.content).toEqual(result.first.content);
});

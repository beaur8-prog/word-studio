import { test, expect } from '@playwright/test';
import { createRequire } from 'node:module';
import fs from 'node:fs/promises';
const fixture = createRequire(import.meta.url)('./hwp-inline-wrap-fixture.cjs');
let samples: Record<string, any>;
test.beforeAll(async () => { samples = await fixture(); });
for (const format of ['hwp', 'hwpx']) for (const placement of ['top', 'edge', 'large-text-edge'])
test(`${format}: multiline inline label ${placement} through two DOCX saves`, async ({ page }, info) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format, placement }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const reader = new HwpReaderSession(); let doc;
    try { await reader.open(new File([new Uint8Array(bytes)], 'wrap.' + format)); doc = importEditableHwp(await reader.editable(), 'wrap.' + format); }
    finally { reader.close(); }
    const edge = placement !== 'top';
    doc.page = { ...doc.page, size: 'Custom', width: 180, height: 150, margin: 10, margins: { top: 10, left: 10, right: 10, bottom: 10 }, header: '', footer: '', headerHtml: '', footerHtml: '', pageNumbers: false };
    if (edge) doc.content.content![0].attrs = { ...doc.content.content![0].attrs, lineSpacing: `${130 * 96 / 25.4 - (placement === 'large-text-edge' ? 36 : 18)}px`, hwpBodyLine: null };
    if (placement === 'large-text-edge') {
      const tail = doc.content.content!.find(node => node.type === 'paragraph' && node.content?.some(text => text.text?.includes('TAIL START')))!;
      tail.attrs = { ...tail.attrs, paragraphSize: '40px' };
      for (const run of tail.content || []) { const style = run.marks?.find(mark => mark.type === 'textStyle'); if (style) style.attrs = { ...style.attrs, fontSize: '40px' }; }
    }
    const first = structuredClone(doc), states = [], files = [];
    for (let save = 0; save < 3; save++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try {
        const table = [...layout.element.querySelectorAll('table')].find(t => t.textContent?.includes('INLINE LABEL'))!, box = table.getBoundingClientRect();
        const paragraphs = [...layout.element.querySelectorAll<HTMLElement>('.print-content > p')].filter(p => p.textContent?.includes('TAIL START'));
        const p = paragraphs[0], walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT), lines = [];
        while (walker.nextNode()) { const r = document.createRange(); r.selectNodeContents(walker.currentNode); lines.push(...[...r.getClientRects()].filter(r => r.width > .1).map(r => ({ left: r.left, right: r.right, top: r.top, bottom: r.bottom }))); }
        const after = [...layout.element.querySelectorAll<HTMLElement>('.print-content > p')].find(p => p.textContent === 'AFTER')!;
        states.push({ save, pages: layout.pages, text: [...layout.element.querySelectorAll('.print-content')].map(p => p.textContent).join(''),
          table: box.toJSON(), lines, bodyLeft: p.closest('.print-content')!.getBoundingClientRect().left,
          tablePage: table.closest<HTMLElement>('.print-page')!.dataset.page, tailPage: p.closest<HTMLElement>('.print-page')!.dataset.page,
          afterTop: after.getBoundingClientRect().top, warnings: layout.warnings });
      } finally { layout.dispose(); }
      if (save < 2) { const blob = await exportDocx(doc); files.push([...new Uint8Array(await blob.arrayBuffer())]); const loaded = await importFile(new File([blob], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
    }
    return { first, states, files };
  }, { bytes: [...samples[format].bytes], format, placement });
  const edge = placement !== 'top';
  for (const [i, bytes] of result.files.entries()) await fs.writeFile(info.outputPath(`hwp-gov-wrap-${format}-${placement}-save${i+1}.docx`), Buffer.from(bytes));
  for (const state of result.states) {
    if (placement !== 'large-text-edge') expect(state.pages, JSON.stringify(result.states)).toBe(edge ? 2 : 1);
    else expect(state.pages).toBeGreaterThanOrEqual(2);
    expect(state.text).toBe(result.states[0].text); expect(state.warnings).toEqual([]);
    expect(state.tablePage).toBe(state.tailPage); expect(state.tablePage).toBe(edge ? '2' : '1');
    expect(state.lines.length).toBeGreaterThan(2);
    expect(state.lines[0].left, JSON.stringify(state)).toBeGreaterThanOrEqual(state.table.right - 1);
    expect(state.lines[0].top, JSON.stringify(state)).toBeLessThan(state.table.bottom);
    expect(state.lines.some(line => line.top >= state.table.bottom && line.left < state.table.right - 20), JSON.stringify(state)).toBe(true);
    expect(state.afterTop).toBeGreaterThanOrEqual(Math.max(...state.lines.map(line => line.bottom), state.table.bottom) - 1);
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'wrap.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.first)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' }), tail = editor.locator(':scope > p').filter({ hasText: 'TAIL START' });
  await expect(tail).toBeVisible();
  const measure = () => tail.evaluate(p => { const table = p.parentElement!.querySelector(':scope > .tableWrapper table')!, r = document.createRange(); r.selectNodeContents(p); return { first: [...r.getClientRects()].find(r => r.width > .1)!.toJSON(), table: table.getBoundingClientRect().toJSON(), height: p.getBoundingClientRect().height }; });
  await expect.poll(async () => (await measure()).first.left - (await measure()).table.right).toBeGreaterThanOrEqual(-1);
  await expect.poll(async () => { const state = await measure(); return state.first.top - state.table.top; }).toBeLessThan(70);
  const before = await tail.textContent(), height = (await measure()).height;
  await tail.fill('TAIL START ' + 'EDIT GROW '.repeat(100)); await expect.poll(async () => (await measure()).height).toBeGreaterThan(height);
  await page.keyboard.press('ControlOrMeta+z'); await expect(tail).toHaveText(before!);
  await expect.poll(async () => (await measure()).height).toBeCloseTo(height, 0);
  await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(tail).toContainText('EDIT GROW');
});

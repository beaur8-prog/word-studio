import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-merge-row-break-fixture.cjs');
let directory: string, files: Record<string, string>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-hwp-merge-')); files = await fixture(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

for (const format of ['hwp', 'hwpx']) for (const policy of [0, 2]) test(`${format} policy ${policy}: a formerly single-page vertical merge respects row-boundary permission after preceding flow changes`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts'), { createPrintLayout } = await import('/src/print-layout.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'merge.' + format));
      let doc = importEditableHwp(await session.editable(), 'merge.' + format);
      // A paper/preceding-flow edit creates a boundary that did not exist in
      // the source. Permission comes from the source policy, not old boxes.
      doc.page = { ...doc.page, size: 'Custom', width: 150, height: 180, margin: 10, margins: undefined, pageNumbers: false };
      const lead = doc.content.content!.find(n => n.type === 'paragraph' && n.content?.some(t => t.text === 'LEAD'))!;
      lead.attrs = { ...lead.attrs, lineSpacing: '350px', spaceBefore: '0px', spaceAfter: '0px' };
      const first = doc, states = [];
      for (let save = 0; save < 3; save++) {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try { states.push({ pages: [...layout.element.querySelectorAll('.print-content')].map(p => p.textContent!.replace(/\s/g, '')), scaled: [...layout.element.querySelectorAll<HTMLElement>('[style]')].some(p => !!p.style.zoom) }); }
        finally { layout.dispose(); }
        if (save < 2) { const again = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...again.document, content: parseDocumentHtml(again.html!) }; }
      }
      return { first, states };
    } finally { session.close(); }
  }, { bytes: [...await fs.readFile(files[`${format}-${policy}`])], format });
  for (const state of result.states) {
    expect(state.scaled).toBe(false); expect(state.pages).toHaveLength(2);
    expect(state.pages.join('')).toBe('LEADMERGEDONCEROWONEROWTWO');
    expect(state.pages[0].includes('ROWONE')).toBe(policy === 2); expect(state.pages[0]).not.toContain('ROWTWO');
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'merge.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.first)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => editor.locator('p').filter({ hasText: 'ROW ONE' }).evaluate((p, zoom) => {
      const sheet = p.closest<HTMLElement>('.document-stack')!, y = (p.getBoundingClientRect().top - sheet.getBoundingClientRect().top) / (zoom / 100);
      return Math.floor(y / (Number(sheet.dataset.pageHeight) + 24));
    }, zoom)).toBe(policy === 2 ? 0 : 1);
  }
  await editor.locator(':scope > p').first().fill('EDIT'); await page.keyboard.press('ControlOrMeta+z'); await expect(editor.locator(':scope > p').first()).toHaveText('LEAD');
});

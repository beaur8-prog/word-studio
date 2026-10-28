import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-empty-lines-fixture.cjs');
let directory: string, files: Record<string, any>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-empty-lines-')); files = await fixture(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });
for (const format of ['hwp', 'hwpx']) for (const enabled of [false, true]) test(`${format} hideEmptyLine=${enabled}: only two overflow blank lines hide, preserving edits and two DOCX saves`, async ({ page }) => {
  await page.goto('/'); const file = files[`${format}-${enabled}`];
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp, exportEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'lines.' + format)); let doc = importEditableHwp(await session.editable(), 'lines.' + format); const first = doc, states = [];
      for (let save = 0; save < 3; save++) {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const after = [...layout.element.querySelectorAll('.print-content > p')].find(p => p.textContent === 'AFTER')!;
          states.push({ pages: layout.pages, policy: doc.page.hideEmptyLines === true, paragraphs: doc.content.content!.length,
            hidden: layout.element.querySelectorAll('.pagination-hidden-empty-line').length, afterPage: Number(after.closest<HTMLElement>('.print-page')!.dataset.page),
            afterY: after.getBoundingClientRect().top - after.closest('.print-page')!.getBoundingClientRect().top });
        } finally { layout.dispose(); }
        if (save < 2) { const again = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...again.document, content: parseDocumentHtml(again.html!) }; }
      }
      const nativePolicies = [];
      for (let save = 0; save < 2; save++) {
        const saved = await exportEditableHwp(doc, format); const reopened = new HwpReaderSession();
        try { await reopened.open(new File([saved.blob], 'again.' + format)); doc = importEditableHwp(await reopened.editable(), 'again.' + format); nativePolicies.push(doc.page.hideEmptyLines === true); }
        finally { reopened.close(); }
      }
      return { first, states, nativePolicies };
    } finally { session.close(); }
  }, { bytes: [...await fs.readFile(file.path)], format });
  if (process.env.HWP_LAYOUT_DIAGNOSTICS) console.log(JSON.stringify({ enabled, format, file, states: result.states, nativePolicies: result.nativePolicies }));
  for (const state of result.states) {
    expect(state.pages).toBe(2); expect(state.policy).toBe(enabled); expect(state.paragraphs).toBe(34); expect(state.hidden).toBe(enabled ? 2 : 0);
    expect(state.afterPage).toBe(2); expect(state.afterY).toBeCloseTo(2835 / 75 + (enabled ? 20 : 60), 0);
  }
  expect(result.nativePolicies).toEqual([enabled, enabled]);
  await page.locator('input[type=file]').first().setInputFiles({ name: 'lines.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.first)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' }), after = editor.locator(':scope > p').filter({ hasText: /^AFTER$/ });
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => editor.locator('.pagination-hidden-empty-line').count()).toBe(enabled ? 2 : 0);
    await expect.poll(() => after.evaluate((p, zoom) => (p.getBoundingClientRect().top - p.closest('.document-stack')!.getBoundingClientRect().top) / (zoom / 100), zoom)).toBeCloseTo((51024 / 75 + 24) + result.states[0].afterY, 0);
  }
  // Hidden blank lines have no click target. Place the caret at their retained
  // model position and exercise the real keyboard input/history path.
  const blank = editor.locator(':scope > p').nth(30);
  await blank.evaluate(p => {
    (p.closest('[contenteditable="true"]') as HTMLElement).focus();
    const range = document.createRange(); range.selectNodeContents(p); range.collapse(true);
    const selection = getSelection()!; selection.removeAllRanges(); selection.addRange(range);
  });
  if (enabled) await editor.evaluate(p => p.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' })));
  await page.keyboard.type('VISIBLE'); await expect(blank).toHaveText('VISIBLE');
  await expect.poll(() => blank.evaluate(p => p.classList.contains('pagination-hidden-empty-line'))).toBe(false);
  await expect.poll(() => blank.evaluate(p => p.getBoundingClientRect().height)).toBeGreaterThan(0);
  if (enabled) await editor.evaluate(p => p.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: 'VISIBLE' })));
  await page.keyboard.press('ControlOrMeta+z'); await expect(blank).toBeEmpty();
  await expect.poll(() => editor.locator('.pagination-hidden-empty-line').count()).toBe(enabled ? 2 : 0);
  await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(blank).toHaveText('VISIBLE');
});

for (const kind of ['highlight', 'spacing']) test(`explicit hide-empty policy preserves ${kind} on an otherwise blank paragraph`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async kind => {
    const { importFile } = await import('/src/document-io.ts'), { defaultPage } = await import('/src/types.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts'), { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const lead = Array.from({ length: 30 }, (_, i) => `<p style="font-size:10pt;line-height:20px;margin:0">LEAD ${i}</p>`).join('');
    const blank = kind === 'highlight' ? '<p style="font-size:10pt;line-height:20px;margin:0"><mark> </mark></p>' : '<p style="font-size:10pt;line-height:20px;margin:0 0 8px"></p>';
    const html = lead + blank + '<p style="font-size:10pt;line-height:20px;margin:0">AFTER</p>';
    const loaded = await importFile(new File([html], 'authored.html'));
    const doc = { ...loaded.document, page: { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10, pageNumbers: false, hideEmptyLines: true }, content: parseDocumentHtml(html) };
    const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
    try {
      const after = [...layout.element.querySelectorAll('.print-content > p')].find(p => p.textContent === 'AFTER')!;
      return { doc, pages: layout.pages, hidden: layout.element.querySelectorAll('.pagination-hidden-empty-line').length,
        afterY: after.getBoundingClientRect().top - after.closest('.print-page')!.getBoundingClientRect().top };
    } finally { layout.dispose(); }
  }, kind);
  expect(result.pages).toBe(2); expect(result.hidden).toBe(0); expect(result.afterY).toBeCloseTo(10 * 96 / 25.4 + 20 + (kind === 'spacing' ? 8 : 0), 0);
  await page.locator('input[type=file]').first().setInputFiles({ name: 'authored.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.doc)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  await expect(editor.locator('.pagination-hidden-empty-line')).toHaveCount(0);
  await expect.poll(() => editor.locator(':scope > p').last().evaluate(p => {
    const sheet = p.closest<HTMLElement>('.document-stack')!;
    return p.getBoundingClientRect().top - sheet.getBoundingClientRect().top - (Number(sheet.dataset.pageHeight) + 24);
  })).toBeCloseTo(result.afterY, 0);
});

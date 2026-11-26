import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createRequire } from 'node:module';

test('an overflowed empty column line can share the next authored page', async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async () => {
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { defaultPage } = await import('/src/types.ts');
    const settings = { ...defaultPage, size: 'Custom' as const, width: 132, height: 41.75,
      margin: 5, pageNumbers: false };
    const outcomes = [];
    for (const { hideEmptyLines, decorated } of [
      { hideEmptyLines: true, decorated: false },
      { hideEmptyLines: false, decorated: false },
      { hideEmptyLines: true, decorated: true },
    ]) {
      const columns = '<div class="document-columns" data-column-count="2" data-column-space="12" data-column-balance="false" style="column-gap:12px;column-fill:auto">';
      const text = (value: string) => '<p style="line-height:120px;font-size:12px;margin:0">' + value + '</p>';
      const blank = '<p style="line-height:18px;font-size:12px;margin:0' +
        (decorated ? ';background-color:rgb(255,230,230)' : '') + '"></p>';
      const html = columns + text('LEFT') + '<div data-column-break="true"></div>' + text('RIGHT') + blank +
        '</div><div data-page-break="true"></div>' + columns + '<p>NEXT</p></div>';
      const layout = await createPrintLayout(html, { ...settings, hideEmptyLines });
      try {
        outcomes.push({
          pages: [...layout.element.querySelectorAll('.print-content')].map(content => content.textContent),
          hidden: layout.element.querySelectorAll('p.pagination-hidden-empty-line').length,
          retainedEmpty: [...layout.element.querySelectorAll('.print-content p')].filter(p => p.textContent === '').length,
        });
      } finally { layout.dispose(); }
    }
    return outcomes;
  });
  expect(states).toEqual([
    { pages: ['LEFTRIGHT', 'NEXT'], hidden: 1, retainedEmpty: 1 },
    { pages: ['LEFTRIGHT', '', 'NEXT'], hidden: 0, retainedEmpty: 1 },
    { pages: ['LEFTRIGHT', '', 'NEXT'], hidden: 0, retainedEmpty: 1 },
  ]);
});

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
  await expect(editor.locator(':scope > p').last()).toHaveText('AFTER');
  await expect(editor.locator('.pagination-hidden-empty-line')).toHaveCount(0);
  await expect.poll(() => editor.locator(':scope > p').last().evaluate(p => {
    const sheet = p.closest<HTMLElement>('.document-stack')!;
    return p.getBoundingClientRect().top - sheet.getBoundingClientRect().top - (Number(sheet.dataset.pageHeight) + 24);
  })).toBeCloseTo(result.afterY, 0);
});

test('the empty line after two suppressed lines can retain its em at the page bottom', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { defaultPage } = await import('/src/types.ts');
    const { importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const lead = Array.from({ length: 30 }, (_, i) => `<p style="font-size:10pt;line-height:20px;margin:0">LEAD ${i}</p>`).join('');
    const blank = (advance: number) => `<p style="font-size:10pt;line-height:${advance}px;margin:0"></p>`;
    const layoutPage = { ...defaultPage, size: 'Custom' as const, width: 150, height: 20 + 615 * 25.4 / 96,
      margin: 10, pageNumbers: false, hideEmptyLines: true };
    const states = [];
    for (const short of [true, false]) {
      const html = lead + blank(26) + blank(26) + blank(short ? 20 : 40) +
        '<p style="font-size:10pt;line-height:20px;margin:0">AFTER</p>';
      const { document } = await importFile(new File([html], 'core-blank-fit.html'));
      const doc = { ...document, page: layoutPage, content: parseDocumentHtml(html) };
      const output = await createPrintLayout(generateHTML(doc.content, editorExtensions), layoutPage);
      try {
        const sheet = output.element.querySelector<HTMLElement>('.print-page[data-page="2"]')!;
        const after = [...sheet.querySelectorAll('p')].find(p => p.textContent === 'AFTER')!;
        states.push({ doc, firstParagraphs: output.element.querySelectorAll('.print-page[data-page="1"] .print-content > p').length,
          hidden: output.element.querySelectorAll('.pagination-hidden-empty-line').length,
          after: after.getBoundingClientRect().top - sheet.getBoundingClientRect().top });
      } finally { output.dispose(); }
    }
    return states;
  });
  const [retained, spacer] = result;
  expect(retained.firstParagraphs).toBe(33);
  expect(spacer.firstParagraphs).toBe(32);
  expect(retained.hidden).toBe(2);
  expect(spacer.hidden).toBe(2);
  expect(spacer.after - retained.after).toBeCloseTo(40, 1);
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'native-blank-fit.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(retained.doc)),
  });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  await expect(editor.locator(':scope > p')).toHaveCount(34);
  await expect(editor.locator('.pagination-hidden-empty-line')).toHaveCount(2);
  await expect.poll(() => editor.locator(':scope > p').last().evaluate(p => {
    const sheet = p.closest<HTMLElement>('.document-stack')!;
    return p.getBoundingClientRect().top - sheet.getBoundingClientRect().top - Number(sheet.dataset.pageHeight) - 24;
  })).toBeCloseTo(retained.after, 1);
});

test('a verified HWP terminal empty line preserves its full advance through two DOCX saves', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { defaultPage } = await import('/src/types.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const paper = { ...defaultPage, size: 'Custom' as const, width: 150,
      height: 20 + 615 * 25.4 / 96, margin: 10, pageNumbers: false, suppressBottomSpacing: true };
    const style = 'style="font-size:10pt;line-height:20px;margin:0"';
    const lead = Array.from({ length: 30 }, (_, i) => `<p ${style}>LEAD ${i}</p>`).join('');
    const marked = `<p ${style} data-hwp-body-line="6 150 13.3333 20"></p>`;
    const after = `<p ${style}>AFTER</p>`;
    const html = lead + marked + after;
    const { document } = await importFile(new File([html], 'verified-terminal-line.html'));
    let doc = { ...document, page: paper, content: parseDocumentHtml(html) };
    const initial = doc;
    async function observe(content: typeof doc.content) {
      const layout = await createPrintLayout(generateHTML(content, editorExtensions), paper);
      try {
        const paragraphs = [...layout.element.querySelectorAll('.print-content > p')];
        const empty = paragraphs.find(p => !p.textContent)!;
        const next = paragraphs.find(p => p.textContent === 'AFTER')!;
        const sheet = next.closest<HTMLElement>('.print-page')!;
        return { emptyPage: empty.closest<HTMLElement>('.print-page')!.dataset.page,
          nextPage: sheet.dataset.page, afterY: next.getBoundingClientRect().top - sheet.getBoundingClientRect().top,
          blankHeight: empty.getBoundingClientRect().height, paragraphs: paragraphs.length,
          hidden: layout.element.querySelectorAll('.pagination-hidden-empty-line').length };
      } finally { layout.dispose(); }
    }
    const states = [];
    for (let save = 0; save < 3; save++) {
      states.push(await observe(doc.content));
      if (save < 2) {
        const second = await importFile(new File([await exportDocx(doc)], 'saved.docx'));
        doc = { ...second.document, content: parseDocumentHtml(second.html!) };
      }
    }
    const ordinary = await observe(parseDocumentHtml(lead + `<p ${style}></p>` + after));
    const edited = await observe(parseDocumentHtml(lead + marked.replace('</p>', 'TYPED</p>') + `<p ${style}></p>` + after));
    return { initial, states, ordinary, edited };
  });
  for (const state of result.states) {
    expect(state).toMatchObject({ emptyPage: '1', nextPage: '2', paragraphs: 32, hidden: 0 });
    expect(state.blankHeight).toBeCloseTo(20, 1);
    expect(result.ordinary.afterY - state.afterY).toBeCloseTo(20, 1);
  }
  expect(result.ordinary.emptyPage).toBe('2');
  expect(result.edited.emptyPage).toBe('2');
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'verified-terminal-line.wordstudio', mimeType: 'application/json',
    buffer: Buffer.from(JSON.stringify(result.initial)),
  });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  await expect(editor.locator(':scope > p')).toHaveCount(32);
  await expect.poll(() => editor.locator(':scope > p').last().evaluate(p => {
    const sheet = p.closest<HTMLElement>('.document-stack')!;
    return p.getBoundingClientRect().top - sheet.getBoundingClientRect().top - Number(sheet.dataset.pageHeight) - 24;
  })).toBeCloseTo(result.states[0].afterY, 1);
});

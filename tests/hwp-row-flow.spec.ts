import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
const createFixture = createRequire(import.meta.url)('./hwp-row-flow-fixture.cjs');
let directory: string, files: Record<string, string>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.resolve('.cache/hwp-row-flow-')); files = await createFixture(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

for (const format of ['hwp', 'hwpx']) test(`${format}: long two-column rows retain each column, widths and full-size text across pages and two DOCX saves`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'rows.' + format));
      let doc = importEditableHwp(await session.editable(), 'rows.' + format);
      const states = [];
      for (let save = 0; save < 3; save++) {
        const html = generateHTML(doc.content, editorExtensions), layout = await createPrintLayout(html, doc.page);
        try {
          const sheets = [...layout.element.querySelectorAll('.print-page')];
          const columns = ['', ''], widths: number[][] = [];
          let maxOverflow = 0, headers = 0;
          for (const sheet of sheets) {
            const boundary = sheet.querySelector('.print-content')!.getBoundingClientRect().bottom;
            for (const row of sheet.querySelectorAll<HTMLTableRowElement>('.print-content tr')) {
              maxOverflow = Math.max(maxOverflow, row.getBoundingClientRect().bottom - boundary);
              if ([...row.cells].every(c => c.tagName === 'TH')) { headers++; continue; }
              widths.push([...row.cells].map(c => c.getBoundingClientRect().width));
              [...row.cells].forEach((cell, i) => { columns[i] += cell.textContent; });
            }
          }
          states.push({ pages: layout.pages, columns, widths, maxOverflow, headers,
            scaled: [...layout.element.querySelectorAll<HTMLElement>('[style]')].some(e => !!e.style.zoom),
            policies: JSON.stringify(doc.content).match(/"rowSplit":"allow"/g)?.length ?? 0 });
        } finally { layout.dispose(); }
        if (save < 2) { const reopened = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...reopened.document, content: parseDocumentHtml(reopened.html!) }; }
      }
      return states;
    } finally { session.close(); }
  }, { bytes: [...await fs.readFile(files[format])], format });
  for (const state of result) {
    expect(state.pages).toBeGreaterThan(1); expect(state.pages).toBeLessThan(5);
    expect(state.policies).toBe(1); expect(state.headers).toBe(state.pages);
    expect(state.maxOverflow).toBeLessThanOrEqual(.5); expect(state.scaled).toBe(false);
    for (const [column, count] of [[0, 110], [1, 75]]) {
      const expected = Array.from({ length: count }, (_, i) => `${column === 0 ? 'LEFT' : 'RIGHT'} ${i} 한글 paragraph`).join('');
      expect(state.columns[column].replace(/\s/g, '')).toBe(expected.replace(/\s/g, ''));
    }
    for (const widths of state.widths) { expect(widths).toHaveLength(2); expect(widths[0]).toBeCloseTo(state.widths[0][0], 0); expect(widths[1]).toBeCloseTo(state.widths[0][1], 0); }
  }
  expect(result.map(s => s.pages)).toEqual([result[0].pages, result[0].pages, result[0].pages]);
});

test('allowed rows share the available page height while keeping empty continuation columns and minimum-height reserve', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { defaultPage } = await import('/src/types.ts');
    const line = (text: string) => `<p style="line-height:20px;margin:0">${text}</p>`;
    const row = `<table style="width:400px;margin:0"><colgroup><col style="width:200px"><col style="width:200px"></colgroup><tr data-row-split="allow" style="height:1100px"><td style="padding:0">${line('LEFT ONCE')}</td><td style="padding:0">${Array.from({ length: 45 }, (_, i) => line('RIGHT ' + i)).join('')}</td></tr></table>`;
    const layout = await createPrintLayout(line('BEFORE') + row + line('AFTER'), { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10 });
    try {
      const rows = [...layout.element.querySelectorAll<HTMLTableRowElement>('tr')];
      return { pages: layout.pages, rowHeights: rows.map(r => r.getBoundingClientRect().height), columns: rows.map(r => [...r.cells].map(c => c.textContent)),
        text: layout.element.textContent, scaled: [...layout.element.querySelectorAll<HTMLElement>('[style]')].some(e => !!e.style.zoom) };
    } finally { layout.dispose(); }
  });
  expect(result.pages).toBe(2); expect(result.scaled).toBe(false);
  expect(result.rowHeights.reduce((a, b) => a + b)).toBeCloseTo(1100, 0);
  expect(result.columns).toHaveLength(2); expect(result.columns[0][0]).toBe('LEFT ONCE'); expect(result.columns[1][0]).toBe('');
  expect(result.columns.map(c => c[1]).join('')).toBe(Array.from({ length: 45 }, (_, i) => 'RIGHT ' + i).join(''));
  expect(result.text).toContain('AFTER');
});

test('unmarked forms, exact rows, vertical merges and explicit keep-together rows retain their boundaries and DOCX policy', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { defaultPage } = await import('/src/types.ts');
    const cases = ['', 'data-row-split="avoid"', 'data-row-split="allow" data-word-row-rule="exact"', 'data-row-split="allow"'];
    const output = [];
    for (const [i, attrs] of cases.entries()) {
      const source = `<p style="line-height:20px;margin:0">OPEN${'<br>line'.repeat(20)}</p><table style="margin:0"><tr ${attrs} style="height:300px"><td${i === 3 ? ' rowspan="2"' : ''}><p>KEEP LEFT</p></td><td><p>KEEP RIGHT</p></td></tr>${i === 3 ? '<tr><td><p>MERGED END</p></td></tr>' : ''}</table>`;
      const doc = await importFile(new File([source], 'policy.html'));
      const layout = await createPrintLayout(doc.html!, { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10 });
      try { output.push([...layout.element.querySelectorAll('.print-content')].map(p => p.textContent)); }
      finally { layout.dispose(); }
    }
    const source = await importFile(new File(['<table><tr data-row-split="avoid"><td><p>KEEP</p></td></tr><tr data-row-split="allow"><td><p>SPLIT</p></td></tr></table>'], 'policies.html'));
    let doc = { ...source.document, content: parseDocumentHtml(source.html!) };
    const policies = [];
    for (let save = 0; save < 2; save++) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; policies.push(doc.content.content?.find(n => n.type === 'table')?.content?.map(n => n.attrs?.rowSplit)); }
    return { output, policies };
  });
  for (const pages of result.output) { expect(pages).toHaveLength(2); expect(pages[0]).not.toContain('KEEP'); expect(pages[1]).toContain('KEEP LEFT'); expect(pages[1]).toContain('KEEP RIGHT'); }
  expect(result.policies).toEqual([['avoid', 'allow'], ['avoid', 'allow']]);
});

test('source-observed vertical merges continue without scaling or duplicating cell text', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { createPrintLayout } = await import('/src/print-layout.ts'), { defaultPage } = await import('/src/types.ts');
    const { importFile } = await import('/src/document-io.ts'), { parseDocumentHtml } = await import('/src/extensions.ts');
    const html = `<table style="width:400px;margin:0"><colgroup><col style="width:100px"><col style="width:300px"></colgroup>${Array.from({ length: 6 }, (_, i) => `<tr style="height:200px">${i === 0 ? '<td rowspan="6" data-cell-split="allow" style="vertical-align:middle"><p>MERGED ONCE</p></td>' : ''}<td><p>ROW ${i}</p></td></tr>`).join('')}</table>`;
    const imported = await importFile(new File([html], 'merge.html'));
    const model = parseDocumentHtml(imported.html!);
    const layout = await createPrintLayout(imported.html!, { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10 });
    try {
      return { pages: layout.pages, model, texts: [...layout.element.querySelectorAll('td')].map(c => c.textContent),
        widths: [...layout.element.querySelectorAll<HTMLTableRowElement>('tr')].filter(r => r.cells.length === 2).map(r => [...r.cells].map(c => c.getBoundingClientRect().width)),
        scaled: [...layout.element.querySelectorAll<HTMLElement>('[style]')].some(e => !!e.style.zoom),
        maxOverflow: Math.max(...[...layout.element.querySelectorAll('.print-content')].map(p => Math.max(...[...p.querySelectorAll('tr')].map(r => r.getBoundingClientRect().bottom - p.getBoundingClientRect().bottom)))),
      };
    } finally { layout.dispose(); }
  });
  expect(result.pages).toBe(2); expect(result.scaled).toBe(false); expect(result.maxOverflow).toBeLessThanOrEqual(.5);
  expect(result.texts.filter(t => t?.includes('MERGED ONCE'))).toHaveLength(1);
  for (let i = 0; i < 6; i++) expect(result.texts.filter(t => t === 'ROW ' + i)).toHaveLength(1);
  expect(result.widths).toHaveLength(2); expect(result.widths[0]).toEqual(result.widths[1]);
  expect(result.model.content?.[0].content?.[0].content?.[0].attrs).toMatchObject({ rowspan: 6, cellSplit: 'allow' });
});

test('continuing merge policy and unscaled fragments survive two DOCX saves', async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async () => {
    const { importFile, exportDocx } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const { defaultPage } = await import('/src/types.ts');
    const html = `<table>${Array.from({ length: 6 }, (_, i) => `<tr style="height:200px">${i === 0 ? '<td rowspan="6" data-cell-split="allow"><p>MERGED ONCE</p></td>' : ''}<td><p>ROW ${i}</p></td></tr>`).join('')}</table>`;
    const loaded = await importFile(new File([html], 'merge.html'));
    let doc = { ...loaded.document, page: { ...defaultPage, pageNumbers: false, size: 'Custom', width: 150, height: 180, margin: 10 }, content: parseDocumentHtml(loaded.html!) }; const states = [];
    for (let i = 0; i < 2; i++) {
      const reopened = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...reopened.document, content: parseDocumentHtml(reopened.html!) };
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try { states.push({ model: doc.content, pages: layout.pages, scaled: [...layout.element.querySelectorAll<HTMLElement>('[style]')].some(e => !!e.style.zoom), texts: [...layout.element.querySelectorAll('td')].map(c => c.textContent) }); } finally { layout.dispose(); }
    }
    return states;
  });
  for (const state of states) {
    expect(state.pages).toBe(2); expect(state.scaled).toBe(false); expect(JSON.stringify(state.model)).toContain('"cellSplit":"allow"');
    expect(state.texts.filter(t => t === 'MERGED ONCE')).toHaveLength(1); for (let i = 0; i < 6; i++) expect(state.texts.filter(t => t === 'ROW ' + i)).toHaveLength(1);
  }
});

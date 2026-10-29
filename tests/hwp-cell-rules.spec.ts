import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';

const require = createRequire(import.meta.url), fixture = require('./hwp-cell-reflow-fixture.cjs');
let directory: string, samples: Record<string, any>;
test.beforeAll(async () => {
  directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-cell-rules-'));
  samples = await fixture(directory);
  const { HwpDocument } = await import(pathToFileURL(path.join(path.dirname(require.resolve('@rhwp/core')), 'rhwp.js')).href);
  // The formatting command stores physical leading at ParaShape's 2x scale;
  // the read-only query returns pixels. Verify the reopened native bytes.
  for (const [rule, spacing] of [['Fixed', 3600], ['SpaceOnly', 1500], ['Minimum', 4000]] as const) {
    const doc = new HwpDocument(await fs.readFile(samples['ordinary-short-hwp'].file));
    try {
      const control = JSON.parse(doc.getControls()).find((c: any) => c.list === 0 && c.ctrlId === 'tbl');
      doc.applyParaFormatInCell(0, 0, control.controlIndex, 0, 0, JSON.stringify({ lineSpacingType: rule, lineSpacing: spacing }));
      for (const format of ['hwp', 'hwpx']) {
        const bytes = format === 'hwp' ? doc.exportHwp() : doc.exportHwpx(), file = path.join(directory, `${rule}.${format}`);
        await fs.writeFile(file, bytes);
        const reopened = new HwpDocument(bytes);
        try {
          const properties = JSON.parse(reopened.getCellParaPropertiesAtByPath(0, 0, JSON.stringify([{ controlIndex: control.controlIndex, cellIndex: 0, cellParaIndex: 0 }])));
          expect(properties.lineSpacingType).toBe(rule); expect(properties.lineSpacing).toBe(Math.round(spacing / 15) / 10);
        } finally { reopened.free(); }
        samples[`${rule}-${format}`] = { file, expected: rule === 'Fixed' ? 24 : rule === 'SpaceOnly' ? 42 : 32 };
      }
    } finally { doc.free(); }
  }
});
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

for (const rule of ['Fixed', 'SpaceOnly', 'Minimum']) for (const format of ['hwp', 'hwpx'])
test(`${format} ${rule}: native cell rule retains its meaning after font growth and two DOCX saves`, async ({ page }, info) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'rule.' + format));
      let doc = importEditableHwp(await session.editable(), 'rule.' + format), target: any;
      const find = (n: any) => { if (n.type === 'paragraph' && JSON.stringify(n.content).includes('Hgyp TEST')) target = n; n.content?.forEach(find); }; find(doc.content);
      const nativePolicy = target.attrs.hwpCellLine;
      target.content[0].text = 'Hgyp TEST '.repeat(8); target.content[0].marks.find((m: any) => m.type === 'textStyle').attrs.fontSize = '24pt';
      const resize = (n: any) => { if (n.type === 'tableCell') n.attrs.colwidth = [120]; n.content?.forEach(resize); }; resize(doc.content);
      const edited = structuredClone(doc), states = [];
      for (let save = 0; save < 3; save++) {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const p = [...layout.element.querySelectorAll<HTMLElement>('td p')].find(p => p.textContent?.includes('Hgyp TEST'))!, tops: number[] = [];
          const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
          while (walker.nextNode()) { const r = document.createRange(); r.selectNodeContents(walker.currentNode);
            for (const box of r.getClientRects()) if (box.width > .1 && box.height > 1 && !tops.some(y => Math.abs(y - box.top) < .5)) tops.push(box.top); }
          tops.sort((a, b) => a - b);
          states.push({ pages: layout.pages, text: p.textContent, policy: p.dataset.hwpCellLine, height: p.getBoundingClientRect().height,
            advances: tops.slice(1).map((top, i) => top - tops[i]) });
        } finally { layout.dispose(); }
        if (save < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
      }
      return { nativePolicy, edited, states };
    } finally { session.close(); }
  }, { bytes: [...await fs.readFile(samples[`${rule}-${format}`].file)], format });
  await fs.writeFile(info.outputPath('cell-rule.json'), JSON.stringify(result, null, 2));
  expect(result.nativePolicy).toContain(rule);
  for (const state of result.states) {
    expect(state.pages).toBe(1); expect(state.policy).toContain(rule); expect(state.text).toBe('Hgyp TEST '.repeat(8));
    expect(state.advances.length).toBeGreaterThan(2);
    for (const pitch of state.advances) expect(Math.abs(pitch - samples[`${rule}-${format}`].expected)).toBeLessThan(.6);
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'rule.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.edited)) });
  const paragraph = page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).locator('td p').filter({ hasText: 'Hgyp TEST' });
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => paragraph.evaluate((p, zoom) => p.getBoundingClientRect().height / (zoom / 100), zoom)).toBeCloseTo(result.states[0].height, 0);
  }
});

for (const nested of [false, true]) for (const format of ['hwp', 'hwpx'])
test(`${format} ${nested ? 'nested' : 'ordinary'}: large whitespace keeps its line budget without enlarging unmarked text`, async ({ page }, info) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'space.' + format));
      let doc = importEditableHwp(await session.editable(), 'space.' + format), target: any;
      const find = (n: any) => { if (n.type === 'paragraph' && JSON.stringify(n.content).includes('Hgyp TEST')) target = n; n.content?.forEach(find); }; find(doc.content);
      const marks = structuredClone(target.content[0].marks); marks.find((m: any) => m.type === 'textStyle').attrs.fontSize = '24pt';
      target.content = [{ type: 'text', text: ' ', marks }, { type: 'text', text: 'Hgyp TEST' }];
      const original = structuredClone(doc), states = [];
      for (let save = 0; save < 3; save++) {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const p = [...layout.element.querySelectorAll<HTMLElement>('td p')].find(p => p.textContent === ' Hgyp TEST')!, walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT), sizes: number[] = [], tops: number[] = [];
          while (walker.nextNode()) if (walker.currentNode.textContent?.trim()) { sizes.push(parseFloat(getComputedStyle(walker.currentNode.parentElement!).fontSize));
            const range = document.createRange(); range.selectNodeContents(walker.currentNode); tops.push(range.getBoundingClientRect().top - p.getBoundingClientRect().top); }
          states.push({ height: p.getBoundingClientRect().height, sizes, tops, policy: p.dataset.hwpCellLine });
        } finally { layout.dispose(); }
        if (save < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
      }
      return { original, states };
    } finally { session.close(); }
  }, { bytes: [...await fs.readFile(samples[`${nested ? 'nested' : 'ordinary'}-short-${format}`].file)], format });
  await fs.writeFile(info.outputPath('mixed-strut.json'), JSON.stringify(result.states, null, 2));
  for (const state of result.states) {
    expect(state.policy).toBeTruthy(); expect(Math.abs(state.height - 51.2)).toBeLessThan(.3);
    for (const font of state.sizes) expect(Math.abs(font - 40 / 3)).toBeLessThan(.1);
    expect(state.tops).toEqual(result.states[0].tops);
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'space.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.original)) });
  const p = page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).locator('td p').filter({ hasText: 'Hgyp TEST' }).last();
  await expect.poll(() => p.evaluate(p => p.getBoundingClientRect().height)).toBeCloseTo(51.2, 0);
  expect(await p.evaluate(p => parseFloat(getComputedStyle(p).fontSize))).toBeCloseTo(40 / 3, 1);
});

import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import JSZip from 'jszip';
const createFixture = createRequire(import.meta.url)('./hwp-cell-reflow-fixture.cjs');
let directory: string, samples: Record<string, any>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-cell-reflow-')); samples = await createFixture(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });
for (const nested of [false, true]) for (const format of ['hwp', 'hwpx']) for (const action of ['text', 'font', 'width'])
  test(`${format}: ${nested ? 'nested' : 'ordinary'} final cell retains native line advance after ${action} reflow and DOCX/Hancom saves`, async ({ page }, info) => {
    const sample = samples[`${nested ? 'nested' : 'ordinary'}-${action === 'width' ? 'wide' : 'short'}-${format}`];
    const expectedAdvance = sample.advance * (action === 'font' ? 32 / sample.fontPx : 1);
    await page.goto('/');
    const result = await page.evaluate(async ({ bytes, format, action }) => {
      const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp, exportEditableHwp } = await import('/src/hwp/editing.ts');
      const { exportDocx, importFile } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
      const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
      const session = new HwpReaderSession();
      try {
        await session.open(new File([new Uint8Array(bytes)], 'reflow.' + format));
        let doc = importEditableHwp(await session.editable(), 'reflow.' + format);
        const visit = (n: any) => {
          if (n.type === 'tableCell' && n.content?.some((c: any) => c.type === 'paragraph' && JSON.stringify(c).includes('Hgyp TEST'))) {
            n.attrs.colwidth = [action === 'width' ? 120 : 240];
            const p = n.content.find((c: any) => c.type === 'paragraph' && JSON.stringify(c).includes('Hgyp TEST'));
            const mark = structuredClone(p.content[0].marks);
            if (action !== 'width') {
              if (action === 'font') mark.find((m: any) => m.type === 'textStyle').attrs.fontSize = '24pt';
              p.content = [{ type: 'text', text: 'Hgyp TEST '.repeat(14), marks: mark }];
            }
          }
          n.content?.forEach(visit);
        };
        visit(doc.content); const edited = structuredClone(doc), states = [];
        for (let save = 0; save < 5; save++) {
          const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
          try {
            const p = [...layout.element.querySelectorAll<HTMLParagraphElement>('.print-content td p')].find(p => p.textContent!.includes('Hgyp TEST'))!;
            const tops: number[] = [], walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
            while (walker.nextNode()) { const range = document.createRange(); range.selectNodeContents(walker.currentNode);
              for (const r of range.getClientRects()) if (r.width > .1 && r.height > 1 && !tops.some(y => Math.abs(y - r.top) < .5)) tops.push(r.top); }
            tops.sort((a, b) => a - b);
            states.push({ pages: layout.pages, text: p.textContent, height: p.getBoundingClientRect().height,
              advances: tops.slice(1).map((y, i) => y - tops[i]), lineHeight: getComputedStyle(p).lineHeight });
          } finally { layout.dispose(); }
          if (save < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
          else if (save < 4) {
            const output = await exportEditableHwp(doc, format as 'hwp' | 'hwpx'), next = new HwpReaderSession();
            try { await next.open(new File([output.blob], 'again.' + format)); doc = importEditableHwp(await next.editable(), 'again.' + format); }
            finally { next.close(); }
          }
        }
        return { edited, states };
      } finally { session.close(); }
    }, { bytes: [...await fs.readFile(sample.file)], format, action });
    await fs.writeFile(info.outputPath('native-and-reflow.json'), JSON.stringify({ native: sample.metrics, expectedAdvance, states: result.states }, null, 2));
    for (const state of result.states) {
      expect(state.pages).toBe(1); expect(state.advances.length).toBeGreaterThan(1);
      expect(state.text).toBe(result.states[0].text);
      for (const advance of state.advances) expect(Math.abs(advance - expectedAdvance)).toBeLessThan(.6);
    }
    await page.locator('input[type=file]').first().setInputFiles({ name: 'reflow.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.edited)) });
    const p = page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).locator('p').filter({ hasText: 'Hgyp TEST' }).last();
    for (const zoom of [50, 100, 125, 200, 100]) {
      const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
      await expect.poll(async () => p.evaluate((p, { zoom, expectedAdvance }) => {
        const tops: number[] = [], walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
        while (walker.nextNode()) { const range = document.createRange(); range.selectNodeContents(walker.currentNode);
          for (const r of range.getClientRects()) if (r.width > .1 && r.height > 1 && !tops.some(y => Math.abs(y - r.top) < .5)) tops.push(r.top); }
        tops.sort((a, b) => a - b);
        return Math.max(...tops.slice(1).map((y, i) => Math.abs((y - tops[i]) / (zoom / 100) - expectedAdvance)));
      }, { zoom, expectedAdvance })).toBeLessThan(.6);
    }
    const index = await p.evaluate(p => [...p.closest('.tiptap')!.querySelectorAll('p')].indexOf(p));
    const target = page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).locator('p').nth(index);
    const original = await target.textContent(); await target.fill('CHANGED'); await page.keyboard.press('ControlOrMeta+z'); await expect(target).toHaveText(original!);
    await page.getByRole('button', { name: '다시 실행 (Ctrl+Y)' }).click(); await expect(target).toHaveText('CHANGED');
    await page.getByRole('button', { name: '실행 취소 (Ctrl+Z)' }).click(); await expect(target).toHaveText(original!);
  });

for (const change of ['line', 'font', 'malformed']) test(`DOCX ${change} changes override optional HWP line metadata`, async ({ page }) => {
  await page.goto('/');
  const bytes = await page.evaluate(async bytes => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx } = await import('/src/document-io.ts'), session = new HwpReaderSession();
    try { await session.open(new File([new Uint8Array(bytes)], 'source.hwp'));
      return [...new Uint8Array(await (await exportDocx(importEditableHwp(await session.editable(), 'source.hwp'))).arrayBuffer())];
    } finally { session.close(); }
  }, [...await fs.readFile(samples['ordinary-short-hwp'].file)]);
  const zip = await JSZip.loadAsync(new Uint8Array(bytes));
  let xml = await zip.file('word/document.xml')!.async('string');
  expect(xml).toContain('wsh:cellLine=');
  if (change === 'line') xml = xml.replace(/w:line="200"/g, 'w:line="450"');
  else if (change === 'font') xml = xml.replace(/w:sz w:val="20"/g, 'w:sz w:val="48"');
  else xml = xml.replace(/wsh:cellLine="[^"]+"/g, 'wsh:cellLine="invalid"');
  zip.file('word/document.xml', xml);
  const result = await page.evaluate(async bytes => {
    const { importFile } = await import('/src/document-io.ts'), { parseDocumentHtml } = await import('/src/extensions.ts');
    const loaded = await importFile(new File([new Uint8Array(bytes)], 'external.docx'));
    const doc = parseDocumentHtml(loaded.html!); let target: any;
    const walk = (n: any) => { if (n.type === 'paragraph' && JSON.stringify(n.content).includes('Hgyp TEST')) target = n; n.content?.forEach(walk); }; walk(doc);
    return target;
  }, [...await zip.generateAsync({ type: 'uint8array' })]);
  expect(result.attrs.hwpCellLine).toBeNull();
  expect(result.attrs.lineSpacing).toBe(change === 'line' ? '22.5pt' : '10pt');
  if (change === 'font') expect(result.content[0].marks.find((m: any) => m.type === 'textStyle').attrs.fontSize).toBe('24pt');
});

for (const nested of [false, true]) test(`${nested ? 'nested' : 'ordinary'} fractional point font keeps the terminal cell line after DOCX quantization`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async bytes => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'fractional.hwp'));
      let doc = importEditableHwp(await session.editable(), 'fractional.hwp');
      const walk = (n: any) => { if (n.type === 'text' && n.text === 'Hgyp TEST') n.marks.find((m: any) => m.type === 'textStyle').attrs.fontSize = '9.6pt'; n.content?.forEach(walk); }; walk(doc.content);
      const states = [];
      for (let save = 0; save < 3; save++) {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const p = [...layout.element.querySelectorAll<HTMLParagraphElement>('.print-content td p')].find(p => p.textContent === 'Hgyp TEST')!;
          states.push({ pages: layout.pages, height: p.getBoundingClientRect().height, policy: p.dataset.hwpCellLine });
        } finally { layout.dispose(); }
        if (save < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
      }
      return states;
    } finally { session.close(); }
  }, [...await fs.readFile(samples[`${nested ? 'nested' : 'ordinary'}-short-hwp`].file)]);
  for (const state of result) { expect(state.pages).toBe(1); expect(state.policy).toBeTruthy(); expect(Math.abs(state.height - 12.8)).toBeLessThan(.25); }
});

for (const nested of [false, true]) for (const format of ['hwp', 'hwpx']) test(`${format}: ${nested ? 'nested' : 'ordinary'} mixed size runs on one baseline retain one compact terminal line`, async ({ page }) => {
  const sample = samples[`${nested ? 'nested' : 'ordinary'}-mixed-${format}`];
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'mixed.' + format));
      let doc = importEditableHwp(await session.editable(), 'mixed.' + format); const states = [];
      for (let save = 0; save < 3; save++) {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const p = [...layout.element.querySelectorAll<HTMLParagraphElement>('.print-content td p')].find(p => p.textContent === 'BIG Hgyp TEST')!;
          states.push({ pages: layout.pages, height: p.getBoundingClientRect().height });
        } finally { layout.dispose(); }
        if (save < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
      }
      return states;
    } finally { session.close(); }
  }, { bytes: [...await fs.readFile(sample.file)], format });
  for (const state of result) { expect(state.pages).toBe(1); expect(Math.abs(state.height - sample.inkHeight)).toBeLessThan(.3); }
});

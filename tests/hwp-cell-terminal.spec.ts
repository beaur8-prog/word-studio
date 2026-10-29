import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-cell-reflow-fixture.cjs');
let directory: string, samples: Record<string, any>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-cell-terminal-')); samples = await fixture(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

for (const format of ['hwp', 'hwpx']) for (const kind of ['wrapped', 'breaks']) for (const mode of ['ordinary', 'nested', 'headers', 'merge-top', 'merge-middle', 'merge-bottom', 'merge-header-top', 'merge-header-middle', 'merge-header-bottom', 'merge-border-top', 'merge-header-border-middle'])
test(`${format} ${mode} ${kind}: terminal cell budget survives page fragments, editing and two DOCX saves`, async ({ page }, info) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format, kind, mode }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'terminal.' + format));
      let doc = importEditableHwp(await session.editable(), 'terminal.' + format);
      doc.page = { ...doc.page, size: 'Custom', width: 150, height: 180, margin: 10, margins: undefined, pageNumbers: false };
      let cell: any, table: any, paragraph: any;
      const find = (n: any, parent: any) => {
        if (n.type === 'tableCell' && n.content?.some((p: any) => p.type === 'paragraph' && JSON.stringify(p.content).includes('Hgyp TEST'))) { cell = n; table = parent.parent; paragraph = n.content.find((p: any) => p.type === 'paragraph' && JSON.stringify(p.content).includes('Hgyp TEST')); }
        n.content?.forEach((c: any) => find(c, { node: n, parent: parent?.node }));
      }; find(doc.content, null);
      const marks = structuredClone(paragraph.content[0].marks);
      const p = (text: string, height = '20px') => ({ type: 'paragraph', attrs: { ...paragraph.attrs, hwpCellLine: null, paragraphSize: '10pt', lineSpacing: height, spaceBefore: '0px', spaceAfter: '0px' }, content: [{ type: 'text', text, marks: structuredClone(marks) }] });
      const words = Array.from({ length: kind === 'breaks' ? 110 : 120 }, (_, i) => `ITEM${String(i).padStart(3, '0')}`);
      paragraph.content = kind === 'wrapped' ? [{ type: 'text', text: words.join(' '), marks }] : words.flatMap((text, i) => [...(i ? [{ type: 'hardBreak' }] : []), { type: 'text', text, marks: structuredClone(marks) }]);
      cell.attrs = { ...cell.attrs, colwidth: [160], paddingTop: '4px', paddingBottom: '4px', paddingLeft: '4px', paddingRight: '4px', verticalAlign: mode.startsWith('merge-') ? mode.split('-').at(-1) : 'top' };
      if (mode.includes('border')) Object.assign(cell.attrs, { borderTop: '4px solid #000000', borderRight: '4px solid #000000', borderBottom: '4px solid #000000', borderLeft: '4px solid #000000', paddingTop: '11px', paddingBottom: '7px' });
      const relax = (n: any) => { if (n.type === 'tableRow') n.attrs = { ...n.attrs, rowSplit: 'allow' }; if (n.type === 'tableCell' && n !== cell) n.attrs = { ...n.attrs, colwidth: [420] }; n.content?.forEach(relax); }; relax(doc.content);
      if (mode.startsWith('merge-')) {
        cell.attrs.rowspan = 6; cell.attrs.cellSplit = 'allow';
        table.content = Array.from({ length: 6 }, (_, i) => ({ type: 'tableRow', attrs: { ...table.content[0].attrs, rowHeight: 140, rowSplit: 'allow' }, content: [
          ...(i ? [] : [cell]), { type: 'tableCell', attrs: { ...cell.attrs, rowspan: 1, cellSplit: null, colwidth: [240], verticalAlign: 'top' }, content: [p('RIGHT' + i)] },
        ] }));
      }
      if (mode === 'headers' || mode.startsWith('merge-header-')) table.content.unshift({ type: 'tableRow', attrs: { ...table.content[0].attrs, rowHeight: 40, rowSplit: 'allow' }, content:
        (mode === 'headers' ? [160] : [160, 240]).map((width, i) => ({ type: 'tableHeader', attrs: { ...cell.attrs, rowspan: 1, cellSplit: null, colwidth: [width] }, content: [p('HEADER' + i)] })) });
      doc.content.content!.unshift(p('LEAD', '140px'));
      const original = structuredClone(doc), states = [];
      for (let save = 0; save < 3; save++) {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const ink: { text: string; tops: number[] }[] = Array.from({ length: layout.pages }, () => ({ text: '', tops: [] }));
          for (const p of layout.element.querySelectorAll<HTMLElement>('.print-content td p')) if (p.textContent!.includes('ITEM')) {
            const paper = p.closest<HTMLElement>('.print-page')!, page = Number(paper.dataset.page) - 1, origin = paper.getBoundingClientRect().top;
            const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
            while (walker.nextNode()) { const node = walker.currentNode as Text;
              for (let i = 0; i < node.length; i++) if (!/\s/.test(node.data[i])) { const range = document.createRange(); range.setStart(node, i); range.setEnd(node, i + 1); ink[page].text += node.data[i]; ink[page].tops.push(range.getBoundingClientRect().top - origin); }
            }
          }
          states.push({ pages: layout.pages, ink, secondary: [...layout.element.querySelectorAll<HTMLElement>('.print-content td p')].filter(p => p.textContent?.startsWith('RIGHT')).map(p => {
            const range = document.createRange(); range.selectNodeContents(p); const paper = p.closest<HTMLElement>('.print-page')!;
            return { text: p.textContent, page: Number(paper.dataset.page) - 1, top: range.getBoundingClientRect().top - paper.getBoundingClientRect().top };
          }), borders: [...layout.element.querySelectorAll<HTMLElement>('.print-content td p')].filter(p => p.textContent?.includes('ITEM')).map(p => parseFloat(getComputedStyle(p.closest('td')!).borderTopWidth)),
            hardBreaks: [...layout.element.querySelectorAll<HTMLElement>('.print-content td p')].filter(p => p.textContent?.includes('ITEM')).reduce((sum, p) => sum + p.querySelectorAll('br').length, 0), scaled: [...layout.element.querySelectorAll<HTMLElement>('[style]')].some(el => !!el.style.zoom),
            terminalFragments: [...layout.element.querySelectorAll<HTMLElement>('[data-hwp-cell-line]')].filter(p => p.textContent?.includes('ITEM')).map(p => ({ continuation: !!p.dataset.hwpCellContinuation, terminal: !!p.dataset.hwpCellTerminal, height: p.getBoundingClientRect().height,
              top: p.getBoundingClientRect().top - p.closest('.print-page')!.getBoundingClientRect().top, cellTop: p.closest('td')!.getBoundingClientRect().top - p.closest('.print-page')!.getBoundingClientRect().top,
              style: p.getAttribute('style'), prefixes: [...p.querySelectorAll<HTMLElement>('[data-hwp-cell-prefix]')].map(s => s.getAttribute('style')) })) });
        } finally { layout.dispose(); }
        if (save < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
      }
      return { original, states, expected: words.join('') };
    } finally { session.close(); }
  }, { bytes: [...await fs.readFile(samples[`${mode === 'nested' ? 'nested' : 'ordinary'}-short-${format}`].file)], format, kind, mode });
  await fs.writeFile(info.outputPath('terminal-model.json'), JSON.stringify(result.original));
  await fs.writeFile(info.outputPath('terminal-pages.json'), JSON.stringify(result.states, null, 2));
  const expected = result.states[0];
  for (const state of result.states) {
    expect(state.scaled).toBe(false); expect(state.pages).toBeGreaterThan(1);
    expect(state.hardBreaks).toBe(kind === 'breaks' ? 109 : 0);
    if (mode.includes('border')) expect(state.borders.every(b => b === 4)).toBe(true);
    expect(state.ink.map(p => p.text).join('')).toBe(result.expected);
    expect(state.ink.map(p => p.text)).toEqual(expected.ink.map(p => p.text));
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'terminal.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.original)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    try { await expect.poll(() => editor.evaluate((root, { zoom, expected }) => {
      const sheet = root.closest<HTMLElement>('.document-stack')!, height = Number(sheet.dataset.pageHeight), stride = height + 24, origin = root.getBoundingClientRect().top, scale = zoom / 100;
      const text: string[] = expected.ink.map(() => ''), counts: number[] = []; let error = 0, errorAt = '';
      for (const p of root.querySelectorAll('td p')) if (p.textContent!.includes('ITEM')) {
        const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
        while (walker.nextNode()) { const node = walker.currentNode as Text; if (node.parentElement?.closest('.ProseMirror-widget')) continue;
          for (let i = 0; i < node.length; i++) if (!/\s/.test(node.data[i])) { const range = document.createRange(); range.setStart(node, i); range.setEnd(node, i + 1); const y = (range.getBoundingClientRect().top - origin) / scale, page = Math.max(0, Math.floor((y + .5) / stride));
            const local = y - page * stride, index = counts[page] || 0;
            const delta = Math.abs(local - (expected.ink[page]?.tops[index] ?? Infinity));
            if (delta > error) { error = delta; errorAt = `${zoom}% p${page + 1}: ${local} vs ${expected.ink[page]?.tops[index]} (${node.data.slice(0, 35)}), p=${p.getAttribute('style')}`; }
            counts[page] = index + 1; text[page] = (text[page] || '') + node.data[i];
          }
        }
      }
      let sideError = 0, seen = 0, sideAt = '';
      for (const p of root.querySelectorAll('td p')) if (p.textContent?.startsWith('RIGHT')) {
        const range = document.createRange(); range.selectNodeContents(p); const y = (range.getBoundingClientRect().top - origin) / scale, page = Math.max(0, Math.floor((y + .5) / stride));
        const reference = expected.secondary.find(r => r.text === p.textContent), delta = reference?.page === page ? Math.abs(y - page * stride - reference.top) : Infinity;
        if (delta > sideError) { sideError = delta; sideAt = `${p.textContent}: page ${page + 1}, y=${y - page * stride}, expected=${JSON.stringify(reference)}`; } seen++;
      }
      return { text, aligned: error < 1.5 || errorAt, sideAligned: seen === expected.secondary.length && sideError < 1.5 || sideAt };
    }, { zoom, expected })).toEqual({ text: expected.ink.map(p => p.text), aligned: true, sideAligned: true });
    } catch (error) {
      await fs.writeFile(info.outputPath('editor-fragments.json'), JSON.stringify(await editor.evaluate(async root => {
        const origin = root.getBoundingClientRect().top;
        const { measureParagraphLines } = await import('/src/editor-lines.ts');
        root.classList.add('pagination-measuring-cells');
        const natural = [...root.querySelectorAll<HTMLElement>('td p')].filter(p => p.textContent?.includes('ITEM')).map(p => ({ height: p.getBoundingClientRect().height, data: { ...p.dataset }, cell: (() => { const c = p.closest('td')!, s = getComputedStyle(c); return { style: c.getAttribute('style'), paddingTop: s.paddingTop, paddingBottom: s.paddingBottom, borderTop: s.borderTopWidth, borderBottom: s.borderBottomWidth }; })(),
          lines: measureParagraphLines({ posAtDOM: () => 0 } as any, p, v => v - origin).map(({ top, bottom, before, after }: any) => ({ top, bottom, before, after })) }));
        root.classList.remove('pagination-measuring-cells');
        return { natural, stack: root.closest('.document-stack')?.outerHTML.slice(0, 1000), paragraphs: [...root.querySelectorAll<HTMLElement>('td p')].filter(p => p.textContent?.includes('ITEM')).map(p => ({
          style: p.getAttribute('style'), top: p.getBoundingClientRect().top - origin, height: p.getBoundingClientRect().height,
          prefixes: [...p.querySelectorAll<HTMLElement>('[data-hwp-cell-prefix]')].map(s => ({ style: s.getAttribute('style'), first: s.textContent?.slice(0, 35), last: s.textContent?.slice(-35) })),
          widgets: [...p.querySelectorAll<HTMLElement>('.ProseMirror-widget')].map(s => ({ html: s.outerHTML, top: s.getBoundingClientRect().top - origin })),
          lines: (() => { const values: { text: string; top: number }[] = [], walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT); while (walker.nextNode()) { const n = walker.currentNode as Text; for (let i = 0; i < n.length; i++) if (!/\s/.test(n.data[i])) { const r = document.createRange(); r.setStart(n, i); r.setEnd(n, i + 1); const top = r.getBoundingClientRect().top - origin, last = values.at(-1); if (last && Math.abs(last.top - top) < .5) last.text += n.data[i]; else values.push({ text: n.data[i], top }); } } return values; })(),
        })) };
      }), null, 2));
      throw error;
    }
  }
  const paragraphs = editor.locator('p'), index = await paragraphs.evaluateAll(ps => ps.findIndex(p => p.textContent!.includes('ITEM'))), target = paragraphs.nth(index);
  await target.fill('EDITED'); await page.keyboard.press('ControlOrMeta+z'); await expect(target).toContainText('ITEM000');
  const download = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
  const saved = JSON.parse(await fs.readFile((await (await download).path())!, 'utf8')); expect(saved.content).toEqual(result.original.content);
});

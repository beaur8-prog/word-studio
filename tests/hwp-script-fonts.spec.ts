import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-script-fonts-fixture.cjs');
let directory: string, files: Record<string, string>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-script-fonts-')); files = await fixture(directory, { latinSize: 125, hangulSpacing: -5, hangulWidth: 95, latinWidth: 110, columns: 2 }); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

for (const format of ['hwp', 'hwpx']) test(`${format}: source script fonts, sizes and widths reach standard DOCX and survive two saves`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const session = new HwpReaderSession(); let doc;
    try { await session.open(new File([new Uint8Array(bytes)], 'scripts.' + format)); doc = importEditableHwp(await session.editable(), 'scripts.' + format); }
    finally { session.close(); }
    const original = doc, states = [];
    const logical = (n: any): string => n.type === 'hardBreak' ? '\n' : (n.text || '') + (n.content || []).map(logical).join('');
    for (let save = 0; save < 3; save++) {
      const blob = await exportDocx(doc), zip = await JSZip.loadAsync(await blob.arrayBuffer());
      const xml = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
      const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      const runs = [...xml.getElementsByTagNameNS(W, 'r')].map(r => ({ text: [...r.getElementsByTagNameNS(W, 't')].map(t => t.textContent).join(''),
        font: r.getElementsByTagNameNS(W, 'rFonts')[0]?.getAttributeNS(W, 'ascii'),
        size: r.getElementsByTagNameNS(W, 'sz')[0]?.getAttributeNS(W, 'val'),
        tracking: r.getElementsByTagNameNS(W, 'spacing')[0]?.getAttributeNS(W, 'val'),
        width: r.getElementsByTagNameNS(W, 'w')[0]?.getAttributeNS(W, 'val') }));
      states.push({ body: logical(doc.content), runs });
      if (save < 2) { const opened = await importFile(new File([blob], 'again.docx')); doc = { ...opened.document, content: parseDocumentHtml(opened.html!) }; }
    }
    return { original, states };
  }, { bytes: [...await fs.readFile(files[format])], format });
  for (const state of result.states) {
    expect(state.body).toBe(result.states[0].body);
    expect(state.body).toContain('한글 Latin 123, Bold words\n다음 Next line.');
    const latin = state.runs.filter(r => /[A-Za-z]/.test(r.text));
    expect(latin.length).toBeGreaterThan(2);
    // The second source paragraph uses a 10pt run. The same 125% shape
    // attribute is relative to each run's base, including that paragraph.
    for (const run of latin) expect(run, run.text).toMatchObject({ font: 'Courier New', size: run.text === 'AFTER' ? '25' : '30', tracking: '0', width: '110' });
    for (const run of state.runs.filter(r => /[가-힣]/.test(r.text))) expect(run, run.text).toMatchObject({ font: 'Arial', size: '24', tracking: '-12', width: '95' });
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'scripts.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.original)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  await expect(editor).toContainText('Next line.');
  await expect.poll(() => editor.evaluate(root => {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) if (walker.currentNode.textContent?.includes('Latin 123,')) {
      const css = getComputedStyle(walker.currentNode.parentElement!);
      return [css.fontFamily, css.fontSize, css.letterSpacing === 'normal' ? 0 : parseFloat(css.letterSpacing)];
    }
    return [];
  })).toEqual(['"Courier New"', '20px', 0]);
  await editor.locator('p').filter({ hasText: 'Latin 123,' }).click(); await page.keyboard.press('End'); await page.keyboard.type(' EDIT');
  await expect(editor).toContainText('EDIT'); await page.keyboard.press('ControlOrMeta+z'); await expect(editor).not.toContainText('EDIT');
  await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(editor).toContainText('EDIT');
});

test('the authored Latin face survives native save and yields to a font edit', async ({ page }, info) => {
  const directoryForAlias = path.join(directory, 'authored-face');
  await fs.mkdir(directoryForAlias);
  const [authoredLatin, resolvedLatin] = ['HCI Poppy', 'Palatino Linotype'];
  const source = await fixture(directoryForAlias, { latinFace: authoredLatin, latinRenderedFace: resolvedLatin,
    hangulWidth: 95, body: '한글 Latin 123, Bold words.' });
  await page.goto('/');
  const result = await page.evaluate(async ({ hwp, hwpx }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp, exportEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_core.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const loadSource = async (bytes: number[], format: 'hwp' | 'hwpx') => {
      const reader = new HwpReaderSession();
      try {
        const input = new File([Uint8Array.from(bytes)], 'native-script-fonts.' + format);
        await reader.open(input);
        return importEditableHwp(await reader.editable(), input.name);
      } finally { reader.close(); }
    };
    let doc = await loadSource(hwp, 'hwp');
    const fromHwpx = await loadSource(hwpx, 'hwpx');
    const authoredHwpx = generateHTML(fromHwpx.content, editorExtensions).includes('HCI Poppy');
    const paragraph = doc.content.content?.find(node => node.content?.some(child => child.text?.startsWith('한글')));
    if (!paragraph) throw Error('Missing authored line');
    paragraph.attrs = { ...paragraph.attrs, lineSpacing: '24px', hwpBodyLine: '1 150 16 24' };
    const docx = [];
    for (let i = 0; i < 2; i++) {
      const saved = await exportDocx(doc);
      docx.push([...new Uint8Array(await saved.arrayBuffer())]);
      const opened = await importFile(new File([saved], 'again.docx'));
      doc = { ...opened.document, content: parseDocumentHtml(opened.html!) };
    }
    const print = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page,
      undefined, undefined, { hwpNativeExport: true, paintHwpCellLines: false });
    let usedStrut;
    try {
      usedStrut = Boolean(print.element.querySelector('span[data-hwp-latin-font] span[data-hwp-body-run="true"]'));
    } finally { print.dispose(); }
    const outputs = [];
    for (const format of ['hwp', 'hwpx'] as const) {
      const out = await exportEditableHwp(doc, format);
      outputs.push({ format, edited: false, bytes: [...new Uint8Array(await out.blob.arrayBuffer())] });
    }
    const update = (node: import('@tiptap/core').JSONContent) => {
      if (node.text?.includes('한글')) {
        const mark = node.marks?.find(m => m.type === 'textStyle');
        if (mark?.attrs) mark.attrs.fontFamily = '맑은 고딕';
      }
      node.content?.forEach(update);
    };
    update(doc.content);
    const edited = await exportEditableHwp(doc, 'hwpx');
    outputs.push({ format: 'hwpx', edited: true, bytes: [...new Uint8Array(await edited.blob.arrayBuffer())] });
    return { outputs, docx, usedStrut, authoredHwpx };
  }, { hwp: [...await fs.readFile(source.hwp)], hwpx: [...await fs.readFile(source.hwpx)] });
  expect(result.usedStrut).toBe(true);
  expect(result.authoredHwpx).toBe(true);

  const require = createRequire(import.meta.url), core = path.dirname(require.resolve('@rhwp/core'));
  const JSZip = require('jszip');
  const { pathToFileURL } = await import('node:url');
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  for (const [index, bytes] of result.docx.entries())
    await fs.writeFile(info.outputPath(`textbox-save-latin-${index + 1}.docx`), Buffer.from(bytes));
  for (const output of result.outputs) {
    const doc = new HwpDocument(Uint8Array.from(output.bytes));
    try {
      const paragraph = Array.from({ length: doc.getParagraphCount(0) }, (_, i) => i)
        .find(i => doc.getTextRange(0, i, 0, 0xffffffff).startsWith('한글 Latin 123,'));
      expect(paragraph, output.format).toBeDefined();
      const { charShapeId, fontFamilies, ratios, fontSize } = JSON.parse(doc.getCharPropertiesAt(0, paragraph!, 0));
      expect(fontFamilies.slice(0, 2)).toEqual(output.edited ? ['맑은 고딕', '맑은 고딕'] : ['Arial', resolvedLatin]);
      const header = await (await JSZip.loadAsync(doc.exportHwpx())).file('Contents/header.xml').async('string');
      const shape = [...header.matchAll(/<hh:charPr id="(\d+)"[^>]*>[\s\S]*?<\/hh:charPr>/g)]
        .find(([, id]) => Number(id) === charShapeId)?.[0];
      const id = /<hh:fontRef\b[^>]*\blatin="(\d+)"/.exec(shape ?? '')?.[1];
      const fonts = /<hh:fontface lang="LATIN"[^>]*>([\s\S]*?)<\/hh:fontface>/.exec(header)?.[1];
      const stored = [...(fonts ?? '').matchAll(/<hh:font id="(\d+)" face="([^"]+)"/g)].find(font => font[1] === id)?.[2];
      expect(stored).toBe(output.edited ? '맑은 고딕' : authoredLatin);
      expect(ratios[0]).toBe(95);
      expect(fontSize).toBe(1200);
    } finally { doc.free(); }
  }
});

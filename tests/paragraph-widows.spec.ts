import { test, expect } from '@playwright/test';
import { Document, Packer, Paragraph } from 'docx';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-script-cells-fixture.cjs');
let directory: string, files: Record<string, string>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-widows-')); files = await fixture(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

for (const format of ['hwp', 'hwpx']) test(`${format}: disabled widow protection survives body, ordinary and nested cells through DOCX`, async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const session = new HwpReaderSession(); let doc;
    try { await session.open(new File([new Uint8Array(bytes)], 'widows.' + format)); doc = importEditableHwp(await session.editable(), 'widows.' + format); }
    finally { session.close(); }
    const states = [], W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    for (let save = 0; save < 3; save++) {
      const blob = await exportDocx(doc), zip = await JSZip.loadAsync(await blob.arrayBuffer());
      const xml = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
      states.push([...xml.getElementsByTagNameNS(W, 'p')].map(p => ({
        text: [...p.getElementsByTagNameNS(W, 't')].map(t => t.textContent).join(''),
        widow: p.getElementsByTagNameNS(W, 'widowControl')[0]?.getAttributeNS(W, 'val') ?? 'true',
      })));
      if (save < 2) { const opened = await importFile(new File([blob], 'again.docx')); doc = { ...opened.document, content: parseDocumentHtml(opened.html!) }; }
    }
    return states;
  }, { bytes: [...await fs.readFile(files[format])], format });
  for (const state of states) for (const text of ['After astral body', 'Table Latin 456', 'Nested Latin 789', 'After nested table']) {
    const paragraph = state.find(p => p.text.includes(text));
    expect(paragraph, text).toBeTruthy();
    expect(paragraph!.widow, text).toMatch(/^(false|0|off)$/);
  }
});

test('external DOCX widow inheritance, direct overrides and defaults survive two saves', async ({ page }) => {
  const external = await Packer.toBuffer(new Document({
    styles: { paragraphStyles: [
      { id: 'Off', name: 'Off', paragraph: { widowControl: false } },
      { id: 'On', name: 'On', paragraph: { widowControl: true } },
    ] },
    sections: [{ children: [
      new Paragraph({ text: 'INHERITED-OFF', style: 'Off' }),
      new Paragraph({ text: 'DIRECT-ON', style: 'Off', widowControl: true }),
      new Paragraph({ text: 'DIRECT-OFF', style: 'On', widowControl: false }),
      new Paragraph({ text: 'DEFAULT-ON' }),
    ] }],
  }));
  await page.goto('/');
  const states = await page.evaluate(async bytes => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    let input = new File([new Uint8Array(bytes)], 'external.docx');
    const states = [], W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    for (let save = 0; save < 3; save++) {
      const opened = await importFile(input), doc = { ...opened.document, content: parseDocumentHtml(opened.html!) };
      const blob = await exportDocx(doc), zip = await JSZip.loadAsync(await blob.arrayBuffer());
      const xml = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
      states.push([...xml.getElementsByTagNameNS(W, 'p')].map(p => ({
        text: [...p.getElementsByTagNameNS(W, 't')].map(t => t.textContent).join(''),
        enabled: !['0', 'false', 'off'].includes(p.getElementsByTagNameNS(W, 'widowControl')[0]?.getAttributeNS(W, 'val') ?? 'true'),
      })));
      input = new File([blob], 'again.docx');
    }
    return states;
  }, [...external]);
  for (const state of states) expect(state).toEqual([
    { text: 'INHERITED-OFF', enabled: false }, { text: 'DIRECT-ON', enabled: true },
    { text: 'DIRECT-OFF', enabled: false }, { text: 'DEFAULT-ON', enabled: true },
  ]);
});

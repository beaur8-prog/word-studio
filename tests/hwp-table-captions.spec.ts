import { test, expect } from '@playwright/test';
import { createRequire } from 'node:module';
import fs from 'node:fs/promises';
const fixture = createRequire(import.meta.url)('./hwp-table-caption-fixture.cjs');

for (const format of ['hwp', 'hwpx']) test(`${format}: nested captions retain text, cells and picture bytes through two DOCX saves and editing`, async ({ page }, info) => {
  const { bytes } = await fixture({ format, side: 'TOP', nestedOwner: true, table: true, recursive: true, followingOwner: true, picture: true });
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts'), { createPrintLayout } = await import('/src/print-layout.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts'), { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const session = new HwpReaderSession(); let doc;
    try { await session.open(new File([new Uint8Array(bytes)], 'captions.' + format)); doc = importEditableHwp(await session.editable(), 'captions.' + format); }
    finally { session.close(); }
    const original = structuredClone(doc), states = [], files = [];
    const digest = async (src: string) => Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', await (await fetch(src)).arrayBuffer()))).map(n => n.toString(16).padStart(2, '0')).join('');
    const text = (node: Node): string => node.nodeName === 'BR' ? '\n' : node.nodeType === Node.TEXT_NODE ? node.textContent || '' : [...node.childNodes].map(text).join('');
    for (let save = 0; save < 3; save++) {
      const html = generateHTML(doc.content, editorExtensions), root = document.createElement('div'); root.innerHTML = html;
      const layout = await createPrintLayout(html, doc.page);
      states.push({ text: text(root), cells: [...root.querySelectorAll('td, th')].map(text), tables: root.querySelectorAll('table').length,
        images: await Promise.all([...root.querySelectorAll('img')].map(i => digest(i.src))), pages: layout.pages,
        printed: text(layout.element), annotations: html.includes('data-hwp-caption-key') }); layout.dispose();
      if (save < 2) { const blob = await exportDocx(doc); files.push([...new Uint8Array(await blob.arrayBuffer())]); const loaded = await importFile(new File([blob], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
    }
    return { states, original, files };
  }, { bytes: [...bytes], format });
  for (const [i, bytes] of result.files.entries()) await fs.writeFile(info.outputPath(`hwp-gov-caption-${format}-save${i+1}.docx`), Buffer.from(bytes));
  const first = result.states[0]; expect(first.tables).toBe(4); expect(first.images).toHaveLength(1);
  const captions = ['CAPTION-BEFORE', 'INNER-CAPTION', 'CAPTION-CELL-A', 'CAPTION-CELL-B', 'CAPTION-AFTER', 'SECOND-CAPTION'];
  for (const state of result.states) {
    expect(state.text).toBe(first.text); expect(state.cells).toEqual(first.cells); expect(state.images).toEqual(first.images);
    expect(state.tables).toBe(4); expect(state.pages).toBe(first.pages); expect(state.annotations).toBe(false);
    for (const caption of captions) { expect(state.text.split(caption)).toHaveLength(2); expect(state.printed.split(caption)).toHaveLength(2); }
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'captions.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.original)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' }), caption = editor.locator('p').filter({ hasText: /^CAPTION-CELL-A/ });
  await expect(caption).toBeVisible(); await expect(editor.locator('table')).toHaveCount(4);
  await caption.click(); await page.keyboard.press('End'); await page.keyboard.type(' EDIT'); await expect(caption).toContainText('EDIT');
  await page.keyboard.press('ControlOrMeta+z'); await expect(caption).not.toContainText('EDIT');
  await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(caption).toContainText('EDIT');
  await page.screenshot({ path: info.outputPath('caption-editor.png') });
});

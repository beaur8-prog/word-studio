import { test, expect, type Page } from '@playwright/test';
import { readFile, writeFile } from 'node:fs/promises';

// Authored controls only; the independent Korean university form stays local.
async function fixture(page: Page) {
  await page.goto('/');
  return page.evaluate(async () => {
    const { exportDocx } = await import('/src/document-io.ts'); const { defaultPage } = await import('/src/types.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const texts = ['MODERN OFF ', 'MODERN ON ', 'CUSTOM ', 'LEGACY DEFAULT ', 'LEGACY OVERRIDE ', 'TABLE '];
    const blob = await exportDocx({ format: 'word-studio', version: 1, title: 'CHECKBOXES', updatedAt: new Date().toISOString(), page: { ...defaultPage, header: 'HEADER ' }, comments: [],
      content: { type: 'doc', content: [...texts.slice(0, 5).map(text => ({ type: 'paragraph', content: [{ type: 'text', text }] })), { type: 'table', content: [{ type: 'tableRow', content: [{ type: 'tableCell', content: [{ type: 'paragraph', content: [{ type: 'text', text: texts[5] }] }] }] }] }] } });
    const zip = await JSZip.loadAsync(await blob.arrayBuffer());
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', X = 'http://schemas.microsoft.com/office/word/2010/wordml';
    const doc = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
    const paragraphs = [...doc.getElementsByTagNameNS(W, 'p')];
    const modern = (d: Document, on: boolean, custom = false) => {
      const sdt = d.createElementNS(W, 'w:sdt'), pr = d.createElementNS(W, 'w:sdtPr'), box = d.createElementNS(X, 'w14:checkbox');
      const state = d.createElementNS(X, 'w14:checked'); state.setAttributeNS(X, 'w14:val', on ? '1' : '0'); box.append(state);
      for (const [name, code] of [['checkedState', custom ? '2713' : '2612'], ['uncheckedState', '2610']]) {
        const el = d.createElementNS(X, 'w14:' + name); el.setAttributeNS(X, 'w14:val', code); el.setAttributeNS(X, 'w14:font', 'Arial'); box.append(el);
      }
      pr.append(box); sdt.append(pr); const content = d.createElementNS(W, 'w:sdtContent'), run = d.createElementNS(W, 'w:r'), text = d.createElementNS(W, 'w:t');
      // The first checked control deliberately has a stale cached character.
      text.textContent = custom ? '✓' : '☐'; run.append(text); content.append(run); sdt.append(content); return sdt;
    };
    paragraphs[0].append(modern(doc, false)); paragraphs[1].append(modern(doc, true)); paragraphs[2].append(modern(doc, true, true)); paragraphs[5].append(modern(doc, false));
    const legacy = (explicit: boolean | null) => {
      const r = doc.createElementNS(W, 'w:r'), begin = doc.createElementNS(W, 'w:fldChar'); begin.setAttributeNS(W, 'w:fldCharType', 'begin');
      const data = doc.createElementNS(W, 'w:ffData'), name = doc.createElementNS(W, 'w:name'); name.setAttributeNS(W, 'w:val', explicit === null ? 'DefaultBox' : 'OverrideBox'); data.append(name);
      const box = doc.createElementNS(W, 'w:checkBox'), size = doc.createElementNS(W, 'w:sizeAuto'), value = doc.createElementNS(W, 'w:default'); value.setAttributeNS(W, 'w:val', '1'); box.append(size, value);
      if (explicit !== null) { const checked = doc.createElementNS(W, 'w:checked'); checked.setAttributeNS(W, 'w:val', explicit ? '1' : '0'); box.append(checked); }
      data.append(box); begin.append(data); r.append(begin);
      const code = doc.createElementNS(W, 'w:r'), instr = doc.createElementNS(W, 'w:instrText'); instr.textContent = ' FORMCHECKBOX '; code.append(instr);
      const endRun = doc.createElementNS(W, 'w:r'), end = doc.createElementNS(W, 'w:fldChar'); end.setAttributeNS(W, 'w:fldCharType', 'end'); endRun.append(end);
      return [r, code, endRun];
    };
    paragraphs[3].append(...legacy(null)); paragraphs[4].append(...legacy(false));
    const headerName = Object.keys(zip.files).find(n => /^word\/header\d+\.xml$/.test(n))!;
    const header = new DOMParser().parseFromString(await zip.file(headerName)!.async('string'), 'application/xml'); header.getElementsByTagNameNS(W, 'p')[0].append(modern(header, true));
    zip.file(headerName, new XMLSerializer().serializeToString(header)); zip.file('word/document.xml', new XMLSerializer().serializeToString(doc));
    return [...await zip.generateAsync({ type: 'uint8array' })];
  });
}

test('modern and legacy checkbox state, custom glyph, table and header survive two DOCX saves', async ({ page }, info) => {
  const input = await fixture(page);
  const result = await page.evaluate(async input => {
    const { importFile, exportDocx } = await import('/src/document-io.ts'); const { parseDocumentHtml } = await import('/src/extensions.ts');
    let file = new File([new Uint8Array(input)], 'checkboxes.docx'); const states = [], saves = [];
    const text = (n: any): string => n.type === 'text' ? n.text : (n.content || []).map(text).join('');
    for (let i = 0; i < 3; i++) {
      const loaded = await importFile(file), doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
      states.push({ text: text(doc.content), header: doc.page.headerHtml, warnings: loaded.warnings });
      if (i < 2) { const b = await exportDocx(doc); saves.push([...new Uint8Array(await b.arrayBuffer())]); file = new File([b], 'again.docx'); }
    }
    return { states, saves };
  }, input);
  for (const [i, bytes] of [input, ...result.saves].entries()) await writeFile(info.outputPath(`textbox-save-checkbox-${i}.docx`), Buffer.from(bytes));
  for (const state of result.states) {
    expect(state.text).toBe('MODERN OFF ☐MODERN ON ☒CUSTOM ✓LEGACY DEFAULT ☒LEGACY OVERRIDE ☐TABLE ☐');
    expect(state.header).toContain('☒');
  }
  expect(result.states[0].warnings.join(' ')).toContain('양식 컨트롤');
});

test('checkbox symbols remain visible and editable through real Word exports and undo/redo', async ({ page }, info) => {
  const input = await fixture(page);
  await page.locator('input[type=file]').first().setInputFiles({ name: 'checkboxes.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(input) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true }); await expect(editor).toContainText('MODERN OFF ☐');
  await editor.locator('p').filter({ hasText: 'MODERN OFF' }).click(); await page.keyboard.press('End'); await page.keyboard.press('Shift+ArrowLeft'); await page.keyboard.insertText('☒');
  await expect(editor).toContainText('MODERN OFF ☒'); await page.keyboard.press('ControlOrMeta+z'); await expect(editor).toContainText('MODERN OFF ☐');
  await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(editor).toContainText('MODERN OFF ☒');
  for (let save = 1; save <= 2; save++) {
    await page.getByRole('button', { name: '파일', exact: true }).click(); const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
    const download = await pending, bytes = [...await readFile((await download.path())!)]; await writeFile(info.outputPath(`textbox-save-checkbox-ui-${save}.docx`), Buffer.from(bytes));
    const result = await page.evaluate(async bytes => { const { importFile } = await import('/src/document-io.ts'); const html = (await importFile(new File([new Uint8Array(bytes)], 'again.docx'))).html!; const template = document.createElement('template'); template.innerHTML = html; return { text: template.content.textContent, inputs: template.content.querySelectorAll('input').length }; }, bytes);
    expect(result.text).toContain('MODERN OFF ☒'); expect(result.text).toContain('CUSTOM ✓'); expect(result.inputs).toBe(0);
  }
});

import { test, expect, type Page } from '@playwright/test';
import { readFile, writeFile } from 'node:fs/promises';

// Synthetic OOXML preservation controls, separate from the Word-authored corpus.
async function source(page: Page, gradient: boolean) {
  await page.goto('/');
  return page.evaluate(async gradient => {
    const { exportDocx } = await import('/src/document-io.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const plain = new Uint8Array(await (await exportDocx({ format: 'word-studio', version: 1,
      title: 'BACKGROUND', updatedAt: new Date().toISOString(), page: { ...defaultPage }, comments: [],
      content: { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'BACKGROUND BODY' }] }] },
    })).arrayBuffer());
    const zip = await JSZip.loadAsync(plain), W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const doc = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
    const bg = doc.createElementNS(W, 'w:background');
    bg.setAttributeNS(W, 'w:color', gradient ? 'A1B2C3' : 'D0DCAF');
    bg.setAttributeNS(W, 'w:themeColor', gradient ? 'accent3' : 'background2');
    bg.setAttributeNS(W, 'w:themeTint', '33');
    if (gradient) {
      const fill = new DOMParser().parseFromString('<v:background xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" id="_x0000_s1025" o:bwmode="white" fillcolor="#a1b2c3 [662]" o:targetscreensize="1024,768"><v:fill color2="fill lighten(0)" focusposition="1" focussize="" method="linear sigma" focus="100%" type="gradientRadial"><o:fill v:ext="view" type="gradientCenter"/></v:fill></v:background>', 'application/xml');
      bg.append(doc.importNode(fill.documentElement, true));
    }
    doc.documentElement.insertBefore(bg, doc.documentElement.firstChild);
    const settings = new DOMParser().parseFromString(await zip.file('word/settings.xml')!.async('string'), 'application/xml');
    settings.getElementsByTagNameNS(W, 'displayBackgroundShape')[0].setAttributeNS(W, 'w:val', gradient ? 'true' : 'false');
    zip.file('word/document.xml', new XMLSerializer().serializeToString(doc));
    zip.file('word/settings.xml', new XMLSerializer().serializeToString(settings));
    return { input: [...await zip.generateAsync({ type: 'uint8array' })], plain: [...plain] };
  }, gradient);
}

async function inspect(page: Page, bytes: number[]) {
  return page.evaluate(async bytes => {
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const zip = await JSZip.loadAsync(new Uint8Array(bytes)), W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const doc = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
    const bg = doc.getElementsByTagNameNS(W, 'background')[0];
    const settings = new DOMParser().parseFromString(await zip.file('word/settings.xml')!.async('string'), 'application/xml');
    return { color: bg?.getAttributeNS(W, 'color') ?? null, theme: bg?.getAttributeNS(W, 'themeColor') ?? null,
      drawing: bg ? new XMLSerializer().serializeToString(bg) : null,
      display: settings.getElementsByTagNameNS(W, 'displayBackgroundShape')[0]?.getAttributeNS(W, 'val') ?? null };
  }, bytes);
}

for (const gradient of [false, true]) test(`DOCX ${gradient ? 'gradient' : 'solid'} background survives two saves without inheriting the export theme`, async ({ page }, info) => {
  const data = await source(page, gradient);
  const result = await page.evaluate(async data => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    let file = new File([new Uint8Array(data.input)], 'background.docx'); const states = [], files = [];
    for (let save = 0; save < 3; save++) {
      const loaded = await importFile(file), doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
      states.push({ background: (doc.page as any).wordBackground ?? null, content: doc.content });
      if (save < 2) { const blob = await exportDocx(doc); const bytes = [...new Uint8Array(await blob.arrayBuffer())]; files.push(bytes); file = new File([blob], 'background.docx'); }
    }
    return { states, files };
  }, data);
  for (const [i, bytes] of [data.input, ...result.files].entries()) await writeFile(info.outputPath(`textbox-save-background-${gradient}-${i}.docx`), Buffer.from(bytes));
  for (const state of result.states) {
    expect(state.background?.color).toBe(gradient ? 'A1B2C3' : 'D0DCAF');
    expect(state.background?.display).toBe(gradient);
    expect(state.content).toEqual(result.states[0].content);
    if (gradient) expect(state.background?.drawingXml).toContain('gradientRadial');
  }
  for (const bytes of result.files) {
    const actual = await inspect(page, bytes);
    expect(actual.color).toBe(gradient ? 'A1B2C3' : 'D0DCAF');
    expect(actual.theme).toBeNull();
    expect(actual.display).toBe(gradient ? 'true' : 'false');
    if (gradient) { expect(actual.drawing).toContain('gradientRadial'); expect(actual.drawing).toContain('gradientCenter'); expect(actual.drawing).not.toContain('[662]'); }
  }
});

test('editing and opening another document preserve then reset the imported background', async ({ page }, info) => {
  const data = await source(page, true);
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  await page.locator('input[type=file]').first().setInputFiles({ name: 'background.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(data.input) });
  await expect(editor).toHaveText('BACKGROUND BODY');
  await editor.click(); await page.keyboard.press('End'); await page.keyboard.insertText(' EDIT');
  for (const plain of [false, true]) {
    if (plain) {
      await page.locator('input[type=file]').first().setInputFiles({ name: 'plain.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(data.plain) });
      if (await page.getByRole('dialog', { name: '변경 내용 보관' }).count()) await page.getByRole('button', { name: '복구본 보관 후 계속' }).click();
      await expect(editor).toHaveText('BACKGROUND BODY');
    }
    await page.getByRole('button', { name: '파일', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
    const download = await pending, bytes = [...await readFile((await download.path())!)];
    await writeFile(info.outputPath(`textbox-save-background-editor-${plain}.docx`), Buffer.from(bytes));
    const actual = await inspect(page, bytes);
    expect(actual.color).toBe(plain ? null : 'A1B2C3');
    if (!plain) expect(actual.drawing).toContain('gradientRadial');
  }
});

test('background metadata rejects external relationships, active markup and unbounded payloads', async ({ page }) => {
  await page.goto('/');
  const rejected = await page.evaluate(async () => {
    const { validateDocument } = await import('/src/document-io.ts'); const { defaultPage } = await import('/src/types.ts');
    const base = { format: 'word-studio', version: 1, title: 'BOUND', updatedAt: new Date().toISOString(), comments: [], content: { type: 'doc', content: [{ type: 'paragraph' }] } };
    const cases = [
      { color: 'not-rgb', display: true },
      { color: 'A1B2C3', display: true, drawingXml: '<script>run()</script>' },
      { color: 'A1B2C3', display: true, drawingXml: '<v:background xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId7"/>' },
      { color: 'A1B2C3', display: true, drawingXml: '<v:background xmlns:v="urn:schemas-microsoft-com:vml" onclick="run()"/>' },
      { color: 'A1B2C3', display: true, drawingXml: 'x'.repeat(32769) },
    ];
    return cases.map(wordBackground => { try { validateDocument({ ...base, page: { ...defaultPage, wordBackground } }); return false; } catch { return true; } });
  });
  expect(rejected).toEqual([true, true, true, true, true]);
});

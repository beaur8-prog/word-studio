import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';

// Independent boundary fixture: the final glyphs fit, but its authored exact
// line box extends ten pixels below the body. No text or font is shortened.
for (const setting of ['on', 'false', 'zero', 'off', 'absent'] as const) {
  test(`Word bottom spacing ${setting}: page policy survives editing and two DOCX saves`, async ({ page }) => {
    await page.goto('/');
    const result = await page.evaluate(async setting => {
      const { defaultPage } = await import('/src/types.ts');
      const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
      const { exportDocx, importFile } = await import('/src/document-io.ts');
      const { createPrintLayout } = await import('/src/print-layout.ts');
      const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
      const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
      const paper = { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10, pageNumbers: false };
      const body = 160 * 96 / 25.4;
      const initial: any = { format: 'word-studio', version: 1, id: 'bottom-spacing', title: 'Bottom spacing',
        updatedAt: new Date().toISOString(), comments: [], page: paper,
        content: parseDocumentHtml(`<p style="font-size:1px;margin:0;line-height:${body - 70}px"></p><p style="font-size:40pt;line-height:80px;margin:0">FINAL</p>`) };
      const zip = await JSZip.loadAsync(await (await exportDocx(initial)).arrayBuffer());
      const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      const settings = new DOMParser().parseFromString(await zip.file('word/settings.xml')!.async('string'), 'application/xml');
      const compat = settings.getElementsByTagNameNS(W, 'compat')[0];
      const flag = settings.getElementsByTagNameNS(W, 'suppressBottomSpacing')[0] || settings.createElementNS(W, 'w:suppressBottomSpacing');
      if (!flag.parentNode) compat.prepend(flag);
      if (setting === 'absent') flag.remove();
      else if (setting !== 'on') flag.setAttributeNS(W, 'w:val', setting === 'zero' ? '0' : setting);
      zip.file('word/settings.xml', new XMLSerializer().serializeToString(settings));
      let file = new File([await zip.generateAsync({ type: 'uint8array' })], 'boundary.docx');
      const states: any[] = [], savedFlags: boolean[] = [], saves: number[][] = [];
      let doc: any;
      for (let save = 0; save < 3; save++) {
        const loaded = await importFile(file);
        doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const final = [...layout.element.querySelectorAll<HTMLElement>('p')].find(p => p.textContent === 'FINAL')!;
          states.push({ enabled: doc.page.suppressBottomSpacing, pages: layout.pages, text: final.textContent,
            height: final.getBoundingClientRect().height, fontSize: getComputedStyle(final).fontSize,
            outside: final.getBoundingClientRect().bottom - final.closest('.print-content')!.getBoundingClientRect().bottom });
        } finally { layout.dispose(); }
        if (save < 2) {
          const output = await exportDocx(doc), bytes = new Uint8Array(await output.arrayBuffer()), outputZip = await JSZip.loadAsync(bytes);
          saves.push([...bytes]);
          const raw = new DOMParser().parseFromString(await outputZip.file('word/settings.xml')!.async('string'), 'application/xml');
          const saved = raw.getElementsByTagNameNS(W, 'suppressBottomSpacing')[0];
          savedFlags.push(!!saved && !['false', '0', 'off'].includes(saved.getAttributeNS(W, 'val') || ''));
          file = new File([output], 'boundary.docx');
        }
      }
      return { states, savedFlags, doc, saves };
    }, setting);
    for (const [index, bytes] of result.saves.entries()) await writeFile(test.info().outputPath(`textbox-save-bottom-spacing-${index + 1}.docx`), Buffer.from(bytes));
    const enabled = setting === 'on';
    for (const state of result.states) {
      expect(state.enabled).toBe(enabled);
      expect(state.pages).toBe(enabled ? 1 : 2);
      expect(state.text).toBe('FINAL');
      expect(state.height).toBeCloseTo(80, 0);
      expect(parseFloat(state.fontSize)).toBeCloseTo(40 * 4 / 3, 1);
      if (enabled) expect(state.outside).toBeGreaterThan(9);
      else expect(state.outside).toBeLessThan(0);
    }
    expect(result.savedFlags).toEqual([enabled, enabled]);
    await page.locator('input[type=file]').first().setInputFiles({ name: 'boundary.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.doc)) });
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
    const last = editor.locator('p').filter({ hasText: /^FINAL$/ });
    await expect(last).toBeVisible();
    for (const zoom of [100, 50, 125, 100]) {
      const slider = page.getByRole('slider', { name: '확대 축소' });
      await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
      await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe(enabled ? '1' : '2');
    }
    const before = await editor.evaluate((el: any) => el.editor.getJSON());
    await last.fill('WRAP '.repeat(100));
    await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).not.toBe('1');
    await page.keyboard.press('ControlOrMeta+z');
    await expect(last).toHaveText('FINAL');
    expect(await editor.evaluate((el: any) => el.editor.getJSON())).toEqual(before);
    await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe(enabled ? '1' : '2');
    await page.keyboard.press('ControlOrMeta+Shift+z');
    await expect(editor.locator('p').nth(1)).toContainText('WRAP');
  });
}

test('backup validation rejects a non-boolean bottom-spacing policy', async ({ page }) => {
  await page.goto('/');
  const error = await page.evaluate(async () => {
    const { defaultPage } = await import('/src/types.ts');
    const { importFile } = await import('/src/document-io.ts');
    const doc = { format: 'word-studio', version: 1, title: 'Invalid', updatedAt: new Date().toISOString(), comments: [],
      page: { ...defaultPage, suppressBottomSpacing: 'false' }, content: { type: 'doc', content: [{ type: 'paragraph' }] } };
    try { await importFile(new File([JSON.stringify(doc)], 'invalid.wordstudio')); return null; }
    catch (error) { return String(error); }
  });
  expect(error).toContain('용지 옵션이 올바르지 않습니다.');
});

import { test, expect } from '@playwright/test';
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';

const paper = { size: 'A4', orientation: 'portrait', margin: 25.4, header: '', footer: '', pageNumbers: false };
for (const [side, x, y] of [['right', 900, 100], ['left', -60, 100], ['bottom', 80, 1400], ['top', 80, -60]] as const) {
  test(`a wholly off-paper ${side} drawing keeps its asset and authored offsets without appearing on paper`, async ({ page }, testInfo) => {
    await page.goto('/');
    const source = await page.evaluate(({ x, y, paper }) => {
      const canvas = document.createElement('canvas'); canvas.width = canvas.height = 4;
      const ctx = canvas.getContext('2d')!; ctx.fillStyle = '#ff0000'; ctx.fillRect(0, 0, 4, 4);
      const image = document.createElement('img'); image.src = canvas.toDataURL(); image.width = image.height = 40;
      image.dataset.inlineImage = 'true'; image.alt = 'Retained off-paper drawing';
      image.dataset.imageFloating = JSON.stringify({ horizontalPosition: { relative: 'page', offset: x * 9525 }, verticalPosition: { relative: 'page', offset: y * 9525 }, wrap: { type: 1 } });
      return { html: `<p style="line-height:24px;margin:0">BEFORE</p><p style="line-height:24px;margin:0">${image.outerHTML}ANCHOR</p><p style="line-height:24px;margin:0">AFTER</p>`, image: image.src, paper };
    }, { x, y, paper });
    await page.locator('input[type=file]').first().setInputFiles({ name: 'off-paper.html', mimeType: 'text/html', buffer: Buffer.from(source.html) });
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
    await expect(editor).toHaveText('BEFOREANCHORAFTER');
    await expect(editor.locator('img[data-image-floating]')).toBeHidden();
    await expect(page.locator('.document-stack')).toHaveAttribute('data-pages', '1');
    const result = await page.evaluate(async source => {
      const { importFile, exportDocx } = await import('/src/document-io.ts');
      const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
      const { createPrintLayout } = await import('/src/print-layout.ts');
      const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
      const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
      const loaded = await importFile(new File([source.html], 'off-paper.html'));
      let doc: any = { ...loaded.document, page: source.paper, content: parseDocumentHtml(loaded.html!) };
      const snapshots = [], files = [];
      for (let save = 0; save <= 2; save++) {
        const images: any[] = []; const walk = (n: any) => { if (['image', 'inlineImage'].includes(n.type)) images.push(n.attrs); n.content?.forEach(walk); }; walk(doc.content);
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const img = layout.element.querySelector('.print-content img') as HTMLElement;
          snapshots.push({ save, images: images.map(i => ({ src: i.src, geometry: JSON.parse(i.floating) })), pages: layout.pages,
            visible: !!img && getComputedStyle(img).display !== 'none' && img.getBoundingClientRect().width > 0,
            text: [...layout.element.querySelectorAll('.print-content')].map(p => p.textContent).join('') });
        } finally { layout.dispose(); }
        if (save < 2) {
          const blob = await exportDocx(doc), zip = await JSZip.loadAsync(await blob.arrayBuffer());
          files.push({ bytes: [...new Uint8Array(await blob.arrayBuffer())], xml: await zip.file('word/document.xml')!.async('string') });
          const reopened = await importFile(new File([blob], 'off-paper.docx')); doc = { ...reopened.document, content: parseDocumentHtml(reopened.html!) };
        }
      }
      return { snapshots, files };
    }, source);
    for (const row of result.snapshots) {
      expect(row.pages).toBe(1); expect(row.visible).toBe(false); expect(row.text).toBe('BEFOREANCHORAFTER');
      expect(row.images).toHaveLength(1); expect(row.images[0].src).toBe(source.image);
      expect(row.images[0].geometry.horizontalPosition).toEqual({ relative: 'page', offset: x * 9525 });
      expect(row.images[0].geometry.verticalPosition).toEqual({ relative: 'page', offset: y * 9525 });
    }
    for (const [i, file] of result.files.entries()) {
      expect(file.xml).not.toContain('display:none');
      expect(file.xml).toContain(`<wp:posOffset>${x * 9525}</wp:posOffset>`);
      expect(file.xml).toContain(`<wp:posOffset>${y * 9525}</wp:posOffset>`);
      const output = testInfo.outputPath(`textbox-save-offpaper-${side}-${i + 1}.docx`);
      await mkdir(path.dirname(output), { recursive: true }); await writeFile(output, Buffer.from(file.bytes));
    }
  });
}

test('partial page ink, alignment anchors and paragraph-relative offsets are not treated as wholly off paper', async ({ page }) => {
  await page.goto('/');
  const cases = await page.evaluate(async () => {
    const { imageOutsidePaper } = await import('/src/image-floating.ts');
    const sheet = document.createElement('section'); sheet.dataset.pageHeight = '1122.52'; sheet.style.cssText = 'position:fixed;left:-10000px;width:793.7px;height:1122.52px'; document.body.append(sheet);
    try {
      const values = [
        { horizontalPosition: { relative: 'page', offset: -20 * 9525 }, verticalPosition: { relative: 'page', offset: 50 * 9525 } },
        { horizontalPosition: { relative: 'page', offset: 780 * 9525 }, verticalPosition: { relative: 'page', offset: 50 * 9525 } },
        { horizontalPosition: { relative: 'column', offset: 900 * 9525 }, verticalPosition: { relative: 'paragraph', offset: 1400 * 9525 } },
        { horizontalPosition: { relative: 'page', align: 'center', offset: 900 * 9525 }, verticalPosition: { relative: 'page', align: 'center', offset: 1400 * 9525 } },
      ];
      return values.map(v => { const img = document.createElement('img'); img.width = img.height = 40; img.dataset.imageFloating = JSON.stringify({ ...v, wrap: { type: 1 } }); sheet.append(img); return imageOutsidePaper(img); });
    } finally { sheet.remove(); }
  });
  expect(cases).toEqual([false, false, false, false]);
});

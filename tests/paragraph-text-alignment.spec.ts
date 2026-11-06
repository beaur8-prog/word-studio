import { test, expect } from '@playwright/test';
for (const override of [false, true]) test(`standard paragraph top alignment ${override ? 'respects an explicit baseline override' : 'is inherited from a Word style'}`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async override => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const canvas = document.createElement('canvas'); canvas.width = 80; canvas.height = 40; const ctx = canvas.getContext('2d')!; ctx.fillStyle = '#cc3344'; ctx.fillRect(0, 0, 80, 40);
    const src = canvas.toDataURL(), html = `<p data-word-text-alignment="top" style="margin:0;font-size:14pt;line-height:28.93px"><img data-inline-image="true" src="${src}" width="80" height="40" alt="SOURCE LOGO"></p><p style="margin:0">SOURCE BODY</p>`;
    const loaded = await importFile(new File([html], 'alignment.html')), doc = { ...loaded.document, page: { ...defaultPage, pageNumbers: false }, content: parseDocumentHtml(html) };
    const zip = await JSZip.loadAsync(await (await exportDocx(doc)).arrayBuffer()), W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const parser = new DOMParser(), xml = parser.parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
    const p = xml.getElementsByTagNameNS(W, 'p')[0], props = p.getElementsByTagNameNS(W, 'pPr')[0], setting = props.getElementsByTagNameNS(W, 'textAlignment')[0];
    if (setting?.getAttributeNS(W, 'val') !== 'top') throw Error('Export must use standard w:textAlignment');
    const styles = parser.parseFromString(await zip.file('word/styles.xml')!.async('string'), 'application/xml');
    const style = styles.createElementNS(W, 'w:style'); style.setAttributeNS(W, 'w:type', 'paragraph'); style.setAttributeNS(W, 'w:styleId', 'LogoTop');
    const name = styles.createElementNS(W, 'w:name'); name.setAttributeNS(W, 'w:val', 'Logo top'); style.append(name);
    const base = styles.createElementNS(W, 'w:basedOn'); base.setAttributeNS(W, 'w:val', 'Normal'); style.append(base);
    const pr = styles.createElementNS(W, 'w:pPr'), top = styles.createElementNS(W, 'w:textAlignment'); top.setAttributeNS(W, 'w:val', 'top'); pr.append(top); style.append(pr); styles.documentElement.append(style);
    let pstyle = props.getElementsByTagNameNS(W, 'pStyle')[0]; if (!pstyle) { pstyle = xml.createElementNS(W, 'w:pStyle'); props.insertBefore(pstyle, props.firstChild); } pstyle.setAttributeNS(W, 'w:val', 'LogoTop');
    if (override) setting.setAttributeNS(W, 'w:val', 'baseline'); else setting.remove();
    // No ignorable Studio paragraph attribute carries this alignment. Its
    // standard style/direct properties are the sole source of truth.
    for (const a of [...p.attributes]) if (a.namespaceURI?.startsWith('urn:wordstudio:')) p.removeAttributeNS(a.namespaceURI, a.localName);
    zip.file('word/document.xml', new XMLSerializer().serializeToString(xml)); zip.file('word/styles.xml', new XMLSerializer().serializeToString(styles));
    let imported = await importFile(new File([await zip.generateAsync({ type: 'blob' })], 'external.docx'));
    const states = [];
    for (let n = 0; n < 3; n++) {
      const doc = { ...imported.document, content: parseDocumentHtml(imported.html!) }, layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try {
        const image = layout.element.querySelector('img')!, paragraph = image.closest('p')!;
        states.push({ policy: doc.content.content![0].attrs?.wordTextAlignment, image: image.getAttribute('src'), height: paragraph.getBoundingClientRect().height, imageHeight: image.getBoundingClientRect().height, align: getComputedStyle(image).verticalAlign, text: layout.element.textContent });
      } finally { layout.dispose(); }
      if (n < 2) imported = await importFile(new File([await exportDocx(doc)], 'again.docx'));
    }
    return { states, src };
  }, override);
  for (const state of result.states) {
    expect(state.policy).toBe(override ? 'baseline' : 'top'); expect(state.image).toBe(result.src); expect(state.text).toContain('SOURCE BODY');
    expect(state.imageHeight).toBeCloseTo(40, 0); expect(state.align).toBe(override ? 'baseline' : 'top');
    if (override) expect(state.height).toBeGreaterThan(45); else expect(state.height).toBeCloseTo(40, 0);
  }
});

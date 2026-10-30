import { test, expect } from '@playwright/test';

for (const kind of ['dxa', 'pct']) test(`${kind}: an indented table keeps its column widths through two DOCX saves`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async kind => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const results = [];
    for (const indent of [0, 80]) {
      const html = `<table style="margin-left:${indent}px"><tr><td colwidth="300" style="padding:4px"><p style="font-family:Arial;font-size:12pt;line-height:20px;margin:0">LEFT</p></td>` +
        '<td colwidth="300" style="padding:4px"><p style="font-family:Arial;font-size:12pt;line-height:20px;margin:0">RIGHT</p></td></tr></table>';
      const opened = await importFile(new File([html], 'source.html'));
      const original = { ...opened.document, content: parseDocumentHtml(opened.html!) };
      let blob = await exportDocx(original);
      if (kind === 'pct') {
        const zip = await JSZip.loadAsync(await blob.arrayBuffer());
        const xml = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
        const width = xml.getElementsByTagNameNS(W, 'tblW')[0];
        width.setAttributeNS(W, 'w:type', 'pct'); width.setAttributeNS(W, 'w:w', '2500');
        zip.file('word/document.xml', new XMLSerializer().serializeToString(xml));
        blob = await zip.generateAsync({ type: 'blob' });
      }
      const states = [];
      for (let save = 0; save < 3; save++) {
        const current = await importFile(new File([blob], 'again.docx'));
        const doc = { ...current.document, content: parseDocumentHtml(current.html!) };
        const table = doc.content.content!.find(n => n.type === 'table')!;
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          let paperOverflow = 0;
          for (const p of layout.element.querySelectorAll('.print-content p')) {
            const paper = p.closest('.print-page')!.getBoundingClientRect();
            const r = document.createRange(); r.selectNodeContents(p);
            for (const box of r.getClientRects()) paperOverflow = Math.max(paperOverflow, paper.left - box.left, box.right - paper.right);
          }
          states.push({ save, indent: parseFloat(table.attrs!.tableIndent || '0'),
            widths: table.content![0].content!.flatMap(c => c.attrs!.colwidth), pages: layout.pages,
            paperOverflow, text: layout.element.querySelector('.print-content')!.textContent });
        } finally { layout.dispose(); }
        if (save < 2) blob = await exportDocx(doc);
      }
      results.push({ indent, states });
    }
    return results;
  }, kind);
  const expected = kind === 'dxa' ? [300, 300] : result[0].states[0].widths;
  for (const row of result) for (const state of row.states) {
    expect(state.widths).toEqual(expected);
    expect(state.indent).toBeCloseTo(row.indent, 1);
    expect(state.pages).toBe(1); expect(state.paperOverflow).toBeLessThanOrEqual(.5);
    expect(state.text).toContain('LEFT'); expect(state.text).toContain('RIGHT');
  }
});

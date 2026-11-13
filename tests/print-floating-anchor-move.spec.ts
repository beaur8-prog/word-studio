import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';

for (const wrap of [1, 3]) for (const paintedAnchor of [false, true]) test(`a floating drawing moved to new paper reflows both pages (wrap ${wrap}, ${paintedAnchor ? 'painted' : 'empty'} anchor), including two DOCX saves`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ wrap, paintedAnchor }) => {
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const canvas = document.createElement('canvas'); canvas.width = canvas.height = 4;
    const context = canvas.getContext('2d')!; context.fillStyle = '#145cff'; context.fillRect(0, 0, 4, 4);
    const image = canvas.toDataURL(), geometry = { horizontalPosition: { relative: 'page', offset: 76 * 9525 },
      verticalPosition: { relative: 'page', offset: 142 * 9525 }, wrap: { type: wrap, side: 'bothSides' }, layoutInCell: true };
    const paragraph = (text: string) => `<p style="font-family:Arial;font-size:11pt;line-height:24px;margin:0" data-word-line-rule="exact">${text}</p>`;
    const html = paragraph('') + paragraph('DRAWING HEADING') + `<p style="font-family:Arial;font-size:11pt;line-height:1px;margin:0" data-word-line-rule="exact"><img data-inline-image="true" src="${image}" width="640" height="890" alt="COMPLETE DRAWING" data-image-floating="${JSON.stringify(geometry).replaceAll('"', '&quot;')}">${paintedAnchor ? 'ANCHOR' : ''}</p><div data-page-break="true"></div>` + paragraph('AFTER DRAWING');
    let doc: any = { format: 'word-studio', version: 1, id: 'moved-drawing', title: 'Moved drawing', comments: [], updatedAt: new Date().toISOString(),
      page: { ...defaultPage, margins: { top: 25, bottom: 20, left: 20, right: 20 }, pageNumbers: false }, content: parseDocumentHtml(html) };
    const states = [], saves = [];
    for (let pass = 0; pass < 3; pass++) {
      const source = document.createElement('div'); source.innerHTML = generateHTML(doc.content, editorExtensions);
      const sourceImage = source.querySelector('img')!;
      const layout = await createPrintLayout(source.innerHTML, doc.page);
      try {
        const point = (element: Element) => { const paper = element.closest<HTMLElement>('.print-page')!, r = element.getBoundingClientRect(), p = paper.getBoundingClientRect();
          return { page: Number(paper.dataset.page), x: r.left - p.left, y: r.top - p.top, width: r.width, height: r.height, paperHeight: p.height }; };
        const heading = [...layout.element.querySelectorAll('.print-content p')].find(p => p.textContent === 'DRAWING HEADING')!;
        const after = [...layout.element.querySelectorAll('.print-content p')].find(p => p.textContent === 'AFTER DRAWING')!;
        const drawing = layout.element.querySelector('img[alt="COMPLETE DRAWING"]')!;
        states.push({ pages: layout.pages, heading: point(heading), after: point(after), drawing: point(drawing),
          text: source.textContent, sourceImage: { src: sourceImage.getAttribute('src'), geometry: (() => { const g = JSON.parse(sourceImage.getAttribute('data-image-floating')!); return { horizontal: g.horizontalPosition, vertical: g.verticalPosition, wrap: g.wrap.type }; })(), width: sourceImage.getAttribute('width'), height: sourceImage.getAttribute('height') } });
      } finally { layout.dispose(); }
      if (pass < 2) { const blob = await exportDocx(doc); saves.push([...new Uint8Array(await blob.arrayBuffer())]);
        const opened = await importFile(new File([blob], 'moved-drawing.docx')); doc = { ...opened.document, content: parseDocumentHtml(opened.html!) }; }
    }
    return { states, saves };
  }, { wrap, paintedAnchor });
  for (const [index, bytes] of result.saves.entries()) await writeFile(test.info().outputPath(`textbox-save-moved-drawing-${wrap}-${index + 1}.docx`), Buffer.from(bytes));
  for (const state of result.states) {
    expect(state.pages).toBe(3); expect(state.heading.page).toBe(1); expect(state.heading.y).toBeCloseTo(25 * 96 / 25.4 + 24, 0);
    expect(state.drawing.page).toBe(2); expect(state.drawing.y).toBeCloseTo(142, 0); expect(state.drawing.height).toBeCloseTo(890, 0);
    expect(state.drawing.y).toBeGreaterThanOrEqual(0); expect(state.drawing.y + state.drawing.height).toBeLessThan(state.drawing.paperHeight);
    expect(state.after.page).toBe(3); expect(state.text).toBe(result.states[0].text); expect(state.sourceImage).toEqual(result.states[0].sourceImage);
  }
  await writeFile(test.info().outputPath('moved-drawing.json'), JSON.stringify(result.states, null, 2));
});

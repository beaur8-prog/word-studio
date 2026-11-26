import { expect, test } from '@playwright/test';

test('witnessed HWP row spacing and inline pair preserve three logical tables through DOCX saves', async ({ page }) => {
  await page.goto('/');
  const results = await page.evaluate(async () => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { hwpRelatedTableLayoutAttribute, hwpRelatedParagraphToken, hwpRelatedTableToken } =
      await import('/src/hwp-related-table-layout.ts');
    const { layoutHwpRelatedTables } = await import('/src/hwp-related-table-print.ts');
    const { pageMargins, paperDimensions } = await import('/src/types.ts');

    const paragraph = (text: string) =>
      `<p data-exact-line-height="true" style="line-height:22.4px;font-size:16px;margin:0">${text}</p>`;
    const row = (left: string, right: string) =>
      `<tr><td colwidth="200">${left}</td><td colwidth="240">${right}</td></tr>`;
    const text = [['left start', 'left tail'], ['right start', 'right tail']];
    const introduction = row(paragraph('A'), paragraph('B'));
    const witnessedRow = row(paragraph('left<br> start') + paragraph(text[0][1]),
      paragraph(text[1][0]) + paragraph(text[1][1]));
    const before = `<table>${introduction}${witnessedRow}${row(paragraph('C'), paragraph('D'))}</table>`;
    const member = (label: string, width: number, offset: number) =>
      `<table style="margin-left:${offset}px"><tr><td colwidth="${width}">${paragraph(label)}</td></tr></table>`;
    const carrier = '<p data-exact-line-height="true" style="line-height:1px;font-size:2px;margin:0"></p>';
    const root = document.createElement('div');
    root.innerHTML = before + member('Left original', 160, 20) + carrier + member('Right original', 190, 226);
    const plain = await importFile(new File([root.innerHTML], 'related.html'));
    const [w, h] = paperDimensions(plain.document.page), margins = pageMargins(plain.document.page);
    const proof = {
      bodyWidth: (w - margins.left - margins.right) * 96 / 25.4,
      bodyHeight: (h - margins.top - margins.bottom) * 96 / 25.4,
      rows: 3, row: 1, rowWidths: [200, 240],
      before: text.map((cell, col) => cell.map((literal, p) => ({
        gap: 14, text: hwpRelatedParagraphToken(literal), font: 16, pitch: 22.4, breaks: col === 0 && p === 0 ? 1 : 0,
      }))),
      pair: [
        { left: 20, width: 160, rows: 1, text: hwpRelatedTableToken(['Left original']) },
        { left: 226, width: 190, rows: 1, text: hwpRelatedTableToken(['Right original']) },
      ],
    };
    root.querySelectorAll('table')[1].setAttribute(hwpRelatedTableLayoutAttribute, JSON.stringify(proof));
    const loaded = await importFile(new File([root.innerHTML], 'related.html'));
    let doc = { ...loaded.document, content: parseDocumentHtml(loaded.html) };

    const states = [];
    for (const stage of [0, 1, 2]) {
      const model = document.createElement('div');
      model.innerHTML = generateHTML(doc.content, editorExtensions);
      const probe = (change?: (copy: HTMLElement) => void) => {
        const copy = model.cloneNode(true) as HTMLElement;
        change?.(copy);
        layoutHwpRelatedTables(copy, doc.page);
        const group = copy.querySelector('table[data-hwp-related-print-group]');
        return {
          group: group ? [...group.querySelectorAll('table')].map(table => table.textContent) : [],
          margins: [...copy.querySelector('table')!.rows[1].querySelectorAll('td > p')]
            .map(p => (p as HTMLParagraphElement).style.marginTop),
          br: copy.querySelector('table')!.rows[1].querySelectorAll('br').length,
        };
      };
      const snapshot = [...model.querySelectorAll('table')].map(table => ({
        rows: table.rows.length,
        columns: [...table.querySelectorAll<HTMLTableColElement>(':scope > colgroup > col')].map(c => c.style.width),
      }));
      const positive = probe();
      const altered = [
        probe(copy => { copy.querySelector('table')!.rows[1].querySelector('p')!.textContent += '!'; }),
        probe(copy => { copy.querySelector('table')!.rows[1].querySelector('br')!.remove(); }),
        probe(copy => { copy.querySelectorAll('table')[2].style.marginLeft = '245px'; }),
      ];
      const untouched = !model.querySelector('[data-hwp-related-print-group]') && model.querySelectorAll('table').length === 3;
      states.push({ snapshot, positive, altered, untouched });
      if (stage < 2) {
        const reopened = await importFile(new File([await exportDocx(doc)], 'related.docx'));
        doc = { ...reopened.document, content: parseDocumentHtml(reopened.html) };
      }
    }
    return states;
  });
  expect(results).toHaveLength(3);
  for (const result of results) {
    expect(result.snapshot).toEqual([
      { rows: 3, columns: ['200px', '240px'] },
      { rows: 1, columns: ['160px'] },
      { rows: 1, columns: ['190px'] },
    ]);
    expect(result.positive).toEqual({ group: ['Left original', 'Right original'], margins: Array(4).fill('14px'), br: 1 });
    expect(result.altered.every(candidate => !candidate.group.length && !candidate.margins.includes('14px'))).toBe(true);
    expect(result.untouched).toBe(true);
  }
});

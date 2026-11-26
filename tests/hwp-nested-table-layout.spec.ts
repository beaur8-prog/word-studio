import { expect, test } from '@playwright/test';

test('witnessed HWP nested cuts preserve a logical merged table through DOCX saves', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const {
      hwpNestedCellToken, hwpNestedLayoutFingerprint, hwpNestedParagraphToken,
      hwpNestedTableLayoutAttribute, hwpNestedTableToken,
    } = await import('/src/hwp-nested-table-layout.ts');
    const { layoutHwpNestedTables } = await import('/src/hwp-nested-table-print.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const labels = ['Start here', 'Then continue here', 'Final paragraph'];
    const nestedCells = [
      [0, 0, 3, 1, 'Merged content'], [0, 1, 1, 1, 'First'],
      [1, 1, 1, 1, 'Second'], [2, 1, 1, 1, 'Third'],
    ] as const;
    const profile = {
      blocks: 4,
      fingerprint: hwpNestedLayoutFingerprint([
        hwpNestedParagraphToken(labels[0]), hwpNestedTableToken(3, 2),
        ...labels.slice(1).map(hwpNestedParagraphToken),
      ]),
      gaps: [[0, 6.65, 13.35, null]],
      cuts: [{
        block: 1, rows: 3, columns: 2, boundary: 2, merges: [[0, 0, 3]],
        fingerprint: hwpNestedLayoutFingerprint(nestedCells.map(cell => hwpNestedCellToken(
          cell[0], cell[1], cell[2], cell[3], cell[4],
        ))),
      }],
    };
    const root = document.createElement('div');
    root.innerHTML = '<table><tr><td colwidth="240"><p>Start here</p>'
      + '<table><tr><td colwidth="90" rowspan="3"><p>Merged content</p></td><td colwidth="90"><p>First</p></td></tr>'
      + '<tr><td colwidth="90"><p>Second</p></td></tr><tr><td colwidth="90"><p>Third</p></td></tr></table>'
      + '<p>Then continue here</p><p>Final paragraph</p></td></tr></table>';
    root.querySelector('table')!.setAttribute(hwpNestedTableLayoutAttribute, JSON.stringify(profile));
    const loaded = await importFile(new File([root.innerHTML], 'nested.html'));
    let doc = { ...loaded.document, content: parseDocumentHtml(loaded.html) };
    const states = [];
    for (const stage of [0, 1, 2]) {
      const model = document.createElement('div');
      model.innerHTML = generateHTML(doc.content, editorExtensions);
      const outer = model.querySelector('table')!, cell = outer.rows[0].cells[0];
      const child = cell.querySelector('table')!;
      const snapshot = { tableCount: model.querySelectorAll('table').length,
        rows: child.rows.length, merge: child.rows[0].cells[0].rowSpan, text: cell.textContent,
        metadata: !!outer.getAttribute(hwpNestedTableLayoutAttribute) };
      const copy = model.cloneNode(true) as HTMLElement;
      layoutHwpNestedTables(copy);
      const copies = [...copy.querySelectorAll('td > table')];
      const continued = copies.map(table => {
        const typed = table as HTMLTableElement;
        return [typed.rows.length, typed.rows[0].cells[0].rowSpan, typed.rows[0].cells[0].textContent];
      });
      const after = copy.querySelector('td')!;
      const changed = model.cloneNode(true) as HTMLElement;
      changed.querySelector('td > table td p')!.textContent += ' edited';
      layoutHwpNestedTables(changed);
      states.push({ snapshot, continued, renderedText: after.textContent,
        gap: (after.firstElementChild as HTMLElement).style.marginBottom,
        editedColumnCount: changed.querySelectorAll('td > table').length,
        modelUnmodified: child.rows.length === 3 && child.rows[0].cells[0].rowSpan === 3 });
      if (stage < 2) {
        const file = await importFile(new File([await exportDocx(doc)], 'roundtrip.docx'));
        doc = { ...file.document, content: parseDocumentHtml(file.html) };
      }
    }
    return states;
  });
  expect(result).toHaveLength(3);
  for (const stage of result) {
    expect(stage.snapshot).toEqual({ tableCount:2, rows:3, merge:3, metadata:true,
      text:'Start hereMerged contentFirstSecondThirdThen continue hereFinal paragraph' });
    expect(stage.continued).toEqual([[2, 2, 'Merged content'], [1, 1, '']]);
    expect(stage.renderedText).toBe(stage.snapshot.text);
    expect(stage.gap).toBe('20px');
    expect(stage.editedColumnCount).toBe(1);
    expect(stage.modelUnmodified).toBe(true);
  }
});

test('native export restores only the faces supplied for pagination', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { clearHwpNestedCellFaceProjection } = await import('/src/hwp-nested-cell-face-print');
    const { hwpSingleBodyFace } = await import('/src/hwp-text-advances');
    const root = document.createElement('div');
    const paragraph = document.createElement('p');
    const measured = document.createElement('span'), authored = document.createElement('span');
    paragraph.dataset.hwpNestedFaceApplied = 'true';
    measured.setAttribute('data-hwp-nested-face-for-measure', '');
    for (const node of [measured, authored]) {
      node.style.fontFamily = `"${hwpSingleBodyFace}", "Noto Sans KR"`;
      node.style.fontSize = '12px';
    }
    measured.textContent = '조판용';
    authored.textContent = '원래 글꼴';
    paragraph.append(measured, ' ', authored);
    root.append(paragraph);
    const copy = root.cloneNode(true) as HTMLElement;
    clearHwpNestedCellFaceProjection(copy);
    const spans = copy.querySelectorAll('span');
    return {
      original: measured.style.fontFamily,
      sourceFace: hwpSingleBodyFace,
      restored: spans[0].style.fontFamily.replace(/^["']|["']$/g, ''),
      authored: spans[1].style.fontFamily,
      sizes: Array.from(spans, node => node.style.fontSize),
      text: copy.textContent,
      markers: copy.querySelectorAll('[data-hwp-nested-face-for-measure],[data-hwp-nested-face-applied]').length,
    };
  });
  expect(result.restored).toBe(result.sourceFace);
  expect(result.authored).toBe(result.original);
  expect(result.sizes).toEqual(['12px', '12px']);
  expect(result.text).toBe('조판용 원래 글꼴');
  expect(result.markers).toBe(0);
});

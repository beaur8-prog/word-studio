import { test, expect } from '@playwright/test';

for (const format of ['hwp', 'hwpx'] as const) test(`${format}: body and nested cell notes remain native controls and editable stories after two saves`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async format => {
    const { exportEditableHwp, importEditableHwp } = await import('/src/hwp/editing.ts');
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const ref = (id: string, kind = 'footnote', label = '1') => `<span data-note-reference="${id}" data-note-kind="${kind}" data-note-label="${label}">${label}</span>`;
    const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';
    let doc: any = { format: 'word-studio', version: 1, id: crypto.randomUUID(), title: 'Notes', comments: [], updatedAt: new Date().toISOString(),
      page: { size: 'A4', orientation: 'portrait', margin: 20, header: '', footer: '', pageNumbers: false },
      content: parseDocumentHtml(`<p>BODY 😀 BEFORE${ref('body')}AFTER</p><table><tr><td><p>CELL BEFORE${ref('cell', 'endnote')}AFTER</p><table><tr><td><p>NESTED BEFORE${ref('nested', 'footnote', '2')}AFTER</p></td></tr></table></td></tr></table>
        <div data-document-note="body" data-note-kind="footnote" data-note-label="1"><p><strong>Footnote 굵게 😀</strong></p><p style="text-align:right">Second paragraph<img data-inline-image="true" src="data:image/png;base64,${png}" width="20" height="20"></p></div>
        <div data-document-note="cell" data-note-kind="endnote" data-note-label="1"><p><em>Endnote 기울임</em></p></div>
        <div data-document-note="nested" data-note-kind="footnote" data-note-label="2"><p>Nested explanation</p><table><tr><td><p>Note table</p></td></tr></table></div>`) };
    const states = [];
    for (let round = 0; round < 2; round++) {
      const exported = await exportEditableHwp(doc, format), session = new HwpReaderSession();
      try { await session.open(new File([exported.blob], 'notes.' + format)); doc = importEditableHwp(await session.editable(), 'notes.' + format); }
      finally { session.close(); }
      const html = generateHTML(doc.content, editorExtensions), root = document.createElement('div'); root.innerHTML = html;
      const refs = [...root.querySelectorAll<HTMLElement>('[data-note-reference]')].map(el => ({ id: el.dataset.noteReference, kind: el.dataset.noteKind, label: el.dataset.noteLabel, parent: el.parentElement!.textContent }));
      const notes = [...root.querySelectorAll<HTMLElement>('[data-document-note]')].map(el => ({ id: el.dataset.documentNote, kind: el.dataset.noteKind, label: el.dataset.noteLabel, text: el.textContent!.replace(/\s+/g, ' ').trim(), bold: !!el.querySelector('strong'), italic: !!el.querySelector('em'), images: el.querySelectorAll('img').length, tables: el.querySelectorAll('table').length }));
      const reopened = await importFile(new File([await exportDocx(doc)], 'notes.docx'));
      const docx = document.createElement('div'); docx.innerHTML = reopened.html!;
      states.push({ refs, notes, warnings: exported.warnings, html, docxRefs: docx.querySelectorAll('[data-note-reference]').length, docxNotes: docx.querySelectorAll('[data-document-note]').length });
    }
    return states;
  }, format);
  for (const state of result) {
    expect(state.refs).toHaveLength(3); expect(state.notes).toHaveLength(3);
    expect(state.docxRefs).toBe(3); expect(state.docxNotes).toBe(3);
    expect(state.refs.map(r => r.parent)).toEqual(['BODY 😀 BEFORE1AFTER', 'CELL BEFORE1AFTER', 'NESTED BEFORE2AFTER']);
    for (const ref of state.refs) expect(state.notes.some(n => n.id === ref.id && n.kind === ref.kind && n.label === ref.label)).toBe(true);
    expect(state.notes.find(n => n.text.includes('Footnote'))).toMatchObject({ bold: true, images: 1 });
    expect(state.notes.find(n => n.text.includes('Endnote'))).toMatchObject({ italic: true, kind: 'endnote' });
    expect(state.notes.find(n => n.text.includes('Nested explanation'))).toMatchObject({ tables: 1 });
    expect(state.html).not.toMatch(/WORDSTUDIO-NOTE|변환하지 못한|원본 보기에서 확인/);
    expect(state.warnings.join('\n')).not.toContain('내용 손실');
    expect(state.warnings.join('\n')).not.toContain('각주·미주는 본문');
  }
});

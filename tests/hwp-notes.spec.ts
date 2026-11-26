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
      let savedDoc = doc;
      const docxSaves = [];
      for (let stage = 0; stage < 2; stage++) {
        const reopened = await importFile(new File([await exportDocx(savedDoc)], 'notes.docx'));
        savedDoc = { ...reopened.document, content: parseDocumentHtml(reopened.html!) };
        const docx = document.createElement('div'); docx.innerHTML = generateHTML(savedDoc.content, editorExtensions);
        const reopenedRefs = [...docx.querySelectorAll<HTMLElement>('[data-note-reference]')];
        const reopenedNotes = [...docx.querySelectorAll<HTMLElement>('[data-document-note]')];
        docxSaves.push({
          refs: reopenedRefs.map(ref => [ref.dataset.noteKind, ref.closest('sup')?.tagName ?? null, ref.textContent]),
          stories: reopenedNotes.map(note => [note.dataset.noteKind, note.textContent?.trim().replace(/\s+/g, ' ')]),
        });
      }
      states.push({ refs, notes, warnings: exported.warnings, html, docxSaves });
    }
    return states;
  }, format);
  for (const state of result) {
    expect(state.refs).toHaveLength(3); expect(state.notes).toHaveLength(3);
    for (const save of state.docxSaves) {
      expect(save.refs).toEqual([['footnote', null, '[1]'], ['endnote', null, '[2]'], ['footnote', null, '[3]']]);
      expect(save.stories).toHaveLength(3);
      expect(save.stories.find(([kind]) => kind === 'endnote')).toEqual(['endnote', 'Endnote 기울임']);
      expect(save.stories.some(([kind, text]) => kind === 'footnote' && text?.includes('Footnote 굵게 😀'))).toBe(true);
      expect(save.stories.some(([kind, text]) => kind === 'footnote' && text?.includes('Nested explanation'))).toBe(true);
    }
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

test('a plain HWP footnote keeps its native label and spacing through two standard DOCX saves', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { unzipSync, strFromU8 } = await import('/node_modules/.vite/deps/fflate.js');
    const native = { mark: '1)', above: 11, below: 8, between: 4, separator: 185, stroke: .5 };
    const proof = JSON.stringify(native).replace(/"/g, '&quot;');
    const marker = 'data-note-reference="owned" data-note-kind="footnote" data-note-label="1"';
    const first = '<p>Before<span ' + marker + ' data-hwp-footnote-mark="1)">1)</span>after</p>';
    const noteStyle = 'font-family:Arial;font-size:12px;line-height:15.6px;margin:0';
    const note = '<div data-document-note="owned" data-note-kind="footnote" data-note-label="1" ' +
      'data-hwp-footnote-layout="' + proof + '"><p style="' + noteStyle + '">Authored note 한글</p></div>';
    const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const ns = 'urn:wordstudio:hwp-footnote-layout:1';
    let documentModel: any = { format: 'word-studio', version: 1, id: crypto.randomUUID(), title: 'Footnote layout',
      comments: [], updatedAt: new Date().toISOString(), page: defaultPage, content: parseDocumentHtml(first + note) };
    const states = [], packages = [];
    for (let stage = 0; stage < 3; stage++) {
      const print = await createPrintLayout(generateHTML(documentModel.content, editorExtensions), documentModel.page);
      try {
        const box = print.element.querySelector<HTMLElement>('.print-footnotes')!;
        const reference = print.element.querySelector<HTMLElement>('[data-note-reference]')!;
        const note = box.querySelector<HTMLElement>('p')!;
        states.push({ pages: print.pages, native: box.dataset.hwpFootnoteLayoutActive, noteLabel: note.dataset.hwpFootnoteLabel,
          noteText: note.textContent, bodyMark: reference.textContent,
          separator: getComputedStyle(box, '::before').width, gap: getComputedStyle(box).paddingTop });
      } finally { print.dispose(); }
      if (stage === 2) break;
      const file = await exportDocx(documentModel);
      const parts = unzipSync(new Uint8Array(await file.arrayBuffer()));
      const xml = (part: string) => new DOMParser().parseFromString(strFromU8(parts[part]), 'application/xml');
      const body = xml('word/document.xml');
      const standard = [...body.getElementsByTagNameNS(w, 'footnoteReference')];
      const id = standard[0]?.getAttributeNS(w, 'id');
      const saved = [...xml('word/footnotes.xml').getElementsByTagNameNS(w, 'footnote')].find(
        entry => entry.getAttributeNS(w, 'id') === id);
      packages.push({
        refs: standard.length, literalBody: [...body.getElementsByTagNameNS(w, 't')].map(t => t.textContent).join(''),
        attrs: JSON.parse(saved?.getAttributeNS(ns, 'layout') || 'null'),
      });
      const reopened = await importFile(new File([file], 'footnote-layout.docx'));
      documentModel = { ...reopened.document, content: parseDocumentHtml(reopened.html!) };
    }
    const complex = parseDocumentHtml(first + note.replace('</div>', '<p>More note content</p></div>'));
    const legacy = await createPrintLayout(generateHTML(complex, editorExtensions), defaultPage);
    try {
      return { states, packages, native, activeOnComplexNote: legacy.element.querySelector('.print-footnotes')?.hasAttribute('data-hwp-footnote-layout-active') };
    } finally { legacy.dispose(); }
  });
  expect(result.states).toHaveLength(3);
  for (const [stage, state] of result.states.entries()) {
    expect(state).toMatchObject({ pages: 1, native: 'true', noteLabel: '1) ', noteText: 'Authored note 한글',
      bodyMark: '1)', separator: '185px', gap: '8px' });
    if (stage < 2) expect(result.packages[stage]).toEqual({ refs: 1, literalBody: 'Beforeafter', attrs: result.native });
  }
  expect(result.activeOnComplexNote).toBe(false);
});

test('a proved body-end pair fits above its own notes without changing either story', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { sourceFootnoteEndBottom } = await import('/src/print-hwp-note-terminal.ts');
    const { clearSourceBodyEndPaint } = await import('/src/hwp-body-line-metrics.ts');
    const { hwpWrapFingerprint, normalizeHwpWrapText, parseHwpSourceWrap } = await import('/src/hwp-text-advances.ts');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
    const fingerprint = (text: string) => hwpWrapFingerprint(normalizeHwpWrapText(text));
    const container = document.createElement('div');
    container.lang = 'en';
    container.innerHTML = `<div data-body-height="440" style="width:240px;height:370px;font:20px/34px Arial">
      <p style="height:240px;margin:0"></p><p style="margin:0">Before the note</p><p style="height:22px;margin:0"></p>
      <p style="width:220px;margin:0 0 0 20px;text-indent:-20px">Editorial
      <span data-note-reference="one" data-hwp-footnote-mark="1)">1)</span>
      <span data-note-reference="two" data-hwp-footnote-mark="2)">2)</span>
      advice studies this case.</p></div>
      <div data-hwp-footnote-layout-active="true">
        <div data-document-note="one">First note</div><div data-document-note="two">Second note</div>
      </div>`;
    document.body.append(container);
    try {
      const body = container.firstElementChild as HTMLElement, notes = container.lastElementChild as HTMLElement;
      const paragraph = body.lastElementChild as HTMLElement;
      const stories = [...notes.children] as HTMLElement[];
      const refs = [...paragraph.querySelectorAll('[data-note-reference]')];
      const literal = paragraph.cloneNode(true) as HTMLElement;
      literal.querySelectorAll('[data-note-reference]').forEach(node => node.remove());
      const source = normalizeHwpWrapText(literal.textContent);
      const lineTop = paragraph.getBoundingClientRect().top - body.getBoundingClientRect().top;
      const wrap = { fingerprint: fingerprint(source), ends: [20, source.length], noteTerminal: {
        width: 240, fullBodyHeight: 440, lineTop, pitch: 34, font: 20, hanging: 20,
        secondBaseline: lineTop + 34 + 16, prior: fingerprint('Before the note'),
        notes: stories.map((story, i) => {
          const mark = (i + 1) + ')';
          story.dataset.hwpFootnoteLayout = JSON.stringify({ mark, above: 11, below: 8, between: 4, separator: 120, stroke: .5 });
          return { mark, fingerprint: fingerprint(story.textContent) };
        }),
      } };
      paragraph.dataset.hwpSourceWrap = JSON.stringify(wrap);
      const valid = !!parseHwpSourceWrap(paragraph.dataset.hwpSourceWrap);
      const naturalLines = paragraphLineOffsets(paragraph).length + 1;
      const priorHtml = paragraph.innerHTML;
      const physicalBottom = body.getBoundingClientRect().top + lineTop + 62;
      body.style.height = (lineTop + 62) + 'px';
      notes.style.marginTop = '2px';
      const bottom = sourceFootnoteEndBottom(paragraph, body, notes);
      const moved = paragraph.style.transform.startsWith('translateY(-');
      const unchanged = paragraph.innerHTML === priorHtml && refs.every(ref => ref.isConnected);
      clearSourceBodyEndPaint(paragraph);
      stories[1].textContent += ' edited';
      const rejected = sourceFootnoteEndBottom(paragraph, body, notes) === undefined && !paragraph.style.transform;
      return { valid, naturalLines, fitted: bottom !== undefined && bottom < physicalBottom, moved, unchanged, rejected };
    } finally { container.remove(); }
  });
  expect(result).toEqual({ valid: true, naturalLines: 2, fitted: true, moved: true, unchanged: true, rejected: true });
});

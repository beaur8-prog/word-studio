import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';

for (const nested of [false, true]) test(`${nested ? 'nested' : 'body'} merged-cell continuation paragraphs do not inherit normal body leading through two DOCX saves`, async ({ page }, info) => {
  await page.goto('/');
  const result = await page.evaluate(async nested => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const cell = (body: string, extra = '') => `<td ${extra} style="padding:0;border:1px solid black;vertical-align:middle">${body}</td>`;
    const small = '<p style="font-size:1pt;line-height:2px;margin:0"></p>';
    const rows = [
      `<tr style="height:12px">${cell(small, 'rowspan="2"')}${cell('<p style="font-size:14pt;line-height:19px;margin:0">MERGED TITLE</p>', 'rowspan="3"')}${cell(small)}</tr>`,
      `<tr style="height:2px">${cell(small, 'rowspan="2"')}</tr>`,
      `<tr style="height:12px">${cell(small)}</tr>`,
      `<tr style="height:400px">${cell('<p style="font-size:12pt;line-height:18px;margin:0">BODY 123 한글</p><p style="font-size:12pt;line-height:35px;margin:0"></p><p style="font-size:12pt;line-height:18px;margin:0">LAST</p>', 'colspan="3"')}</tr>`,
    ].join('');
    const table = `<table style="width:420px;margin:0"><colgroup><col style="width:110px"><col style="width:200px"><col style="width:110px"></colgroup>${rows}</table>`;
    const source = nested ? `<table style="width:430px;margin:0"><tr><td style="padding:0">${table}<p style="font-size:1pt;line-height:1px;margin:0"></p></td></tr></table>` : table;
    const loaded = await importFile(new File([source], 'merged.html'));
    let doc = { ...loaded.document, page: { ...loaded.document.page, size: 'Custom' as const, width: 160, height: 200, margin: 10, margins: { top: 10, bottom: 10, left: 10, right: 10 } }, content: parseDocumentHtml(loaded.html!) };
    const summarize = (root: any) => { const cells: any[] = [], rows: any[] = []; const walk = (n: any) => { if (['tableCell', 'tableHeader'].includes(n.type)) cells.push({ span: [n.attrs.rowspan, n.attrs.colspan], paragraphs: n.content.filter((p: any) => p.type === 'paragraph').map((p: any) => ({ text: (p.content || []).map((n: any) => n.type === 'hardBreak' ? '\n' : n.text || '').join(''), line: Math.round(parseFloat(p.attrs.lineSpacing) * (String(p.attrs.lineSpacing).endsWith('pt') ? 4 / 3 : 1) * 1000) / 1000 })) }); if (n.type === 'tableRow') rows.push(n.attrs.rowHeight); n.content?.forEach(walk); }; walk(root); return { cells, rows }; };
    const states = [], files = [];
    for (let save = 0; save < 3; save++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try { states.push({ ...summarize(doc.content), pages: layout.pages, text: layout.element.textContent }); } finally { layout.dispose(); }
      if (save === 2) break;
      const blob = await exportDocx(doc), bytes = new Uint8Array(await blob.arrayBuffer()), zip = await JSZip.loadAsync(bytes);
      const xml = await zip.file('word/document.xml')!.async('string');
      const root = new DOMParser().parseFromString(xml, 'application/xml'), W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      const continuations = [...root.getElementsByTagNameNS(W, 'tc')].filter(tc => [...tc.children].some(p => p.localName === 'tcPr' && [...p.children].some(v => v.localName === 'vMerge' && v.getAttributeNS(W, 'val') === 'continue'))).map(tc => {
        const paragraphs = [...tc.children].filter(p => p.localName === 'p');
        const p = paragraphs[0], spacing = p?.getElementsByTagNameNS(W, 'spacing')[0];
        return { ignoreCellMarker: !!tc.getElementsByTagNameNS(W, 'hideMark').length, paragraphs: paragraphs.length, text: p?.textContent, before: spacing?.getAttributeNS(W, 'before'), after: spacing?.getAttributeNS(W, 'after'), line: spacing?.getAttributeNS(W, 'line'), rule: spacing?.getAttributeNS(W, 'lineRule') };
      });
      files.push({ bytes: [...bytes], continuations });
      const reopened = await importFile(new File([blob], 'merged.docx')); doc = { ...reopened.document, content: parseDocumentHtml(reopened.html!) };
    }
    return { states, files };
  }, nested);
  await fs.writeFile(info.outputPath('merge-continuation-states.json'), JSON.stringify(result.states, null, 2));
  for (const [save, file] of result.files.entries()) {
    await fs.writeFile(info.outputPath(`textbox-save-merge-continuation-${nested}-${save + 1}.docx`), Buffer.from(file.bytes));
    expect(file.continuations).toHaveLength(4);
    for (const cell of file.continuations) expect(cell).toEqual({ ignoreCellMarker: true, paragraphs: 1, text: '', before: '0', after: '0', line: '1', rule: 'exact' });
  }
  for (const state of result.states) {
    expect(state.pages).toBe(1);
    expect(state.cells).toEqual(result.states[0].cells);
    expect(state.rows).toEqual(result.states[0].rows);
    // The authored 35px empty paragraph remains a real part of the cell.
    expect(state.cells.flatMap(c => c.paragraphs).filter(p => p.line === 35)).toHaveLength(1);
    expect(state.text).toContain('BODY 123 한글'); expect(state.text).toContain('LAST');
  }
});

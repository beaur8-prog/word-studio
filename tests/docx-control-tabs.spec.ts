import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';

test('literal and styled tabs use standard Word controls in body, tables, links, comments and furniture through two saves', async ({ page }, info) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const marked = [{ type: 'bold' }, { type: 'underline' }, { type: 'link', attrs: { href: 'https://example.com/' } }];
    let doc: any = { format: 'word-studio', version: 1, title: 'Control tabs', updatedAt: new Date().toISOString(),
      page: { size: 'A4', orientation: 'portrait', margin: 25.4, header: '', headerHtml: '<p>H\t\tEND\t</p>', footer: 'F\t\tEND\t', pageNumbers: false },
      comments: [{ id: 'c1', quote: 'END\t', text: 'REVIEW\tTAB', author: 'QA', createdAt: new Date().toISOString(), resolved: false }],
      content: { type: 'doc', content: [
        { type: 'paragraph', content: [{ type: 'text', text: 'START' }, { type: 'text', text: '\t\t', marks: marked },
          { type: 'text', text: 'END\t', marks: [{ type: 'italic' }, { type: 'commentAnchor', attrs: { ids: 'c1' } }] },
          { type: 'hardBreak' }, { type: 'documentTab', marks: marked }, { type: 'text', text: 'NEXT' }] },
        { type: 'table', content: [{ type: 'tableRow', content: [{ type: 'tableCell', attrs: { colspan: 1, rowspan: 1, colwidth: [300] }, content: [
          { type: 'paragraph', content: [{ type: 'text', text: '\tCELL\t\tEND\t' }] }] }] }] }
      ] } };
    const logical = (n: any): string => n.type === 'text' ? n.text : n.type === 'documentTab' ? '\t' : n.type === 'hardBreak' ? '\n' : (n.content || []).map(logical).join('');
    const expected = logical(doc.content), rows: any[] = [], files: number[][] = [];
    for (let save = 1; save <= 2; save++) {
      const blob = await exportDocx(doc), bytes = [...new Uint8Array(await blob.arrayBuffer())], zip = await JSZip.loadAsync(await blob.arrayBuffer());
      const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', xml = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
      const literalTabs: string[] = [];
      for (const name of Object.keys(zip.files).filter(f => /^word\/.*\.xml$/.test(f))) {
        const root = new DOMParser().parseFromString(await zip.file(name)!.async('string'), 'application/xml');
        for (const text of root.getElementsByTagNameNS(W, 't')) if (text.textContent?.includes('\t')) literalTabs.push(name);
      }
      const linkedTabs = [...xml.getElementsByTagNameNS(W, 'hyperlink')].flatMap(link => [...link.getElementsByTagNameNS(W, 'tab')]);
      const loaded = await importFile(new File([blob], 'tabs.docx'));
      doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
      rows.push({ literalTabs, bodyTabControls: xml.getElementsByTagNameNS(W, 'tab').length, linkedTabs: linkedTabs.length,
        styledLinkedTabs: linkedTabs.every(tab => { const run = tab.parentElement!; return !!run.getElementsByTagNameNS(W, 'b').length && !!run.getElementsByTagNameNS(W, 'u').length; }),
        logical: logical(doc.content), comment: doc.comments[0]?.text, quote: doc.comments[0]?.quote, header: doc.page.header, footer: doc.page.footer });
      files.push(bytes);
    }
    return { expected, rows, files };
  });
  for (const row of result.rows) {
    expect(row.literalTabs).toEqual([]); expect(row.bodyTabControls).toBe(8);
    expect(row.linkedTabs).toBe(3); expect(row.styledLinkedTabs).toBe(true);
    expect(row.logical).toBe(result.expected); expect(row.comment).toBe('REVIEW\tTAB'); expect(row.quote).toBe('END\t');
    expect(row.header).toBe('H\t\tEND\t'); expect(row.footer).toBe('F\t\tEND\t');
  }
  for (const [i, file] of result.files.entries()) await writeFile(info.outputPath(`textbox-save-control-tabs-${i + 1}.docx`), Buffer.from(file));
});

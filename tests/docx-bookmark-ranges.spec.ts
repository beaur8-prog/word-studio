import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';

test('named destinations cover their following text without duplicating it, while original names and links survive two DOCX saves', async ({ page }, info) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const { defaultPage } = await import('/src/types.ts');
    const mark = (name: string) => ({ type: 'documentBookmark', attrs: { name } });
    const target = '<참고 8>지역특화(주력)산업';
    let doc: any = { format: 'word-studio', version: 1, title: 'Bookmark interoperability', updatedAt: new Date().toISOString(),
      page: defaultPage, comments: [], content: { type: 'doc', content: [
        { type: 'paragraph', content: [mark(target), { type: 'text', text: '지역주력산업', marks: [{ type: 'bold' }] }] },
        { type: 'paragraph', content: [{ type: 'text', text: '😀Before' }, mark('첫째 & Bookmark'), mark('Second'), { type: 'text', text: 'TARGET' }, mark('EndAnchor')] },
        { type: 'paragraph', content: [mark('ws_bookmark_0'), { type: 'text', text: 'Go', marks: [{ type: 'link', attrs: { href: '#' + target } }] }] },
        { type: 'table', content: [{ type: 'tableRow', content: [{ type: 'tableCell', attrs: { colspan: 1, rowspan: 1, colwidth: [300] }, content: [
          { type: 'paragraph', content: [mark('셀 & 책갈피'), { type: 'text', text: 'CELL' }] }] }] }] }
      ] } };
    const logical = (n: any): string => n.text || (n.content || []).map(logical).join('');
    const expected = logical(doc.content), rows = [], files = [];
    for (let save = 1; save <= 2; save++) {
      const blob = await exportDocx(doc), bytes = [...new Uint8Array(await blob.arrayBuffer())], zip = await JSZip.loadAsync(await blob.arrayBuffer());
      const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      const xml = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
      const starts = [...xml.getElementsByTagNameNS(W, 'bookmarkStart')], ends = [...xml.getElementsByTagNameNS(W, 'bookmarkEnd')];
      const ranges = starts.map(start => {
        const end = ends.find(e => e.getAttributeNS(W, 'id') === start.getAttributeNS(W, 'id'))!;
        const range = document.createRange(); range.setStartAfter(start); range.setEndBefore(end);
        return { name: start.getAttributeNS(W, 'name')!, text: range.toString() };
      });
      const anchors = [...xml.getElementsByTagNameNS(W, 'hyperlink')].map(e => e.getAttributeNS(W, 'anchor'));
      const loaded = await importFile(new File([blob], 'bookmarks.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
      const names: string[] = [], positions: string[] = []; let before = '';
      const walk = (n: any) => { if (n.type === 'documentBookmark') { names.push(n.attrs.name); positions.push(before); } if (n.type === 'text') before += n.text; (n.content || []).forEach(walk); }; walk(doc.content);
      rows.push({ ranges, anchors, names, positions, text: logical(doc.content) }); files.push(bytes);
    }
    return { expected, rows, files };
  });
  for (const row of result.rows) {
    expect(row.text).toBe(result.expected);
    expect(row.ranges).toHaveLength(6);
    for (const range of row.ranges) expect(range.name).toMatch(/^[\p{L}_][\p{L}\p{N}_]{0,39}$/u);
    expect(row.ranges.filter(r => r.name !== 'EndAnchor').map(r => r.text)).toEqual(['지역주력산업', 'TARGET', 'TARGET', 'Go', 'CELL']);
    expect(row.ranges.find(r => r.name === 'EndAnchor')?.text).toBe('');
    expect(row.anchors).toEqual([row.ranges[0].name]);
    expect(row.names).toEqual(['<참고 8>지역특화(주력)산업', '첫째 & Bookmark', 'Second', 'EndAnchor', 'ws_bookmark_0', '셀 & 책갈피']);
    expect(row.positions).toEqual(['', '지역주력산업😀Before', '지역주력산업😀Before', '지역주력산업😀BeforeTARGET', '지역주력산업😀BeforeTARGET', '지역주력산업😀BeforeTARGETGo']);
  }
  for (const [i, file] of result.files.entries()) await writeFile(info.outputPath(`textbox-save-bookmark-ranges-${i + 1}.docx`), Buffer.from(file));
});

test('collapsed destinations before media or controls never absorb pictures, breaks or tabs through two saves', async ({ page }, info) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const canvas = document.createElement('canvas'); canvas.width = canvas.height = 2;
    canvas.getContext('2d')!.fillRect(0, 0, 2, 2); const src = canvas.toDataURL();
    const bookmark = (name: string) => ({ type: 'documentBookmark', attrs: { name } });
    let doc: any = { format: 'word-studio', version: 1, title: 'Bookmark barriers', updatedAt: new Date().toISOString(), page: defaultPage, comments: [],
      content: { type: 'doc', content: [
        { type: 'paragraph', content: [bookmark('Picture'), { type: 'inlineImage', attrs: { src, width: 20, height: 20, alt: 'PIXEL' } }, { type: 'text', text: 'CAPTION' }] },
        { type: 'paragraph', content: [bookmark('Break'), { type: 'hardBreak' }, { type: 'text', text: 'AFTER' }] },
        { type: 'paragraph', content: [bookmark('Tab'), { type: 'documentTab' }, { type: 'text', text: 'END' }] },
        { type: 'paragraph', content: [bookmark('LinkedTab'), { type: 'text', text: 'LINK\tEND', marks: [{ type: 'link', attrs: { href: 'https://example.com/' } }] }] },
        { type: 'paragraph', content: [bookmark('Field'), { type: 'documentField', attrs: { kind: 'PAGE' } }, { type: 'text', text: 'PAGE_AFTER' }] },
        { type: 'paragraph', content: [bookmark('Text'), { type: 'text', text: 'SAFE' }] }
      ] } };
    const logical = (n: any): string => n.type === 'text' ? n.text : n.type === 'documentTab' ? '\t' : n.type === 'hardBreak' ? '\n' : (n.content || []).map(logical).join('');
    const expected = logical(doc.content), rows = [], files = [];
    for (let save = 1; save <= 2; save++) {
      const blob = await exportDocx(doc), zip = await JSZip.loadAsync(await blob.arrayBuffer()), W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      const xml = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
      const ends = [...xml.getElementsByTagNameNS(W, 'bookmarkEnd')];
      const ranges = [...xml.getElementsByTagNameNS(W, 'bookmarkStart')].map(start => {
        const end = ends.find(e => e.getAttributeNS(W, 'id') === start.getAttributeNS(W, 'id'))!, range = document.createRange();
        range.setStartAfter(start); range.setEndBefore(end); const fragment = range.cloneContents();
        return { name: start.getAttributeNS(W, 'name'), text: range.toString(), controls: [...fragment.querySelectorAll('*')].filter(e => ['drawing', 'br', 'tab', 'fldChar', 'instrText'].includes(e.localName)).length };
      });
      const loaded = await importFile(new File([blob], 'barriers.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
      const pictures: string[] = []; const walk = (n: any) => { if (['image', 'inlineImage'].includes(n.type)) pictures.push(n.attrs.src); (n.content || []).forEach(walk); }; walk(doc.content);
      rows.push({ ranges, body: logical(doc.content), pictures }); files.push([...new Uint8Array(await blob.arrayBuffer())]);
    }
    return { expected, src, rows, files };
  });
  for (const row of result.rows) {
    expect(row.body).toBe(result.expected); expect(row.pictures).toEqual([result.src]);
    expect(row.ranges.slice(0, 3).map(r => r.text)).toEqual(['', '', '']);
    // Reimport splits a linked literal tab into separate text/tab runs. Its
    // pure LINK prefix may become a range, while the tab must stay outside.
    expect(['', 'LINK']).toContain(row.ranges[3].text);
    expect(row.ranges[4].text).toBe(''); expect(row.ranges[5].text).toBe('SAFE');
    expect(row.ranges.every(r => r.controls === 0)).toBe(true);
  }
  for (const [i, file] of result.files.entries()) await writeFile(info.outputPath(`textbox-save-bookmark-media-${i + 1}.docx`), Buffer.from(file));
});

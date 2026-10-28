import { test, expect } from '@playwright/test';

test('long HWP bookmark names become valid DOCX targets and retain their names through two saves', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const names = ['[별지3] 금융정보 등 제공동의서 공공주택특별법 시행규칙 [별지 제4호서식] 개정2017 앞면', 'A'.repeat(48), 'A'.repeat(47) + 'B', 'ws_bookmark_0'];
    let doc = { format: 'word-studio', version: 1, id: 'bookmarks', title: 'Bookmarks', comments: [], updatedAt: new Date().toISOString(), page: defaultPage,
      content: { type: 'doc', content: names.map(name => ({ type: 'paragraph', content: [{ type: 'documentBookmark', attrs: { name } }, { type: 'text', text: 'LINK', marks: [{ type: 'link', attrs: { href: '#' + name } }] }] })) } };
    const states = [];
    for (let i = 0; i < 2; i++) {
      const blob = await exportDocx(doc), zip = await JSZip.loadAsync(await blob.arrayBuffer());
      const xml = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
      const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      const exportedNames = [...xml.getElementsByTagNameNS(W, 'bookmarkStart')].map(e => e.getAttributeNS(W, 'name'));
      const targets = [...xml.getElementsByTagNameNS(W, 'hyperlink')].map(e => e.getAttributeNS(W, 'anchor'));
      const imported = await importFile(new File([blob], 'again.docx')); doc = { ...imported.document, content: parseDocumentHtml(imported.html!) };
      const restoredNames: string[] = [], restoredTargets: string[] = [];
      const walk = n => { if (n.type === 'documentBookmark') restoredNames.push(n.attrs.name); for (const m of n.marks || []) if (m.type === 'link') restoredTargets.push(m.attrs.href); n.content?.forEach(walk); }; walk(doc.content);
      states.push({ exportedNames, targets, restoredNames, restoredTargets });
    }
    return { names, states };
  });
  for (const state of result.states) {
    expect(state.exportedNames.every(n => n && n.length <= 40)).toBe(true);
    expect(new Set(state.exportedNames).size).toBe(result.names.length);
    expect(state.targets).toEqual(state.exportedNames);
    expect(state.restoredNames).toEqual(result.names); expect(state.restoredTargets).toEqual(result.names.map(n => '#' + n));
  }
});

test('external bookmark renames and removed optional name metadata keep their standard targets', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const name = 'BOOKMARK'.repeat(12), W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const doc = { format: 'word-studio', version: 1, id: 'external-bookmark', title: 'External', comments: [], updatedAt: new Date().toISOString(), page: defaultPage,
      content: { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'documentBookmark', attrs: { name } }, { type: 'text', text: 'LINK', marks: [{ type: 'link', attrs: { href: '#' + name } }] }] }] } };
    const blob = await exportDocx(doc), results = [];
    for (const change of ['remove', 'rename']) {
      const zip = await JSZip.loadAsync(await blob.arrayBuffer());
      const xml = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
      const bookmark = xml.getElementsByTagNameNS(W, 'bookmarkStart')[0];
      const target = change === 'rename' ? 'EXTERNALLY_RENAMED' : bookmark.getAttributeNS(W, 'name');
      if (change === 'rename') { bookmark.setAttributeNS(W, 'w:name', target!); xml.getElementsByTagNameNS(W, 'hyperlink')[0].setAttributeNS(W, 'w:anchor', target!); }
      else {
        const props = new DOMParser().parseFromString(await zip.file('docProps/custom.xml')!.async('string'), 'application/xml');
        [...props.getElementsByTagName('*')].find(p => p.getAttribute('name') === 'WordStudio.BookmarkNames')!.remove();
        zip.file('docProps/custom.xml', new XMLSerializer().serializeToString(props));
      }
      zip.file('word/document.xml', new XMLSerializer().serializeToString(xml));
      const loaded = await importFile(new File([await zip.generateAsync({ type: 'blob' })], 'external.docx'));
      const content = parseDocumentHtml(loaded.html!).content[0].content;
      results.push({ target, bookmark: content.find(n => n.type === 'documentBookmark').attrs.name, link: content.find(n => n.type === 'text').marks.find(m => m.type === 'link').attrs.href });
    }
    return results;
  });
  for (const state of result) { expect(state.bookmark).toBe(state.target); expect(state.link).toBe('#' + state.target); }
});

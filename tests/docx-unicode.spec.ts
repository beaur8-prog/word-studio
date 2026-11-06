import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';

test('supplementary Unicode survives DOCX XML chunk boundaries and two saves', async ({ page }, info) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const text = '\u{f0854}'.repeat(9000) + '천원' + '\u{f0855}'.repeat(9000) + 'x' + '\u{1f642}'.repeat(9000);
    const source = { format: 'word-studio', version: 1, title: 'Unicode 경계',
      content: { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text }] }] },
      page: { size: 'A4', orientation: 'portrait', margin: 20, header: '', footer: '', pageNumbers: false },
      comments: [{ id: 'unicode-comment', quote: '', text: '메모 \u{f0854}', createdAt: '2026-11-06T00:00:00.000Z', resolved: false }], updatedAt: '2026-11-06T00:00:00.000Z' };
    let current: any = source;
    const states = [];
    for (let save = 1; save <= 2; save++) {
      const blob = await exportDocx(current), bytes = new Uint8Array(await blob.arrayBuffer());
      const zip = await JSZip.loadAsync(bytes), invalid = [];
      for (const part of Object.values(zip.files)) if (!part.dir && /\.(xml|rels)$/.test(part.name)) {
        try { new TextDecoder('utf-8', { fatal: true }).decode(await part.async('uint8array')); }
        catch { invalid.push(part.name); }
      }
      const types = new DOMParser().parseFromString(await zip.file('[Content_Types].xml')!.async('string'), 'application/xml');
      const commentsType = [...types.getElementsByTagName('Override')].find(e => e.getAttribute('PartName') === '/word/commentsExtended.xml')?.getAttribute('ContentType');
      const xml = await zip.file('word/document.xml')!.async('string');
      const body = new DOMParser().parseFromString(xml, 'application/xml');
      const actual = [...body.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 't')].map(t => t.textContent).join('');
      const loaded = await importFile(new File([bytes], 'unicode.docx'));
      current = { ...loaded.document, content: loaded.html === undefined ? loaded.document.content : parseDocumentHtml(loaded.html) };
      const logical = (n: any): string => (n.text || '') + (n.content || []).map(logical).join('');
      states.push({ save, invalid, commentsType, textEqual: actual === text, importedEqual: logical(current.content) === text, bytes: [...bytes] });
    }
    return states;
  });
  for (const state of result) {
    await writeFile(info.outputPath(`hwp-gov-unicode-save${state.save}.docx`), Buffer.from(state.bytes));
    expect(state.invalid, `save ${state.save}: all XML parts must be strict UTF-8`).toEqual([]);
    expect(state.commentsType).toBe('application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml');
    expect(state.textEqual).toBe(true);
    expect(state.importedEqual).toBe(true);
  }
});

test('DOCX metadata preparation keeps supplementary scalars as strict UTF-8', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const { prepareDocxMetadata } = await import('/src/docx-metadata.ts');
    const text = '\u{f0854}'.repeat(9000) + 'x' + '\u{1f642}'.repeat(9000);
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
    const zip = new JSZip(); zip.file('word/document.xml', new TextEncoder().encode(xml));
    const body = new DOMParser().parseFromString(xml, 'application/xml');
    const result = await prepareDocxMetadata(zip, body, [{ size: 'A4', orientation: 'portrait', margin: 20, header: '', footer: '', pageNumbers: false }]);
    const prepared = await JSZip.loadAsync(result.buffer);
    try {
      const xml = new TextDecoder('utf-8', { fatal: true }).decode(await prepared.file('word/document.xml')!.async('uint8array'));
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      return { utf8: true, equal: [...doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 't')].map(t => t.textContent).join('') === text };
    } catch { return { utf8: false, equal: false }; }
  });
  expect(result).toEqual({ utf8: true, equal: true });
});

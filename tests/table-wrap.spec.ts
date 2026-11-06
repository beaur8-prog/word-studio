import { test, expect } from '@playwright/test';

for (const edit of ['position', 'width', 'remove-private-metadata', 'different-anchor'])
test(`standard narrow DOCX floating table respects external ${edit}`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async edit => {
    const { exportDocx, importFile } = await import('/src/document-io.ts'), { parseDocumentHtml } = await import('/src/extensions.ts');
    const { defaultPage } = await import('/src/types.ts'), { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const anchor = JSON.stringify({ x: 12, y: 0, bottom: 0, reference: 'paragraph', horizontal: 'margin', wrap: 'right', right: 6 }).replaceAll('"', '&quot;');
    const html = `<table data-table-anchor="${anchor}" style="margin-bottom:0"><colgroup><col style="width:160px"></colgroup><tr><td colwidth="160" style="padding:0;width:160px"><p>LABEL</p></td></tr></table><p style="line-height:24px;margin:0">Editable text beside the label and below it. ${'WORD TABLE WRAP '.repeat(20)}</p>`;
    const loaded = await importFile(new File([html], 'wrap.html'));
    const doc = { ...loaded.document, page: { ...defaultPage, pageNumbers: false }, content: parseDocumentHtml(html) };
    if (doc.content.content![0].content![0].content![0].attrs?.colwidth?.[0] !== 160) throw Error('the fixture must retain its narrow column width');
    const blob = await exportDocx(doc), zip = await JSZip.loadAsync(await blob.arrayBuffer());
    const before = await zip.file('word/document.xml')!.async('string'), xml = new DOMParser().parseFromString(before, 'application/xml');
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', tbl = xml.getElementsByTagNameNS(W, 'tbl')[0], floating = tbl.getElementsByTagNameNS(W, 'tblpPr')[0];
    if (!floating) throw Error('standard OOXML tblpPr is required');
    const original = { x: floating.getAttributeNS(W, 'tblpX'), y: floating.getAttributeNS(W, 'tblpY'), vert: floating.getAttributeNS(W, 'vertAnchor'), horz: floating.getAttributeNS(W, 'horzAnchor'), right: floating.getAttributeNS(W, 'rightFromText') };
    if (edit === 'position') { floating.setAttributeNS(W, 'w:tblpX', '360'); floating.setAttributeNS(W, 'w:rightFromText', '180'); }
    if (edit === 'width') { for (const name of ['tblW','tcW','gridCol']) for (const el of [...tbl.getElementsByTagNameNS(W, name)]) el.setAttributeNS(W, 'w:w', '6000'); }
    if (edit === 'different-anchor') floating.setAttributeNS(W, 'w:vertAnchor', 'page');
    if (edit === 'remove-private-metadata') for (const a of [...tbl.attributes]) if (a.namespaceURI?.startsWith('urn:wordstudio:')) tbl.removeAttributeNS(a.namespaceURI, a.localName);
    zip.file('word/document.xml', new XMLSerializer().serializeToString(xml));
    let round = await importFile(new File([await zip.generateAsync({ type: 'blob' })], 'external.docx'));
    const imported = { ...round.document, content: parseDocumentHtml(round.html!) }, table = imported.content.content!.find(n => n.type === 'table')!;
    const next = await importFile(new File([await exportDocx(imported)], 'again.docx'));
    return { original, anchor: table.attrs?.tableAnchor ? JSON.parse(table.attrs.tableAnchor) : null, width: table.content![0].content![0].attrs?.colwidth?.[0], second: parseDocumentHtml(next.html!).content!.find(n => n.type === 'table')!.attrs?.tableAnchor,
      text: (function t(n: any): string { return (n.text || '') + (n.content || []).map(t).join(''); })(imported.content) };
  }, edit);
  expect(result.original).toEqual({ x: '180', y: '0', vert: 'text', horz: 'margin', right: '90' });
  expect(result.text).toContain('LABEL'); expect(result.text).toContain('WORD TABLE WRAP');
  if (edit === 'different-anchor') { expect(result.anchor).toBe(null); expect(result.second).toBeFalsy(); }
  else {
    expect(result.anchor).toMatchObject({ x: edit === 'position' ? 24 : 12, y: 0, right: edit === 'position' ? 12 : 6, wrap: 'right' });
    expect(JSON.parse(result.second)).toEqual(result.anchor); expect(result.width).toBe(edit === 'width' ? 400 : 160);
  }
});

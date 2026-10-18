import { test, expect } from '@playwright/test';

test('DOCX header/footer inheritance follows each type and explicit blank overrides through multiple sections', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    // @ts-expect-error Vite source module.
    const io = await import('/src/document-io.ts');
    // @ts-expect-error Vite dependency import.
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    // @ts-expect-error Vite source module.
    const { editorExtensions } = await import('/src/extensions.ts');
    // @ts-expect-error Vite dependency import.
    const { generateJSON } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const base = { size: 'A4', orientation: 'portrait', margin: 25.4, header: 'Shared Header', footer: 'Shared Footer', pageNumbers: true };
    const settings = [base, { ...base, size: 'Letter', orientation: 'landscape' }, { ...base, header: '' },
      { ...base, header: '', footer: '', pageNumbers: false }, { ...base, header: '', footer: '', pageNumbers: false }];
    const content: any[] = [];
    settings.forEach((p, i) => {
      if (i) content.push({ type: 'sectionBreak', attrs: { page: JSON.stringify(p) } });
      content.push({ type: 'paragraph', content: [{ type: 'text', text: `SECTION ${i}` }] });
    });
    const doc = { format: 'word-studio', version: 1, id: 'linked', title: 'Linked', updatedAt: new Date().toISOString(), page: base, comments: [], content: { type: 'doc', content } };
    const zip = await JSZip.loadAsync(await (await io.exportDocx(doc)).arrayBuffer());
    const xml = new DOMParser().parseFromString(await zip.file('word/document.xml').async('string'), 'application/xml');
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const sections = [...xml.getElementsByTagNameNS(W, 'sectPr')];
    sections.forEach((section, index) => {
      const remove = index === 1 || index === 4 ? ['headerReference', 'footerReference']
        : index === 2 ? ['footerReference'] : index === 3 ? ['headerReference'] : [];
      [...section.children].filter(el => remove.includes(el.localName)).forEach(el => el.remove());
    });
    zip.file('word/document.xml', new XMLSerializer().serializeToString(xml));
    const loaded = await io.importFile(new File([await zip.generateAsync({ type: 'blob' })], 'linked.docx'));
    const getPages = (loaded: any) => {
      const template = document.createElement('template'); template.innerHTML = loaded.html;
      return [loaded.document.page, ...[...template.content.querySelectorAll('[data-section-page]')].map(el => JSON.parse(el.getAttribute('data-section-page')!))];
    };
    const again = await io.importFile(new File([await io.exportDocx({ ...loaded.document, content: generateJSON(loaded.html, editorExtensions) })], 'again.docx'));
    return { pages: getPages(loaded), roundtrip: getPages(again) };
  });
  const expected = [
    { header: 'Shared Header', footer: 'Shared Footer', pageNumbers: true },
    { header: 'Shared Header', footer: 'Shared Footer', pageNumbers: true, size: 'Letter', orientation: 'landscape' },
    { header: '', footer: 'Shared Footer', pageNumbers: true },
    { header: '', footer: '', pageNumbers: false },
    { header: '', footer: '', pageNumbers: false },
  ];
  expect(result.pages).toMatchObject(expected);
  expect(result.roundtrip).toMatchObject(expected);
});

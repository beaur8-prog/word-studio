import { expect, test } from '@playwright/test';
import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function spacingDocument() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  </Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="document" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  </Relationships>`);
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  </Relationships>`);
  zip.file('word/styles.xml', `<w:styles xmlns:w="${W}">
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
    <w:style w:type="paragraph" w:styleId="Tracked"><w:name w:val="Tracked QA"/>
      <w:basedOn w:val="Normal"/><w:rPr><w:spacing w:val="-40"/></w:rPr></w:style>
  </w:styles>`);
  zip.file('word/document.xml', `<w:document xmlns:w="${W}"><w:body>
    <w:p><w:pPr><w:pStyle w:val="Tracked"/></w:pPr>
      <w:r><w:rPr><w:b/><w:color w:val="234567"/><w:spacing w:val="0"/><w:sz w:val="24"/></w:rPr>
        <w:t>Zero resets tracking</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Tracked"/></w:pPr><w:r><w:t>Inherited tracking</w:t></w:r></w:p>
    <w:p><w:r><w:t>Unspecified tracking</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Tracked"/></w:pPr>
      <w:r><w:rPr><w:spacing w:val="0"/></w:rPr><w:t>Left</w:t><w:tab/><w:t>Right</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body></w:document>`);
  return [...await zip.generateAsync({ type: 'uint8array' })];
}

test('DOCX saves keep explicit zero tracking distinct from inherited or unspecified spacing', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, ns }) => {
    const { unzipSync, zipSync, strFromU8, strToU8 } = await import('/node_modules/fflate/esm/browser.js');
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const parse = (xml: string) => new DOMParser().parseFromString(xml, 'application/xml');
    const find = (element: Element | Document | undefined | null, name: string) =>
      element?.getElementsByTagNameNS(ns, name)[0];
    const open = (data: Blob | Uint8Array<ArrayBuffer>) =>
      importFile(new File([data], 'character-spacing.docx'));
    const content = (file: Awaited<ReturnType<typeof open>>) => parseDocumentHtml(file.html!);
    const save = (file: Awaited<ReturnType<typeof open>>) => exportDocx({ ...file.document, content: content(file) });

    const source = await open(new Uint8Array(bytes));
    const saved1 = await save(source), reopened1 = await open(saved1);
    const saved2 = await save(reopened1), reopened2 = await open(saved2);
    const xmlRuns = async (file: Blob) => {
      const archive = unzipSync(new Uint8Array(await file.arrayBuffer()));
      const tree = parse(strFromU8(archive['word/document.xml']));
      return [...tree.getElementsByTagNameNS(ns, 'r')].filter(run => find(run, 't') || find(run, 'tab')).map(run => ({
        text: [...run.getElementsByTagNameNS(ns, 't')].map(t => t.textContent).join(''),
        tab: !!find(run, 'tab'),
        tracking: find(find(run, 'rPr'), 'spacing')?.getAttributeNS(ns, 'val') ?? null,
      }));
    };
    const modelRuns = (file: Awaited<ReturnType<typeof open>>) => content(file).content!.flatMap(paragraph =>
      (paragraph.content ?? []).filter(run => run.type === 'text' || run.type === 'documentTab').map(run => ({
        text: run.text ?? '',
        tab: run.type === 'documentTab',
        tracking: run.marks?.find(mark => mark.type === 'textStyle')?.attrs?.letterSpacing ?? null,
      })));

    // Model changing the exported paragraph style in an external editor.
    const changed = unzipSync(new Uint8Array(await saved2.arrayBuffer()));
    const exportedDocument = parse(strFromU8(changed['word/document.xml']));
    const styleId = find(find(exportedDocument, 'p'), 'pStyle')!.getAttributeNS(ns, 'val');
    const styles = parse(strFromU8(changed['word/styles.xml']));
    const tracked = [...styles.getElementsByTagNameNS(ns, 'style')].find(style => style.getAttributeNS(ns, 'styleId') === styleId)!;
    let runProperties = find(tracked, 'rPr');
    if (!runProperties) {
      runProperties = styles.createElementNS(ns, 'w:rPr');
      tracked.append(runProperties);
    }
    const spacing = styles.createElementNS(ns, 'w:spacing');
    spacing.setAttributeNS(ns, 'w:val', '-40');
    runProperties.append(spacing);
    changed['word/styles.xml'] = strToU8(new XMLSerializer().serializeToString(styles));
    const external = await open(zipSync(changed));
    const firstPrintedWidth = async (file: Awaited<ReturnType<typeof open>>) => {
      const layout = await createPrintLayout(generateHTML(content(file), editorExtensions), file.document.page);
      try {
        const first = layout.element.querySelector('.print-content p')!;
        const range = document.createRange();
        range.selectNodeContents(first);
        return range.getBoundingClientRect().width;
      } finally { layout.dispose(); }
    };
    return {
      model: [source, reopened1, reopened2, external].map(modelRuns),
      savedRuns: await Promise.all([saved1, saved2].map(xmlRuns)),
      widths: [await firstPrintedWidth(source), await firstPrintedWidth(external)],
    };
  }, { bytes: await spacingDocument(), ns: W });

  expect(result.savedRuns[0].map(run => run.tracking)).toEqual(['0', '-40', null, '0', '0', '0']);
  expect(result.savedRuns[1]).toEqual(result.savedRuns[0]);
  expect(result.savedRuns[1].filter(run => run.tab)).toHaveLength(1);
  expect(result.model[0].map(run => run.tracking)).toEqual(['0pt', '-2pt', null, '0pt', '0pt', '0pt']);
  for (const roundTrip of result.model.slice(1)) expect(roundTrip).toEqual(result.model[0]);
  expect(result.widths[1]).toBeCloseTo(result.widths[0], 1);
});

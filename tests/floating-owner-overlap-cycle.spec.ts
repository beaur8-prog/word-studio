import { readFile } from 'node:fs/promises';
import { expect, test, type Locator, type Page } from '@playwright/test';
import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const packageRelationship = 'http://schemas.openxmlformats.org/package/2006/relationships';
const drawing = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const pic = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const head = 'QA FIRST DRAWING HEADING';
const tail = 'QA AFTER THE LAST PAGE BREAK';
const ids = ['QA RED RECTANGLE', 'QA BLUE RECTANGLE'];
// Self-authored, solid-color 68 × 86 and 68 × 73 PNGs.
const pixels = [
  'iVBORw0KGgoAAAANSUhEUgAAAEQAAABWCAIAAAD5b6d1AAAAfUlEQVR4nO3PAQkAIADAMCPZwe7mMoZXBg/wjT3XN43rBzAwbwVTDaYaTDWYajDVYKrBVIOpBlMNphpMNZhqMNVgqsFUg6kGUw2mGkw1mGow1WCqwVSDqQZTDaYaTDWYajDVYKrBVIOpBlMNphpMNZhqMNVgqsFUg6n2FeYAv7/vrqlGY2MAAAAASUVORK5CYII=',
  'iVBORw0KGgoAAAANSUhEUgAAAEQAAABJCAIAAAAL7xc7AAAAbklEQVR4nO3PAQkAIADAMLPYzMwWMoZXBgvwj7n2N8b1AjNm3mKmykyVmSozVWaqzFSZqTJTZabKTJWZKjNVZqrMVJmpMlNlpspMlZkqM1VmqsxUmakyU2WmykyVmSozVWaqzFSZqTJTZabKTNUB1afpECr5U08AAAAASUVORK5CYII=',
];
type Box = [x: number, y: number, width: number, height: number];
const overlapping: [Box, Box] = [[720090, 1654493, 6479858, 8223885], [720090, 1275398, 6479858, 6983730]];
const separate = [[96, 176, 160, 260], [416, 176, 160, 260]]
  .map(box => box.map(value => value * 9525) as Box) as [Box, Box];

function picture(index: number, [x, y, width, height]: Box) {
  const label = ids[index];
  return `<w:r><w:drawing>
    <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${height}"
      behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/>
      <wp:positionH relativeFrom="page"><wp:posOffset>${x}</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="page"><wp:posOffset>${y}</wp:posOffset></wp:positionV>
      <wp:extent cx="${width}" cy="${height}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>
      <wp:wrapSquare wrapText="bothSides" distT="0" distB="0" distL="0" distR="0"/>
      <wp:docPr id="${index + 1}" name="${label}" descr="${label}"/>
      <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
      <a:graphic><a:graphicData uri="${pic}"><pic:pic>
        <pic:nvPicPr><pic:cNvPr id="${index + 1}" name="figure-${index}.png"/><pic:cNvPicPr/></pic:nvPicPr>
        <pic:blipFill><a:blip r:embed="rId${index + 2}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
        <pic:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </pic:spPr>
      </pic:pic></a:graphicData></a:graphic>
    </wp:anchor></w:drawing></w:r>`;
}

function paragraph(content = '', line = 384, size = 24) {
  return `<w:p><w:pPr>
    <w:keepNext w:val="0"/><w:keepLines w:val="0"/><w:widowControl w:val="0"/>
    <w:spacing w:before="0" w:after="0" w:line="${line}" w:lineRule="exact"/>
    <w:jc w:val="both"/><w:rPr><w:sz w:val="${size}"/></w:rPr>
  </w:pPr>${content}</w:p>`;
}

async function fixture(middleBreak: boolean, boxes: [Box, Box]) {
  const zip = new JSZip();
  const breakParagraph = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  const textRun = (value: string) => `<w:r><w:t>${value}</w:t></w:r>`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Default Extension="png" ContentType="image/png"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  </Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="${packageRelationship}">
    <Relationship Id="rId1" Type="${relationship}/officeDocument" Target="word/document.xml"/>
    <Relationship Id="rId2" Type="${packageRelationship}/metadata/core-properties" Target="docProps/core.xml"/>
  </Relationships>`);
  zip.file('docProps/core.xml', `<cp:coreProperties
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>Independent Layout QA</dc:creator></cp:coreProperties>`);
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="${packageRelationship}">
    <Relationship Id="rId1" Type="${relationship}/styles" Target="styles.xml"/>
    <Relationship Id="rId2" Type="${relationship}/image" Target="media/figure-0.png"/>
    <Relationship Id="rId3" Type="${relationship}/image" Target="media/figure-1.png"/>
  </Relationships>`);
  zip.file('word/styles.xml', `<w:styles xmlns:w="${w}">
    <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>
    </w:rPr></w:rPrDefault></w:docDefaults>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  </w:styles>`);
  zip.file('word/document.xml', `<w:document xmlns:w="${w}" xmlns:wp="${wp}" xmlns:r="${relationship}"
    xmlns:a="${drawing}" xmlns:pic="${pic}"><w:body>
    ${paragraph('', 352, 22)}${paragraph(textRun(head), 352, 22)}${paragraph(picture(0, boxes[0]), 15, 22)}
    ${middleBreak ? breakParagraph : ''}${paragraph()}${paragraph(picture(1, boxes[1]), 15)}${paragraph()}
    ${breakParagraph}${paragraph()}${paragraph(textRun(tail))}
    <w:sectPr><w:pgSz w:w="11906" w:h="16838" w:orient="portrait"/>
    <w:pgMar w:top="1417" w:right="1134" w:bottom="1134" w:left="1134" w:header="850" w:footer="567" w:gutter="0"/></w:sectPr>
  </w:body></w:document>`);
  for (const [index, image] of pixels.entries()) zip.file(`word/media/figure-${index}.png`, image, { base64: true });
  return zip.generateAsync({ type: 'nodebuffer' });
}

function snapshot(root: Locator, print: boolean) {
  return root.evaluate(async (scope, { print, head, tail }) => {
    const { documentScale } = await import('/src/document-scale.ts');
    const stack = scope.closest<HTMLElement>('.document-stack');
    const printPages = [...scope.querySelectorAll('.print-page')];
    const scale = documentScale(scope as HTMLElement);
    const paper = Number(stack?.dataset.pageHeight);
    const stride = paper + 24;
    const rootTop = scope.getBoundingClientRect().top;
    const pageOf = (el: Element) => print
      ? Number(el.closest<HTMLElement>('.print-page')?.dataset.page)
      : Math.max(1, 1 + Math.floor(((el.getBoundingClientRect().top - rootTop) / scale + .5) / stride));
    const images = [...scope.querySelectorAll<HTMLImageElement>('img[data-image-floating]')];
    const markers = [...scope.querySelectorAll('p')].filter(p => p.textContent === head || p.textContent === tail);
    return {
      pages: print ? printPages.length : Number(stack?.dataset.pages),
      images: images.map(img => {
        const box = img.getBoundingClientRect();
        const sheet = img.closest('.print-page')?.getBoundingClientRect();
        const onPrintPaper = !sheet || (box.top >= sheet.top - .5 && box.bottom <= sheet.bottom + .5);
        return {
          label: img.alt, paintedPage: pageOf(img), anchorPage: pageOf(img.closest('p')!),
          drawn: box.height > 0 && box.width > 0 && onPrintPaper,
        };
      }),
      markers: markers.map(p => ({ text: p.textContent, page: pageOf(p) })),
    };
  }, { print, head, tail });
}

async function sourceStructure(page: Page, bytes: Buffer) {
  const zip = await JSZip.loadAsync(bytes);
  const xml = await zip.file('word/document.xml')!.async('string');
  return page.evaluate(({ xml, w, wp }) => {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const children = (el: Document | Element, ns: string, name: string) => [...el.getElementsByTagNameNS(ns, name)];
    return {
      text: children(doc, w, 't').map(el => el.textContent).filter(Boolean),
      breaks: children(doc, w, 'br').filter(el => el.getAttributeNS(w, 'type') === 'page').length,
      anchors: children(doc, wp, 'anchor').map(el => {
        const horizontal = children(el, wp, 'positionH')[0];
        const vertical = children(el, wp, 'positionV')[0];
        return {
          reference: [horizontal.getAttribute('relativeFrom'), vertical.getAttribute('relativeFrom')],
          offset: [horizontal, vertical].map(position => Number(children(position, wp, 'posOffset')[0].textContent)),
          size: ['cx', 'cy'].map(axis => Number(children(el, wp, 'extent')[0].getAttribute(axis))),
          wrap: children(el, wp, 'wrapSquare')[0]?.getAttribute('wrapText'),
        };
      }),
    };
  }, { xml, w, wp });
}

test('two floating pictures keep their settled editor and print papers across an authored break or natural flow', async ({ page }) => {
  test.setTimeout(90000);
  const cases = [
    { name: 'overlap-with-authored-break', middleBreak: true, boxes: overlapping, pages: 3, figurePages: [1, 2] },
    { name: 'overlap-without-authored-break', middleBreak: false, boxes: overlapping, pages: 3, figurePages: [1, 2] },
    { name: 'nonoverlap-on-one-paper', middleBreak: false, boxes: separate, pages: 2, figurePages: [1, 1] },
  ];
  for (const scenario of cases) await test.step(scenario.name, async () => {
    await page.goto('/');
    const bytes = await fixture(scenario.middleBreak, scenario.boxes);
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
    await page.locator('input[type=file]').first().setInputFiles({
      name: `${scenario.name}.docx`, buffer: bytes,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(scenario.name);
    await expect(page.locator('.busy-indicator')).toBeHidden();
    await page.evaluate(() => document.fonts.ready);
    const expected = {
      pages: scenario.pages,
      images: ids.map((label, i) => ({ label, paintedPage: scenario.figurePages[i], anchorPage: scenario.figurePages[i], drawn: true })),
      markers: [{ text: head, page: 1 }, { text: tail, page: scenario.pages }],
    };
    await expect.poll(() => snapshot(editor, false), { timeout: 12000 }).toEqual(expected);
    // A transient matching page count must not pass a repeating layout cycle.
    for (let sample = 0; sample < 4; sample++) {
      await page.waitForTimeout(120);
      expect(await snapshot(editor, false)).toEqual(expected);
    }
    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const print = page.locator('.print-layout');
    await expect.poll(() => snapshot(print, true)).toEqual(expected);

    if (scenario.name !== 'overlap-without-authored-break') return;
    await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
    await page.getByRole('button', { name: '파일', exact: true }).click();
    const exported = page.waitForEvent('download');
    await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
    const source = await sourceStructure(page, bytes);
    expect(source).toEqual({
      text: [head, tail], breaks: 1,
      anchors: scenario.boxes.map(([x, y, width, height]) => ({
        reference: ['page', 'page'], offset: [x, y], size: [width, height], wrap: 'bothSides',
      })),
    });
    expect(await sourceStructure(page, await readFile(await (await exported).path()))).toEqual(source);
  });
});

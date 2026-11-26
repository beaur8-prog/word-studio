import { expect, test } from '@playwright/test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import fs from 'node:fs/promises';
import { createRequire } from 'node:module';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const req = createRequire(import.meta.url);
const confirmed = ['CHECKED FIRST', 'CHECKED NEXT'];
const unconfirmed = ['UNVERIFIED FIRST', 'UNVERIFIED NEXT'];

async function source(format: 'hwp' | 'hwpx') {
  const core = path.dirname(req.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  const doc = HwpDocument.createEmpty();
  try {
    doc.createBlankDocument();
    const paragraph = (text: string) => `<p style="font-size:10pt;line-height:130%;margin:0">${text}</p>`;
    const row = (key: string, content: string[]) => `<tr><td style="width:100px">${paragraph(key)}</td>` +
      `<td style="width:180px">${content.map(paragraph).join('')}</td></tr>`;
    doc.pasteHtml(0, 0, 0, `<table style="width:280px">${row('PROVEN', confirmed)}${row('CONTROL', unconfirmed)}</table>`);
    const table = JSON.parse(doc.getControls()).find(c => c.ctrlId === 'tbl');
    const address = [0, table.para, table.controlIndex] as const;
    doc.setTableProperties(...address, JSON.stringify({ pageBreak: 2, treatAsChar: false, textWrap: 'TopAndBottom' }));
    for (let cell = 0; cell < 4; cell++)
      doc.setCellProperties(...address, cell, JSON.stringify({ height: 1500, paddingTop: 141, paddingBottom: 141, applyInnerMargin: true }));
    for (const cell of [1, 3]) for (let paragraph = 0; paragraph < 2; paragraph++)
      doc.applyParaFormatInCell(...address, cell, paragraph, JSON.stringify({ spacingBefore: 600, spacingAfter: 0, lineSpacing: 130, lineSpacingType: 'Percent' }));
    doc.reflowLinesegs();

    // The author's layout includes the first before-space in the proven cell.
    // Keep the control's normal reflow without that first stored offset.
    const members = unzipSync(doc.exportHwpx());
    const section = 'Contents/section0.xml';
    const xml = strFromU8(members[section]);
    let adjusted = 0;
    members[section] = strToU8(xml.replace(/<hp:t>CHECKED (?:FIRST|NEXT)<\/hp:t>[\s\S]*?<\/hp:p>/g, paragraph =>
      paragraph.replace(/vertpos="(\d+)"/g, (_, position) => {
        adjusted++;
        return `vertpos="${Number(position) + 300}"`;
      })));
    expect(adjusted).toBe(2);
    const hwpx = zipSync(members);
    const authored = new HwpDocument(hwpx);
    try {
      const bytes = format === 'hwp' ? authored.exportHwp() : hwpx;
      const reopened = new HwpDocument(bytes);
      try {
        const first = (cellIndex: number) => {
          const reference = JSON.stringify([{ controlIndex: address[2], cellIndex, cellParaIndex: 0 }]);
          return JSON.parse(reopened.getParagraphLineMetrics(0, address[1], reference)).lines[0].verticalPosition;
        };
        expect([first(1), first(3)]).toEqual([4, 0]);
        expect([1, 3].map(cell => JSON.parse(reopened.getCellParaPropertiesAt(...address, cell, 0)).spacingBefore)).toEqual([8, 8]);
        expect(JSON.parse(reopened.getTableCellBboxes(...address)).map(cell => cell.h)).toEqual([42.4, 42.4, 38.4, 38.4]);
        return [...bytes];
      } finally { reopened.free(); }
    } finally { authored.free(); }
  } finally { doc.free(); }
}

for (const format of ['hwp', 'hwpx'] as const)
  test(`${format}: an intact cell retains verified before-spacing in print and two DOCX saves`, async ({ page }, info) => {
    const bytes = await source(format);
    await page.goto('/');
    const result = await page.evaluate(async ({ bytes, format, confirmed, unconfirmed }) => {
      const { HwpReaderSession } = await import('/src/hwp/reader.ts');
      const { importEditableHwp } = await import('/src/hwp/editing.ts');
      const { exportDocx, importFile } = await import('/src/document-io.ts');
      const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
      const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
      const { createPrintLayout } = await import('/src/print-layout.ts');
      const { default: Zip } = await import('/node_modules/.vite/deps/jszip.js');
      const session = new HwpReaderSession();
      const filename = 'Cell spacing.' + format;
      const markers = [...confirmed, ...unconfirmed];
      const states = [], saves: number[][] = [];
      try {
        await session.open(new File([Uint8Array.from(bytes)], filename));
        let doc = importEditableHwp(await session.editable(), filename);
        for (let stage = 0; stage < 3; stage++) {
          const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
          try {
            const nodes = [...layout.element.querySelectorAll('td > p')];
            const paragraphs = markers.map(text => nodes.find(node => node.textContent === text)!);
            const row = paragraphs[0].closest('tr')!;
            states.push({ before: paragraphs.map(node => getComputedStyle(node).marginTop),
              rowHeight: row.getBoundingClientRect().height, cells: [...row.cells].map(cell => [cell.colSpan, cell.rowSpan]),
              text: paragraphs.map(node => node.textContent), pages: layout.pages });
          } finally { layout.dispose(); }
          if (stage === 2) break;
          const blob = await exportDocx(doc);
          const packed = await blob.arrayBuffer();
          saves.push([...new Uint8Array(packed)]);
          const zip = await Zip.loadAsync(packed);
          const markup = new DOMParser().parseFromString(await zip.file('word/document.xml').async('string'), 'application/xml');
          const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
          const xmlBefore = markers.map(text => [...markup.getElementsByTagNameNS(ns, 'p')]
            .find(node => node.textContent === text)?.getElementsByTagNameNS(ns, 'spacing')[0]?.getAttributeNS(ns, 'before'));
          if (xmlBefore.some((value, index) => index < confirmed.length ? value !== '60' : Number(value) !== 0))
            throw new Error('Unexpected standard before-spacing: ' + JSON.stringify(xmlBefore));
          const reopened = await importFile(new File([blob], 'Again.docx'));
          doc = { ...reopened.document, content: parseDocumentHtml(reopened.html!) };
        }
        return { states, saves };
      } finally { session.close(); }
    }, { bytes, format, confirmed, unconfirmed });

    for (const [i, bytes] of result.saves.entries())
      await fs.writeFile(info.outputPath(`textbox-save-cell-before-${format}-${i + 1}.docx`), Buffer.from(bytes));
    for (const state of result.states) {
      expect(state).toMatchObject({ before: ['4px', '4px', '0px', '0px'], cells: [[1, 1], [1, 1]], text: [...confirmed, ...unconfirmed], pages: 1 });
      expect(state.rowHeight).toBeCloseTo(42.4, 0);
    }
    if (format !== 'hwp') return;
    const input = page.locator('input[type=file]').first();
    await input.setInputFiles({ name: 'Cell spacing.hwp', mimeType: 'application/octet-stream', buffer: Buffer.from(bytes) });
    const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
    await reader.getByRole('button', { name: 'Word 화면에서 편집', exact: true }).click();
    await reader.getByRole('button', { name: '편집본 만들기', exact: true }).click();
    const paragraph = page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).locator('td p').filter({ hasText: confirmed[0] });
    await expect(paragraph).toHaveCSS('margin-top', '4px');
    await paragraph.click();
    await paragraph.evaluate(element => {
      const range = document.createRange();
      range.selectNodeContents(element);
      range.collapse(false);
      window.getSelection()!.removeAllRanges();
      window.getSelection()!.addRange(range);
    });
    await page.keyboard.insertText(' EDIT');
    await expect(paragraph).toHaveText(confirmed[0] + ' EDIT');
    await expect(paragraph).toHaveCSS('margin-top', '4px');
    await page.keyboard.press('ControlOrMeta+z');
    await expect(paragraph).toHaveText(confirmed[0]);
    await page.keyboard.press('ControlOrMeta+Shift+z');
    await expect(paragraph).toHaveText(confirmed[0] + ' EDIT');
    await page.getByRole('button', { name: '파일', exact: true }).click();
    const downloaded = page.waitForEvent('download');
    await page.getByRole('button', { name: 'Word 문서로 내보내기 DOCX', exact: true }).click();
    const edited = info.outputPath('textbox-save-cell-before-edited.docx');
    await (await downloaded).saveAs(edited);
    await input.setInputFiles(edited);
    await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(path.basename(edited, '.docx'));
    await expect(paragraph).toHaveText(confirmed[0] + ' EDIT');
    await expect(paragraph).toHaveCSS('margin-top', '4px');
  });

import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import { BorderStyle, Document, LineRuleType, Packer, Paragraph, Table, TableCell, TableRow, TextRun, WidthType } from 'docx';
import JSZip from 'jszip';

const namespace = 'urn:wordstudio:paragraph-table:1';
const margin = { top: 90, bottom: 90, left: 60, right: 60 };
const stroke = { style: BorderStyle.SINGLE, size: 6, color: '102030' } as const;
const doubleStroke = { ...stroke, style: BorderStyle.DOUBLE };
const modes = ['ordinary', 'centered', 'double', 'inline', 'merged'] as const;

async function fixture(mode: typeof modes[number]) {
  const p = (label: string, advance = 300) => new Paragraph({
    children: [new TextRun({ text: label, font: 'Arial', size: 18 })],
    spacing: { before: 0, after: 0, line: advance, lineRule: LineRuleType.EXACT },
  });
  const spans = mode === 'inline' || mode === 'merged';
  const rows = Array.from({ length: 8 }, (_, row) => new TableRow({
    tableHeader: row === 0,
    cantSplit: true,
    children: (spans && (row === 2 || row === 3) ? ['A'] : ['A', 'B']).map(column => {
      const merged = spans && row === 1 && column === 'B';
      return new TableCell({ margins: margin, verticalAlign: 'center', ...(merged ? { rowSpan: 3 } : {}),
        children: [p(column + row), ...(merged && mode === 'inline' ? [p('B1b')] : [])],
        borders: { top: mode === 'double' || mode === 'inline' && row === 1 ? doubleStroke : stroke,
          bottom: mode === 'inline' && row === 0 ? doubleStroke : stroke, left: stroke, right: stroke },
      });
    }),
  }));
  const table = new Table({ rows, layout: 'fixed', columnWidths: [2250, 2250], width: { size: 4500, type: WidthType.DXA } });
  const bytes = await Packer.toBuffer(new Document({ sections: [{ properties: { page: {
    size: { width: 6000, height: 6000 }, margin: { top: 750, bottom: 750, left: 750, right: 750 },
  } }, children: [p('LEAD', 650), table] }] }));
  if (mode === 'ordinary') return bytes;
  const zip = await JSZip.loadAsync(bytes), xml = await zip.file('word/document.xml')!.async('string');
  const escape = (value: string) => value.replaceAll('"', '&quot;');
  const anchor = escape(JSON.stringify({ x: 0, y: 0, bottom: 0, reference: 'paragraph', horizontal: 'margin' }));
  const attributes = mode === 'inline' ? 'mc:PreserveAttributes="ws:centeredBorders" ws:centeredBorders="inline"'
    : `mc:PreserveAttributes="ws:anchor ws:centeredBorders" ws:anchor="${anchor}" ws:centeredBorders="1"`;
  zip.file('word/document.xml', xml.replace('<w:tbl>', `<w:tbl xmlns:ws="${namespace}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="ws" ${attributes}>`));
  return zip.generateAsync({ type: 'nodebuffer' });
}

for (const mode of modes) {
  test(`HWP row border centers preserve painted strokes and authored insets: ${mode}`, async ({ page }, info) => {
    let bytes = await fixture(mode);
    await page.goto('/');
    const spans = mode === 'inline' || mode === 'merged';
    const active = mode === 'centered' || spans;
    for (let pass = 0; pass <= (active ? 2 : 0); pass++) {
      const name = 'hwp-bordered-rows-' + mode + pass;
      await page.locator('input[type=file]').first().setInputFiles({
        name: name + '.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: bytes,
      });
      await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(name);
      const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
      const editorTable = editor.locator('table');
      await expect(editorTable).toHaveCount(1);
      if (active) await expect(editorTable).toHaveAttribute('data-hwp-centered-borders-active', 'true');
      else await expect(editor.locator('table[data-hwp-centered-borders-active]')).toHaveCount(0);
      await expect(editor.locator('td').filter({ hasText: /^A7$/ })).toHaveCSS('padding-top', active ? '5.5px' : '6px');

      await page.getByRole('button', { name: '보기', exact: true }).click();
      await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
      const preview = page.locator('.print-layout');
      const finalCell = preview.locator('.print-content td').filter({ hasText: /^A7$/ });
      await expect(finalCell).toHaveCSS('padding-top', active ? '5.5px' : '6px');
      await expect(finalCell).toHaveCSS('border-bottom-width', '1px');
      await expect(finalCell).toHaveCSS('border-bottom-color', 'rgb(16, 32, 48)');
      const geometry = await preview.locator('.print-content td > p').evaluateAll(paragraphs => paragraphs.map(p => {
        const range = document.createRange();
        range.selectNodeContents(p);
        const ink = range.getBoundingClientRect(), cell = p.parentElement!.getBoundingClientRect();
        const sheet = p.closest('.print-page')!.getBoundingClientRect();
        return { label: p.textContent, page: Number(p.closest<HTMLElement>('.print-page')!.dataset.page),
          inside: ink.width > 0 && ink.height > 0 && ink.top >= cell.top && ink.bottom <= cell.bottom &&
            ink.top >= sheet.top && ink.bottom <= sheet.bottom };
      }));
      await info.attach('stage-' + pass, { body: JSON.stringify(geometry), contentType: 'application/json' });
      const labels = spans ? ['A1', 'B1', ...(mode === 'inline' ? ['B1b'] : []), 'A2', 'A3', 'A4', 'B4', 'A5', 'B5', 'A6', 'B6', 'A7', 'B7'] :
        Array.from({ length: 7 }, (_, i) => ['A' + (i + 1), 'B' + (i + 1)]).flat();
      expect(geometry.map(r => r.label)).toEqual(labels);
      if (spans) {
        await expect(editorTable.locator('td[rowspan="3"]')).toHaveCount(1);
        await expect(preview.locator('.print-content td[rowspan="3"]')).toHaveCount(1);
      }
      expect(geometry.every(r => r.inside)).toBe(true);
      if (mode !== 'inline') expect(geometry.at(-1)?.page).toBe(active ? 1 : 2);
      await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
      if (!active || pass === 2) break;
      await page.getByRole('button', { name: '파일', exact: true }).click();
      const pending = page.waitForEvent('download');
      await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
      const saved = info.outputPath(`textbox-save-centered-borders-${pass + 1}.docx`);
      await (await pending).saveAs(saved);
      bytes = await readFile(saved);
      const documentXml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
      expect(documentXml).toContain(`ws:centeredBorders="${mode === 'inline' ? 'inline' : '1'}"`);
      expect([...documentXml.matchAll(/<w:top w:type="dxa" w:w="90"\/>/g)]).toHaveLength(spans ? 14 : 16);
      expect([...documentXml.matchAll(/<w:bottom w:val="single" w:color="102030" w:sz="6"\/>/g)]).toHaveLength(mode === 'inline' ? 14 : 16);
      if (mode === 'inline')
        expect([...documentXml.matchAll(/<w:bottom w:val="double" w:color="102030" w:sz="6"\/>/g)]).toHaveLength(2);
    }
  });
}

test('a centered source grid can fit paint rounding without admitting a wrapped line', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { layoutHwpCenteredTableBorders, tableBorderFits } = await import('/src/hwp-centered-table-borders');
    const paper = document.createElement('div');
    paper.className = 'print-page';
    paper.style.cssText = 'position:fixed;left:20px;top:20px;width:240px;height:300px';
    const style = 'border:1px solid black;padding:4px';
    const line = 'line-height:20px;margin:0;font:12px/20px Arial';
    paper.innerHTML = `<div class="print-content">
      <table style="border-collapse:collapse;width:100px" data-hwp-centered-borders="true"><tbody>
        <tr><th style="${style}"><p data-word-line-rule="exact" style="${line}">Heading</p></th></tr>
        <tr><td style="${style}"><p data-word-line-rule="exact" style="${line}"><span style="vertical-align:1px">First</span> Second</p></td></tr>
      </tbody></table></div>`;
    document.body.append(paper);
    try {
      const body = paper.firstElementChild as HTMLElement;
      const table = body.querySelector('table')!;
      layoutHwpCenteredTableBorders(table);
      const row = table.rows[1], height = 2 * (20 + 4 + 4);
      const logicalBottom = table.rows[0].getBoundingClientRect().top + height;
      const physicalBottom = row.getBoundingClientRect().bottom;
      const fits = tableBorderFits(table, body, logicalBottom);
      const tight = tableBorderFits(table, body, logicalBottom - .25);
      const p = row.cells[0].querySelector('p')!;
      p.textContent = 'A longer phrase wraps naturally on the existing fixed column without an authored line break';
      const wrapped = tableBorderFits(table, body, logicalBottom);
      return { active:table.dataset.hwpCenteredBordersActive,
        physicalExcess:physicalBottom - logicalBottom, fits, tight, wrapped };
    } finally { paper.remove(); }
  });
  expect(result.active).toBe('true');
  expect(result.physicalExcess).toBeGreaterThan(0);
  expect(result.fits).toBe(true);
  expect(result.tight).toBe(false);
  expect(result.wrapped).toBe(false);
});

test('a headerless HWP grid keeps its stroke centres through two DOCX saves', async ({ page }, info) => {
  const labelCell = (text: string) => new TableCell({
    margins: margin, borders: { top: stroke, bottom: stroke, left: stroke, right: stroke },
    children: [new Paragraph({
      spacing: { before: 0, after: 0, line: 270, lineRule: LineRuleType.EXACT },
      children: [new TextRun({ text, font: 'Arial', size: 18 })],
    })],
  });
  const separator = new TableCell({
    rowSpan: 3, margins: margin, children: [new Paragraph({ text: '' })],
    borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE } },
  });
  const rows = ['Alpha', 'Bravo', 'Charlie'].map((label, index) =>
    new TableRow({ cantSplit: true, children: [labelCell(label), ...(index === 0 ? [separator] : [])] }));
  const table = new Table({ rows, layout: 'fixed', columnWidths: [2250, 750], width: { size: 3000, type: WidthType.DXA } });
  const initial = await Packer.toBuffer(new Document({ sections: [{ children: [table] }] }));
  const zip = await JSZip.loadAsync(initial);
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace('<w:tbl>',
    `<w:tbl xmlns:ws="${namespace}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="ws" mc:PreserveAttributes="ws:centeredBorders" ws:centeredBorders="inline">`));
  let bytes = await zip.generateAsync({ type: 'nodebuffer' });
  await page.goto('/');
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  const live = editor.locator('table');
  for (let pass = 0; pass < 3; pass++) {
    const filename = `hwp-headerless-grid-${pass}`;
    await page.locator('input[type=file]').first().setInputFiles({
      name: filename + '.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: bytes,
    });
    await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(filename);
    await expect(live.locator('th')).toHaveCount(0);
    await expect(live.locator('td[rowspan="3"]')).toHaveCount(1);
    await expect(live).toHaveAttribute('data-hwp-centered-borders-active', 'true');
    const first = live.locator('td').filter({ hasText: /^Alpha$/ });
    await expect(first).toHaveCSS('padding-top', '5.5px');

    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const printed = page.locator('.print-layout.preview table');
    await expect(printed.locator('td').filter({ hasText: /^Charlie$/ })).toHaveCSS('padding-bottom', '5.5px');
    await expect(printed.locator('td[rowspan="3"]')).toHaveCount(1);
    await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();

    if (pass === 2) {
      await live.locator('td[rowspan="3"] p').click();
      await page.keyboard.insertText('Live separator contents');
      await expect(live).not.toHaveAttribute('data-hwp-centered-borders-active', 'true');
      await expect(first).toHaveCSS('padding-top', '6px');
      await page.keyboard.press('ControlOrMeta+z');
      await expect(live).toHaveAttribute('data-hwp-centered-borders-active', 'true');
      break;
    }
    await page.getByRole('button', { name: '파일', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
    const saved = info.outputPath(`hwp-headerless-inline-save-${pass + 1}.docx`);
    await (await pending).saveAs(saved);
    bytes = await readFile(saved);
    const exported = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
    expect(exported).toContain('ws:centeredBorders="inline"');
    expect(exported).toContain('<w:vMerge');
    expect(exported).not.toContain('<w:tblHeader');
    expect(exported).toContain('<w:top w:type="dxa" w:w="90"/>');
  }
});

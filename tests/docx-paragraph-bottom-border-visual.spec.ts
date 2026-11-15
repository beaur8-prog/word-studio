import { expect, test, type Locator, type Page } from '@playwright/test';
import { BorderStyle, Document, HeadingLevel, LineRuleType, Packer, Paragraph, Table, TableCell, TableRow, WidthType } from 'docx';
import type { Editor } from '@tiptap/core';

const marker = 'data-word-bottom-border-layout';
const headingText = 'VISIBLE GROUP HEADING';
const boundaryText = 'PLAIN BODY BOUNDARY';
const noneText = 'EXPLICITLY CLEARED HEADING';
const cellText = 'CELL HEADING CONTROL';
const plainCellText = 'PLAIN CELL CONTROL';
const visible = [headingText, boundaryText, noneText, cellText, plainCellText];

async function fixture() {
  const bottom = { style: BorderStyle.SINGLE, color: '94C600', size: 16, space: 3 };
  const heading = { heading: HeadingLevel.HEADING_1 } as const;
  return Packer.toBuffer(new Document({
    styles: { default: {
      document: { run: { font: 'Arial', size: 22, color: '101010' } },
      heading1: {
        run: { font: 'Arial', size: 44, color: '101010', bold: false },
        paragraph: { border: { bottom }, spacing: { before: 0, after: 180, line: 600, lineRule: LineRuleType.EXACT } },
      },
    } },
    sections: [{ children: [
      new Paragraph({ ...heading, children: [] }),
      new Paragraph({ ...heading, text: headingText }),
      new Paragraph({ text: boundaryText, spacing: { before: 0, after: 180 } }),
      new Paragraph({ ...heading, text: noneText, border: { bottom: { style: BorderStyle.NONE } } }),
      new Table({ width: { size: 6800, type: WidthType.DXA }, columnWidths: [3800, 3000], rows: [
        new TableRow({ children: [
          new TableCell({ width: { size: 3800, type: WidthType.DXA }, children: [
            new Paragraph({ ...heading, text: cellText, border: { bottom } }),
          ] }),
          new TableCell({ width: { size: 3000, type: WidthType.DXA }, children: [
            new Paragraph(plainCellText),
          ] }),
        ] }),
      ] }),
    ] }],
  }));
}

async function inspectGeometry(root: Locator) {
  return root.evaluate((root, marker) => {
    const origin = root.getBoundingClientRect();
    const selected = (query: string) => [...root.querySelectorAll<HTMLElement>(query)];
    const headings = selected(':scope > h1');
    return {
      markerText: selected('[' + marker + ']').map(item => item.textContent),
      cellMarkers: selected('td [' + marker + ']').length,
      emptyHeading: headings[0]?.textContent,
      emptyHeight: headings[0]?.getBoundingClientRect().height ?? 0,
      rows: selected('h1,p').filter(p => p.textContent?.trim()).map(p => {
        const glyphs = [];
        const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
        while (walker.nextNode()) {
          const node = walker.currentNode as Text;
          for (const grapheme of new Intl.Segmenter().segment(node.data)) {
            if (!grapheme.segment.trim()) continue;
            const range = document.createRange();
            range.setStart(node, grapheme.index);
            range.setEnd(node, grapheme.index + grapheme.segment.length);
            const rect = range.getBoundingClientRect();
            glyphs.push({ value: grapheme.segment, x: rect.left - origin.left,
              y: rect.top - origin.top, width: rect.width, height: rect.height });
          }
        }
        return { text: p.textContent!, glyphs };
      }),
    };
  }, marker);
}

async function paintedRules(page: Page, png: Buffer) {
  return page.evaluate(async b64 => {
    const image = new Image();
    image.src = 'data:image/png;base64,' + b64;
    await image.decode();
    const canvas = document.createElement('canvas');
    canvas.width = image.width;
    canvas.height = image.height;
    const ctx = canvas.getContext('2d')!;
    ctx.drawImage(image, 0, 0);
    const { data } = ctx.getImageData(0, 0, image.width, image.height);
    const bands: { top: number; bottom: number; width: number }[] = [];
    for (let y = 0; y < image.height; y++) {
      let width = 0;
      for (let x = 0; x < image.width; x++) {
        const i = 4 * (y * image.width + x);
        if (Math.abs(data[i] - 148) <= 8 && Math.abs(data[i + 1] - 198) <= 8 && data[i + 2] <= 12 && data[i + 3] > 220) width++;
      }
      if (width < Math.max(60, image.width / 3)) continue;
      const current = bands.at(-1);
      if (current && current.bottom === y) {
        current.bottom = y + 1;
        current.width = Math.max(width, current.width);
      } else bands.push({ top: y, bottom: y + 1, width });
    }
    return bands;
  }, png.toString('base64'));
}

test('a grouped Word paragraph bottom rule paints once beside the intact editor and print text', async ({ page }, info) => {
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'independent-visual-bottom-border.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: await fixture(),
  });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  await expect(editor.locator(':scope > h1')).toHaveCount(3);
  await expect(editor.locator('td')).toHaveCount(2);
  await page.evaluate(() => document.fonts.ready);
  const snapshot = () => editor.evaluate(el => (el as HTMLElement & { editor: Editor }).editor.getJSON());
  const before = await snapshot();

  async function checkView(root: Locator, prefix: string) {
    await expect(root.locator('[' + marker + ']')).toHaveText([headingText]);
    const geometry = await inspectGeometry(root);
    expect(geometry.markerText).toEqual([headingText]);
    expect(geometry.cellMarkers).toBe(0);
    expect(geometry.emptyHeading).toBe('');
    expect(geometry.emptyHeight).toBeGreaterThan(20);
    expect(geometry.rows.map(row => row.text)).toEqual(visible);
    for (const row of geometry.rows) {
      expect(row.glyphs.map(glyph => glyph.value).join('')).toBe(row.text.replace(/\s/g, ''));
      expect(row.glyphs.every(glyph => glyph.width > 0 && glyph.height > 0 && glyph.x >= 0 && glyph.y >= 0)).toBe(true);
    }

    // The screenshot verifies paint even when a renderer leaves a valid marker
    // on the node but loses its pseudo-element or paints a second cloned rule.
    const png = await root.screenshot({ path: info.outputPath(prefix + '-paragraph-bottom-rule.png'), scale: 'css' });
    const bands = await paintedRules(page, png);
    expect(bands).toHaveLength(1);
    const heading = geometry.rows[0].glyphs, next = geometry.rows[1].glyphs;
    expect(bands[0].top).toBeGreaterThan(Math.max(...heading.map(glyph => glyph.y + glyph.height)) - 1);
    expect(bands[0].bottom).toBeLessThan(Math.min(...next.map(glyph => glyph.y)) + 1);
    return geometry;
  }

  const edited = await checkView(editor, 'editor');
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  await expect(page.locator('.print-layout.preview .print-page')).toHaveCount(1);
  const printed = await checkView(page.locator('.print-layout.preview .print-content'), 'print');
  expect(Math.abs(printed.emptyHeight - edited.emptyHeight)).toBeLessThan(1);
  await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
  expect(await snapshot()).toEqual(before);
});

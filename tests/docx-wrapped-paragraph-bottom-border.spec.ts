import { expect, test, type Locator, type Page } from '@playwright/test';
import {
  AlignmentType, BorderStyle, Document, FootnoteReferenceRun, HeadingLevel, HorizontalPositionRelativeFrom,
  ImageRun, LineRuleType, PageNumber, Packer, Paragraph, TextRun, TextWrappingType, VerticalPositionRelativeFrom,
} from 'docx';

const title = 'WRAPPED GROUP TITLE';
const body = 'AFTER THE GROUP';
const footnote = 'SYNTHETIC NOTE BODY';
const fieldSelector = '[data-page-number],[data-document-field]';
const marker = '[data-word-bottom-border-layout]';
const squareName = 'SQUARE PICTURE';
const overlayName = 'NONWRAPPING PICTURE';
const lineTwips = 480;
const gapPoints = 3;
const strokeEighths = 16;
const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==', 'base64');

const cases: Scenario[] = [
  { name: 'shared right slot', squareHeight: 152, paints: true },
  { name: 'stroke below picture', squareHeight: 59 },
  { name: 'multiple wrapped lines', squareHeight: 152, multiline: true },
  { name: 'note on later page', squareHeight: 152, laterNote: true },
  { name: 'inline PAGE field', pageField: true },
  { name: 'nonwrapping picture crossing stroke', squareHeight: 152, overlay: true },
];

interface Scenario {
  name: string;
  squareHeight?: number;
  multiline?: boolean;
  laterNote?: boolean;
  pageField?: boolean;
  overlay?: boolean;
  paints?: boolean;
}

type Wrap = (typeof TextWrappingType)[keyof typeof TextWrappingType];

function picture(name: string, width: number, height: number, x: number, y: number, wrap: Wrap) {
  return new ImageRun({
    type: 'png', data: png, transformation: { width, height },
    altText: { name, title: name, description: name },
    floating: {
      horizontalPosition: { relative: HorizontalPositionRelativeFrom.MARGIN, offset: x * 9525 },
      verticalPosition: { relative: VerticalPositionRelativeFrom.PAGE, offset: y * 9525 },
      wrap: { type: wrap },
    },
  });
}

async function fixture(scenario: Scenario, text: string) {
  const heading = { heading: HeadingLevel.HEADING_1 } as const;
  const children: Paragraph[] = [];
  if (scenario.squareHeight) {
    const images = [picture(squareName, 220, scenario.squareHeight, 0, 96, TextWrappingType.SQUARE)];
    if (scenario.overlay) images.push(picture(overlayName, 60, 35, 540, 158, TextWrappingType.NONE));
    children.push(new Paragraph({
      spacing: { before: 0, after: 0, line: 15, lineRule: LineRuleType.EXACT }, children: images,
    }));
  }
  children.push(
    new Paragraph({ ...heading, children: [] }),
    new Paragraph({
      ...heading,
      children: scenario.pageField
        ? [new TextRun(text + ' '), new TextRun({ children: [PageNumber.CURRENT] })]
        : [new TextRun(text)],
    }),
    new Paragraph({ text: body, spacing: { before: 0, after: 180 } }),
  );
  if (scenario.laterNote) children.push(new Paragraph({
    pageBreakBefore: true, children: [new TextRun('LATER PAGE NOTE'), new FootnoteReferenceRun(1)],
  }));
  return Packer.toBuffer(new Document({
    styles: { default: {
      document: { run: { font: 'Arial', color: '101010', size: 22 } },
      heading1: {
        run: { font: 'Arial', color: '101010', size: 36, bold: false },
        paragraph: {
          alignment: AlignmentType.LEFT,
          spacing: { before: 0, after: 90, line: lineTwips, lineRule: LineRuleType.EXACT },
          border: { bottom: { style: BorderStyle.SINGLE, color: '94C600', size: strokeEighths, space: gapPoints } },
        },
      },
    } },
    ...(scenario.laterNote ? { footnotes: { '1': { children: [new Paragraph(footnote)] } } } : {}),
    sections: [{ children }],
  }));
}

async function inspect(root: Locator) {
  return root.evaluate((root, names) => {
    const origin = root.getBoundingClientRect();
    const [empty, title] = root.querySelectorAll(':scope > h1');
    const rect = (value: DOMRect) => ({
      left: value.left - origin.left, right: value.right - origin.left,
      top: value.top - origin.top, bottom: value.bottom - origin.top,
    });
    const image = (name: string) => {
      const image = root.querySelector('img[alt="' + name + '"]');
      return image ? rect(image.getBoundingClientRect()) : null;
    };
    const glyphs = [];
    const walker = document.createTreeWalker(title, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const node = walker.currentNode as Text;
      for (const { segment, index } of new Intl.Segmenter().segment(node.data)) {
        if (!segment.trim()) continue;
        const range = document.createRange();
        range.setStart(node, index);
        range.setEnd(node, index + segment.length);
        glyphs.push({ value: segment, ...rect(range.getBoundingClientRect()) });
      }
    }
    return {
      emptyHeight: empty.getBoundingClientRect().height, title: rect(title.getBoundingClientRect()), glyphs,
      square: image(names.squareName), overlay: image(names.overlayName),
    };
  }, { squareName, overlayName });
}

async function paintedBands(page: Page, capture: Buffer) {
  return page.evaluate(async base64 => {
    const image = new Image();
    image.src = 'data:image/png;base64,' + base64;
    await image.decode();
    const canvas = document.createElement('canvas');
    canvas.width = image.width;
    canvas.height = image.height;
    const ctx = canvas.getContext('2d')!;
    ctx.drawImage(image, 0, 0);
    const { data } = ctx.getImageData(0, 0, image.width, image.height);
    const bands: { top: number; bottom: number; left: number; right: number }[] = [];
    for (let y = 0; y < image.height; y++) {
      let count = 0, left = image.width, right = 0;
      for (let x = 0; x < image.width; x++) {
        const pixel = 4 * (y * image.width + x);
        if (Math.abs(data[pixel] - 148) <= 8 && Math.abs(data[pixel + 1] - 198) <= 8 &&
            data[pixel + 2] <= 12 && data[pixel + 3] > 220) {
          count++;
          left = Math.min(left, x);
          right = x + 1;
        }
      }
      if (count < Math.max(60, image.width / 3)) continue;
      const current = bands.at(-1);
      if (current && current.bottom === y) {
        current.bottom = y + 1;
        current.left = Math.min(current.left, left);
        current.right = Math.max(current.right, right);
      } else bands.push({ top: y, bottom: y + 1, left, right });
    }
    return bands;
  }, capture.toString('base64'));
}

for (const scenario of cases) {
  test(`paints a wrapped paragraph border only in a clear matching slot · ${scenario.name}`, async ({ page }, info) => {
    const text = scenario.multiline ? (title + ' ').repeat(5).trim() : title;
    await page.goto('/');
    await page.locator('input[type=file]').first().setInputFiles({
      name: 'paragraph-rule.docx', buffer: await fixture(scenario, text),
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
    const heading = editor.locator(':scope > h1').last();
    await expect(heading).toContainText(text);
    await page.evaluate(() => document.fonts.ready);
    if (scenario.squareHeight) await expect.poll(() => heading.locator('[data-flow-widget]').count()).toBeGreaterThan(0);
    if (scenario.laterNote) await expect(editor.locator('[data-note-reference]')).toHaveCount(1);

    async function check(root: Locator, mode: string) {
      const headings = root.locator(':scope > h1');
      await expect(headings).toHaveCount(2);
      await expect(headings.first()).toBeEmpty();
      await expect(root.getByText(body, { exact: true })).toHaveCount(1);
      if (scenario.pageField) await expect(headings.last().locator(fieldSelector)).toHaveCount(1);

      const frame = await inspect(root);
      expect(frame.emptyHeight).toBeGreaterThan(20);
      if (!scenario.pageField) expect(frame.glyphs.map(g => g.value).join('')).toBe(text.replace(/\s/g, ''));
      expect(frame.glyphs.every(g => g.right > g.left && g.bottom > g.top)).toBe(true);
      expect(new Set(frame.glyphs.map(g => Math.round(g.top))).size > 1).toBe(!!scenario.multiline);
      if (scenario.squareHeight) {
        expect(frame.square).not.toBeNull();
        expect(frame.glyphs[0].left).toBeGreaterThan(frame.square!.right - 1);
        expect(frame.glyphs[0].top).toBeLessThan(frame.square!.bottom);
      }
      const strokeTop = frame.title.top + (lineTwips / 20 + gapPoints) * 4 / 3;
      if (scenario.squareHeight === 59) expect(strokeTop).toBeGreaterThan(frame.square!.bottom);
      if (scenario.overlay) {
        expect(frame.overlay).not.toBeNull();
        expect(Math.max(...frame.glyphs.map(g => g.right))).toBeLessThan(frame.overlay!.left);
        expect(frame.overlay!.right).toBeLessThan(frame.title.right);
        expect(frame.overlay!.top).toBeLessThan(strokeTop);
        expect(frame.overlay!.bottom).toBeGreaterThan(strokeTop + strokeEighths / 8 * 4 / 3);
      }

      const expected = scenario.paints ? 1 : 0;
      expect.soft(await root.locator(marker).count(), mode + ' paint marker').toBe(expected);
      const capture = await root.screenshot({ scale: 'css', path: info.outputPath(mode + '.png') });
      const bands = await paintedBands(page, capture);
      expect.soft(bands, mode + ' visible rule').toHaveLength(expected);
      if (scenario.paints && bands.length === 1) {
        const [rule] = bands;
        await expect(headings.last()).toHaveAttribute('data-word-bottom-border-floating');
        expect(Math.abs(rule.left - frame.square!.right)).toBeLessThan(2);
        expect(Math.abs(rule.right - frame.title.right)).toBeLessThan(2);
        expect(rule.top).toBeGreaterThan(Math.max(...frame.glyphs.map(g => g.bottom)) - 1);
        expect(rule.bottom).toBeLessThan(frame.square!.bottom + 1);
      }
    }

    await check(editor, 'editor');
    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const sheets = page.locator('.print-layout.preview .print-page');
    await expect(sheets).toHaveCount(scenario.laterNote ? 2 : 1);
    const first = sheets.first().locator('.print-content');
    if (scenario.laterNote) {
      await expect(first.locator('[data-note-reference]')).toHaveCount(0);
      await expect(sheets.last().locator('[data-note-reference]')).toHaveCount(1);
      await expect(sheets.last().locator('.print-footnotes')).toContainText(footnote);
    }
    await check(first, 'print');
  });
}

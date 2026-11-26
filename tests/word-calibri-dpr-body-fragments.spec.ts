import { expect, test, type Locator } from '@playwright/test';
import { Document, LineRuleType, Packer, Paragraph, TextRun } from 'docx';

test.use({ deviceScaleFactor: 2 });

const faceAsset = '/src/assets/fonts/texgyreadventor-regular.otf';
const fallbackProbe = 'Studio Apple Gothic space probe';
const mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const run = { font: 'Calibri', size: 22 };
const wordLine = 276;
const rawFraction = .5;

function longText(prefix: string) {
  return Array.from({ length: 48 }, (_, i) => `${prefix}${i.toString().padStart(2, '0')}sample`).join(' ');
}

const paragraphs = [
  { id: 'QaWhole', ascii: 'Uniform short paragraph', tail: '', split: false, eligible: true },
  { id: 'QaAscii', ascii: longText('AS'), tail: '', split: true, eligible: true },
  { id: 'QaMixed', ascii: longText('MX'), tail: ' Georgia ending', font: 'Georgia', split: true, eligible: false },
  { id: 'QaNbsp', ascii: longText('NB'), tail: ' tail pair', split: true, eligible: false },
  { id: 'QaAccent', ascii: longText('AC'), tail: ' tail café', split: true, eligible: false },
];

async function fixture() {
  return Packer.toBuffer(new Document({
    creator: 'Independent layout QA author',
    description: 'Self-authored Calibri continuation fixture',
    compatibility: { wrapTrailSpaces: true },
    styles: {
      default: { document: { run } },
      paragraphStyles: paragraphs.map(({ id }) => ({
        id, name: id, basedOn: 'Normal', run,
        paragraph: { spacing: { before: 0, after: 0, line: wordLine, lineRule: LineRuleType.AUTO } },
      })),
    },
    // 320 × 260 CSS px, with 260 × 200 inside the margins. Each long
    // paragraph wraps past one body without an explicit break inside a run.
    sections: [{
      properties: { page: { size: { width: 4800, height: 3900 }, margin: { top: 450, right: 450, bottom: 450, left: 450 } } },
      children: paragraphs.map(({ id, ascii, tail, font }) => new Paragraph({
        style: id,
        children: [
          new TextRun({ ...run, text: ascii }),
          ...(tail ? [new TextRun({ ...run, font: font ?? run.font, text: tail })] : []),
        ],
      })),
    }],
  }));
}

function snapshot(root: Locator) {
  return root.locator('p[data-word-paragraph-style^="Qa"]').evaluateAll(elements => elements.map(element => {
    const p = element as HTMLParagraphElement;
    const css = getComputedStyle(p);
    const range = document.createRange();
    range.selectNodeContents(p);
    const bounds = range.getBoundingClientRect();
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    const families = new Set<string>();
    while (walker.nextNode()) {
      const node = walker.currentNode;
      if (node.textContent?.trim() && node.parentElement) families.add(getComputedStyle(node.parentElement).fontFamily);
    }
    return {
      id: p.dataset.wordParagraphStyle,
      text: p.textContent ?? '',
      sourceRule: p.dataset.wordLineRule,
      sourceLine: Number(p.dataset.wordLine),
      lineHeight: parseFloat(css.lineHeight),
      fontFamilies: [...families],
      fragment: p.hasAttribute('data-print-paragraph-fragment'),
      inkHeight: bounds.height,
    };
  }));
}

test('Calibri auto line metrics keep the complete paragraph decision when print splits its ASCII prefix', async ({ page }) => {
  await page.addInitScript(({ faceAsset, fallbackProbe, rawFraction }) => {
    const NativeFontFace = window.FontFace;
    window.FontFace = class extends NativeFontFace {
      constructor(family: string, source: string | BufferSource, options?: FontFaceDescriptors) {
        if (family === 'Studio Calibri presence probe') source = 'local("WS Calibri QA deliberately unavailable")';
        if (family === fallbackProbe) source = `url("${faceAsset}")`;
        super(family, source, options);
      }
    };
    const originalRect = Element.prototype.getBoundingClientRect;
    Element.prototype.getBoundingClientRect = function () {
      if (!(this instanceof HTMLSpanElement)) return originalRect.call(this);
      const { style } = this;
      if (style.position !== 'fixed' || style.visibility !== 'hidden' || style.left !== '-10000px' || style.whiteSpace !== 'pre')
        return originalRect.call(this);
      const family = style.fontFamily.split(',')[0].replace(/["']/g, '');
      if (style.fontSize === '64px' && family === 'Calibri' && ['Wm', 'i', '0', ' '].includes(this.textContent ?? '')) {
        const original = style.fontFamily;
        style.fontFamily = JSON.stringify(fallbackProbe);
        try { return originalRect.call(this); }
        finally { style.fontFamily = original; }
      }
      if (style.lineHeight !== 'normal' || !['Calibri', 'Georgia'].includes(family) ||
          Math.abs(parseFloat(style.fontSize) - 44 / 3) > .001) return originalRect.call(this);
      const rect = originalRect.call(this);
      const canvas = document.createElement('canvas').getContext('2d')!;
      canvas.font = `${style.fontSize} ${JSON.stringify(fallbackProbe)}`;
      const metrics = canvas.measureText('Ag');
      // Model the observed half-pixel quantization in hidden fallback probes.
      // All actual paragraph widths, wrappers, and text ranges use real layout.
      return new DOMRect(rect.x, rect.y, rect.width,
        metrics.fontBoundingBoxAscent + metrics.fontBoundingBoxDescent + rawFraction);
    };
  }, { faceAsset, fallbackProbe, rawFraction });

  await page.goto('/');
  await expect.poll(() => page.evaluate(family => [...document.fonts].some(face =>
    face.family.replace(/["']/g, '') === family && face.status === 'loaded'), fallbackProbe)).toBe(true);
  expect(await page.evaluate(() => devicePixelRatio)).toBe(2);
  const strut = await page.evaluate(family => {
    const canvas = document.createElement('canvas').getContext('2d')!;
    canvas.font = `14.6667px ${JSON.stringify(family)}`;
    const metrics = canvas.measureText('Ag');
    return metrics.fontBoundingBoxAscent + metrics.fontBoundingBoxDescent;
  }, fallbackProbe);
  expect(strut).toBeGreaterThan(0);

  await page.locator('input[type=file]').first().setInputFiles({
    name: 'calibri-continuations.docx', mimeType, buffer: await fixture(),
  });
  await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue('calibri-continuations');
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  const multiplier = wordLine / 240;
  const expectedHeight = (eligible: boolean) => (strut + (eligible ? 0 : rawFraction)) * multiplier;
  await expect.poll(async () => {
    const rows = await snapshot(editor);
    return rows.length === paragraphs.length && paragraphs.every((p, index) =>
      Math.abs(rows[index].lineHeight - expectedHeight(p.eligible)) < .02);
  }).toBe(true);

  const editing = await snapshot(editor);
  expect(editing.map(row => row.id)).toEqual(paragraphs.map(p => p.id));
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  const print = page.locator('.print-layout');
  await expect.poll(() => print.locator('.print-page').count()).toBeGreaterThan(1);
  const printed = await snapshot(print.locator('.print-content'));

  for (const [index, source] of paragraphs.entries()) {
    const whole = editing[index];
    const fragments = printed.filter(row => row.id === source.id);
    const exactText = source.ascii + source.tail;
    expect(whole.text, source.id).toBe(exactText);
    expect(fragments.map(p => p.text).join(''), source.id).toBe(exactText);
    expect(fragments.length, source.id).toBeGreaterThanOrEqual(source.split ? 2 : 1);
    if (!source.split) expect(fragments).toHaveLength(1);
    for (const fragment of fragments) {
      expect(fragment.fragment, source.id).toBe(source.split);
      expect(fragment.sourceRule, source.id).toBe(whole.sourceRule);
      expect(fragment.sourceLine, source.id).toBeCloseTo(multiplier, 3);
      expect(fragment.lineHeight, source.id).toBeCloseTo(whole.lineHeight, 2);
      expect(fragment.inkHeight, source.id).toBeGreaterThan(0);
    }
    if (!source.eligible) {
      // A fitter that reevaluates this prefix alone would wrongly opt in.
      expect(fragments[0].text).toMatch(/^[\x20-\x7e]+$/);
      expect(fragments[0].fontFamilies).toHaveLength(1);
      if (source.font) expect(whole.fontFamilies).toHaveLength(2);
      expect(fragments[0].lineHeight).toBeCloseTo(expectedHeight(false), 2);
    }
  }
});

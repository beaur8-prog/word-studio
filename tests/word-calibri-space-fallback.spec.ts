import { expect, test, type Locator, type Page } from '@playwright/test';
import { Document, Packer, Paragraph, Tab, TextRun } from 'docx';

const knownFaceAsset = '/src/assets/fonts/texgyreadventor-regular.otf';
const appleProbe = 'Studio Apple Gothic space probe';
const testTexts = [
  'Ordinary A B C',
  'Authored tracking unchanged',
  'NBSP A\u00a0B C',
  'Tab A\tB C',
  'Mixed A B C',
];

async function fixture() {
  const run = { font: 'Calibri', size: 24 };
  return Packer.toBuffer(new Document({
    creator: 'Independent layout QA author',
    description: 'Self-authored Calibri spacing eligibility fixture',
    compatibility: { wrapTrailSpaces: true },
    styles: { default: { document: { run, paragraph: { spacing: { line: 240, after: 0 } } } } },
    sections: [{ children: [
      new Paragraph({ children: [new TextRun({ ...run, text: testTexts[0] })] }),
      new Paragraph({ children: [new TextRun({ ...run, text: testTexts[1], characterSpacing: 20 })] }),
      new Paragraph({ children: [new TextRun({ ...run, text: testTexts[2] })] }),
      new Paragraph({ children: [new TextRun({ ...run, children: ['Tab A', new Tab(), 'B C'] })] }),
      new Paragraph({ children: [new TextRun({ ...run, text: 'Mixed A ' }), new TextRun({ ...run, text: 'B C', font: 'Georgia' })] }),
    ] }],
  }));
}

function read(root: Locator) {
  return root.locator('p').evaluateAll(paragraphs => paragraphs.filter(p => p.textContent?.trim()).map(p => ({
    text: p.textContent,
    wordSpacing: parseFloat(getComputedStyle(p).wordSpacing) || 0,
    letterSpacing: [...p.querySelectorAll('span')].map(span => parseFloat(getComputedStyle(span).letterSpacing) || 0),
    hasTab: !!p.querySelector('[data-docx-tab]'),
  })));
}

async function preview(page: Page) {
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  const root = page.locator('.print-layout');
  await expect(root.locator('.print-page')).toHaveCount(1);
  return root.locator('.print-content');
}

test('measured Calibri fallback corrects ordinary spaces and stops when Calibri becomes available', async ({ page }) => {
  await page.addInitScript(({ knownFaceAsset, appleProbe }) => {
    const NativeFontFace = window.FontFace;
    window.FontFace = class extends NativeFontFace {
      constructor(family: string, source: string | BufferSource, options?: FontFaceDescriptors) {
        if (family === 'Studio Calibri presence probe') source = 'local("WS Calibri QA deliberately unavailable")';
        // A known bundled face stands in for the Apple probe on any host. Its
        // space is wider than Calibri's, so the candidate must apply a delta.
        if (family === appleProbe) source = `url("${knownFaceAsset}")`;
        super(family, source, options);
      }
    };
    const boundingRect = Element.prototype.getBoundingClientRect;
    Element.prototype.getBoundingClientRect = function () {
      if (!(this instanceof HTMLSpanElement)) return boundingRect.call(this);
      const { style } = this;
      const original = style.fontFamily;
      // Model a matching resolved fallback only in the offscreen DOM probe.
      // The document itself still uses the browser's normal font resolution.
      if (style.position !== 'fixed' || style.visibility !== 'hidden' || style.fontSize !== '64px' ||
          !['Wm', 'i', '0', ' '].includes(this.textContent ?? '') || original.replace(/["']/g, '') !== 'Calibri')
        return boundingRect.call(this);
      style.fontFamily = JSON.stringify(appleProbe);
      try { return boundingRect.call(this); }
      finally { style.fontFamily = original; }
    };
  }, { knownFaceAsset, appleProbe });
  await page.goto('/');
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'calibri-space-eligibility.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: await fixture(),
  });
  await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue('calibri-space-eligibility');
  await expect.poll(() => page.evaluate(family => [...document.fonts].some(face =>
    face.family.replace(/["']/g, '') === family && face.status === 'loaded'), appleProbe)).toBe(true);

  const verify = async (root: Locator, eligible: boolean) => {
    await expect.poll(async () => (await read(root))[0]?.wordSpacing).toBeLessThan(eligible ? -.4 : .01);
    const rows = await read(root);
    expect(rows.map(row => row.text)).toEqual(testTexts);
    expect(rows.slice(1).map(row => row.wordSpacing)).toEqual([0, 0, 0, 0]);
    expect(rows[1].letterSpacing.some(spacing => spacing > 0)).toBe(true);
    expect(rows.map(row => row.hasTab)).toEqual([false, false, false, true, false]);
    if (!eligible) expect(rows[0].wordSpacing).toBe(0);
  };
  await verify(editor, true);
  await verify(await preview(page), true);
  await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();

  await page.evaluate(async source => {
    const face = new FontFace('Calibri', `url("${source}?late-calibri")`);
    // Register before loading so the real FontFaceSet event invalidates the
    // existing editor decoration without a document edit or synthetic event.
    document.fonts.add(face);
    await face.load();
    await document.fonts.ready;
  }, knownFaceAsset);
  await expect.poll(async () => (await read(editor))[0]?.wordSpacing).toBe(0);
  await verify(editor, false);
  await verify(await preview(page), false);
});

import { expect, test, type Locator, type Page } from '@playwright/test';
import { Document, LevelFormat, LineRuleType, Packer, Paragraph, TextRun } from 'docx';
import JSZip from 'jszip';

const family = 'Century Gothic';
const bodyText = 'BODY shows normal Word line spacing.   Next words fit naturally.';
const bulletText = 'BULLET shows normal Word line spacing.   Next words fit naturally.';
const mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

async function fixture(legacy: boolean) {
  const run = { font: family, size: 22 };
  const bytes = await Packer.toBuffer(new Document({
    creator: legacy ? 'Word Studio' : 'Independent QA author',
    description: legacy ? 'Created with Word Studio' : 'Independent OOXML layout fixture',
    compatibility: legacy ? undefined : { wrapTrailSpaces: true },
    styles: {
      default: { document: { run } },
      paragraphStyles: [{
        id: 'BodyCheck', name: 'Body Check', run,
        paragraph: { spacing: { before: 0, after: 0, line: 276, lineRule: LineRuleType.AUTO } },
      }],
    },
    numbering: { config: [{
      reference: 'symbol',
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: '\uf0b7',
        style: { run: { font: 'Symbol' }, paragraph: { indent: { left: 720, hanging: 360 } } },
      }],
    }] },
    sections: [{ children: [
      new Paragraph({ style: 'BodyCheck', children: [new TextRun({ text: bodyText, ...run })] }),
      new Paragraph({ style: 'BodyCheck', numbering: { reference: 'symbol', level: 0 },
        children: [new TextRun({ text: bulletText, ...run })] }),
    ] }],
  }));
  if (legacy) return bytes;
  const zip = await JSZip.loadAsync(bytes);
  const settings = await zip.file('word/settings.xml')!.async('string');
  const flag = '<w:wrapTrailSpaces/>';
  if (!settings.includes(flag)) throw Error('The fixture needs a wrapping flag in w:compat');
  zip.file('word/settings.xml', settings.replace(flag, '<w:wrapTrailSpaces w:val="true"/>'));
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, name: string, bytes: Buffer) {
  await page.locator('input[type=file]').first().setInputFiles({ name: name + '.docx', mimeType, buffer: bytes });
  const title = page.getByRole('textbox', { name: '문서 제목', exact: true });
  const protect = page.getByRole('dialog', { name: '변경 내용 보관' });
  await expect.poll(async () => await title.inputValue() === name || await protect.isVisible()).toBe(true);
  if (await protect.isVisible()) await protect.getByRole('button', { name: '복구본 보관 후 계속' }).click();
  await expect(title).toHaveValue(name);
  await page.evaluate(async () => { await document.fonts.ready; });
}

function measure(root: Locator) {
  return root.evaluate((el, { bodyText, bulletText }) => {
    const body = [...el.querySelectorAll('p')].find(p => p.textContent === bodyText);
    const item = el.querySelector('li[data-list-font="Symbol"]');
    const bullet = item?.querySelector('p');
    if (!body || !item || !bullet || bullet.textContent !== bulletText) {
      throw Error('Expected the two unchanged synthetic paragraphs');
    }
    return {
      whiteSpace: getComputedStyle(body).whiteSpace,
      bodyLine: parseFloat(getComputedStyle(body).lineHeight),
      bodyHeight: body.getBoundingClientRect().height,
      bulletLine: parseFloat(getComputedStyle(bullet).lineHeight),
      itemLine: parseFloat(getComputedStyle(item).lineHeight),
      itemExtra: item.getBoundingClientRect().height - bullet.getBoundingClientRect().height,
      marker: getComputedStyle(item, '::marker').content,
    };
  }, { bodyText, bulletText });
}

test('an explicit Word wrapping flag enables body and Symbol metrics independently of whitespace mode', async ({ page }) => {
  await page.addInitScript(() => {
    const NativeFontFace = window.FontFace;
    window.FontFace = class extends NativeFontFace {
      private missingCentury: boolean;
      constructor(family: string, source: string | BufferSource, descriptors?: FontFaceDescriptors) {
        super(family, source, descriptors);
        this.missingCentury = family === 'Century Gothic' && typeof source === 'string' && source.startsWith('local(');
      }
      load() {
        return this.missingCentury
          ? Promise.reject(new DOMException('Century Gothic unavailable in this fixture', 'NetworkError'))
          : super.load();
      }
    };
  });
  await page.goto('/');
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  const results = [];
  for (const [name, legacy] of [['word-explicit-trailing-spaces', false], ['legacy-studio-trailing-spaces', true]] as const) {
    await open(page, name, await fixture(legacy));
    // Wait for the bundled regular face and the asynchronous line decorations.
    await expect.poll(() => page.evaluate(() => [...document.fonts].some(face =>
      face.family.replace(/["']/g, '') === 'Century Gothic' && face.weight === '400' && face.style === 'normal' && face.status === 'loaded',
    ))).toBe(true);
    await expect.poll(async () => (await measure(editor)).bulletLine).toBeCloseTo(20.1, 1);
    const editing = await measure(editor);
    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const preview = page.locator('.print-layout');
    await expect(preview.locator('.print-page')).toHaveCount(1);
    await page.getByRole('button', { name: '인쇄 미리보기 배율 초기화', exact: true }).click();
    results.push({ editing, printing: await measure(preview.locator('.print-content')) });
    await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
  }

  const [word, legacy] = results;
  for (const view of ['editing', 'printing'] as const) {
    expect([word[view].whiteSpace, legacy[view].whiteSpace]).toEqual(['break-spaces', 'break-spaces']);
    expect([word[view].marker, legacy[view].marker]).toEqual(['"• "', '"• "']);
    expect(word[view].bodyHeight - legacy[view].bodyHeight).toBeGreaterThan(.4);
    expect(word[view].bodyLine - legacy[view].bodyLine).toBeGreaterThan(.4);
    expect(word[view].bulletLine).toBeCloseTo(legacy[view].bulletLine, 2);
    expect(word[view].itemLine).toBeCloseTo(word[view].bulletLine, 2);
    expect(word[view].itemExtra).toBeCloseTo(0, 1);
    expect(legacy[view].itemLine).toBeGreaterThan(legacy[view].bulletLine + 1);
    expect(legacy[view].itemExtra).toBeGreaterThan(.5);
  }
});

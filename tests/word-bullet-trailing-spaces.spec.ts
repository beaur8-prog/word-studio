import { expect, test, type Locator, type Page } from '@playwright/test';
import { Document, LevelFormat, LevelSuffix, Packer, Paragraph, TabStopType, TextRun } from 'docx';
import JSZip from 'jszip';

const mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const tail = 'TAIL';
const proof = (head: string) => head + ' '.repeat(72) + tail;
const cases = [
  { head: 'EXPLICIT', numberingTab: 1080, aligned: true },
  { head: 'IMPLICIT', numberingTab: undefined, aligned: true },
  { head: 'OFFSET', numberingTab: 1320, aligned: false },
] as const;

async function fixture(wrapTrailSpaces: boolean) {
  const bytes = await Packer.toBuffer(new Document({
    creator: 'Independent QA author',
    description: 'Independent OOXML Symbol list fixture',
    compatibility: wrapTrailSpaces ? { wrapTrailSpaces: true } : undefined,
    styles: { default: { document: {
      run: { font: 'Courier New', size: 24 },
      paragraph: { spacing: { after: 0, line: 240 } },
    } } },
    numbering: { config: [{
      reference: 'symbol',
      levels: [{
        level: 0, format: LevelFormat.BULLET, suffix: LevelSuffix.TAB, text: '\uf0b7',
        style: { run: { font: 'Symbol' }, paragraph: { indent: { left: 1080, hanging: 360 } } },
      }],
    }] },
    sections: [{
      properties: { page: {
        size: { width: 6000, height: 15840 },
        margin: { left: 1400, right: 1400, top: 1440, bottom: 1440 },
      } },
      children: cases.map(({ head, numberingTab }) => new Paragraph({
        numbering: { reference: 'symbol', level: 0 },
        ...(numberingTab === undefined ? {} : { tabStops: [{ type: TabStopType.NUM, position: numberingTab }] }),
        children: [new TextRun(proof(head))],
      })),
    }],
  }));
  if (!wrapTrailSpaces) return bytes;
  const zip = await JSZip.loadAsync(bytes);
  const settings = await zip.file('word/settings.xml')!.async('string');
  const flag = '<w:wrapTrailSpaces/>';
  if (!settings.includes(flag)) throw Error('The fixture needs the standard flag inside w:compat');
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

async function assertLayout(root: Locator, wrapTrailSpaces: boolean) {
  const items = root.locator('li[data-list-font="Symbol"]');
  await expect(items).toHaveCount(cases.length);
  for (const [index, { head, numberingTab, aligned }] of cases.entries()) {
    const paragraph = items.nth(index).locator(':scope > p');
    const hangs = aligned && !wrapTrailSpaces;
    await expect(paragraph).toHaveCSS('white-space', hangs ? 'pre-wrap' : 'break-spaces');
    expect(await paragraph.textContent()).toBe(proof(head));

    const rendered = await paragraph.evaluate((element, { head, tail }) => {
      const p = element as HTMLElement;
      const item = p.parentElement!;
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      const tops = new Map<string, number>();
      while (walker.nextNode()) {
        const node = walker.currentNode as Text;
        for (const marker of [head, tail]) {
          const offset = node.data.indexOf(marker);
          if (offset < 0) continue;
          const range = document.createRange();
          range.setStart(node, offset);
          range.setEnd(node, offset + 1);
          const rect = range.getBoundingClientRect();
          if (rect.width > 0 && rect.height > 0) tops.set(marker, rect.top);
        }
      }
      if (!tops.has(head) || !tops.has(tail)) throw Error('Missing visible text range');
      const scale = p.getBoundingClientRect().width / p.offsetWidth;
      return {
        lineSeparation: (tops.get(tail)! - tops.get(head)!) / (scale * parseFloat(getComputedStyle(p).lineHeight)),
        marker: getComputedStyle(item, '::marker').content,
        sourceGlyph: item.dataset.listPattern,
        level: item.dataset.listLevel,
        numberingTabs: p.dataset.wordNumberingTabs ?? null,
      };
    }, { head, tail });
    expect(rendered.marker).toBe('"• "');
    expect(rendered.sourceGlyph).toBe('\uf0b7');
    expect(rendered.level).toBe('0');
    expect(rendered.numberingTabs).toBe(numberingTab === undefined ? null : JSON.stringify([numberingTab]));
    if (hangs) expect(rendered.lineSeparation).toBeCloseTo(1, 1);
    else expect(rendered.lineSeparation).toBeGreaterThan(2);
  }
}

test('aligned Symbol bullets hang trailing spaces in editing and print only when Word allows it', async ({ page }) => {
  await page.goto('/');
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  for (const wrapTrailSpaces of [false, true]) {
    const name = wrapTrailSpaces ? 'symbol-wrap-trailing-spaces' : 'symbol-default-trailing-spaces';
    await open(page, name, await fixture(wrapTrailSpaces));
    await assertLayout(editor, wrapTrailSpaces);
    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const print = page.locator('.print-layout');
    await expect(print.locator('.print-page')).toHaveCount(1);
    await assertLayout(print, wrapTrailSpaces);
    await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
  }
});

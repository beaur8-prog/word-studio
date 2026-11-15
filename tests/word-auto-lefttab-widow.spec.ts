import { expect, test } from '@playwright/test';
import { Document, LineRuleType, Packer, Paragraph, Tab, TabStopType, TextRun } from 'docx';

const lines = ['LINE0', 'LINE1', 'LINE2', 'LINE3'];
const cases = [
  { name: 'explicit left tab, one line fits', explicit: true, widow: true, available: 1, retained: 0 },
  { name: 'default left tab, one line fits', explicit: false, widow: true, available: 1, retained: 0 },
  { name: 'widow control disabled, one line fits', explicit: true, widow: false, available: 1, retained: 1 },
  { name: 'widow control enabled, legal two-line split', explicit: true, widow: true, available: 2, retained: 2 },
];

async function fixture(leadTwips: number, explicit: boolean, widow: boolean) {
  const font = 'Century Gothic';
  return Packer.toBuffer(new Document({
    sections: [{
      properties: { page: { size: { width: 6000, height: 6000 }, margin: { top: 750, right: 750, bottom: 750, left: 750 } } },
      children: [
        new Paragraph({
          children: [new TextRun({ text: 'LEAD', font, size: 22 })],
          spacing: { before: 0, after: 0, line: leadTwips, lineRule: LineRuleType.EXACT },
        }),
        new Paragraph({
          children: [new TextRun({
            children: ['6.7', new Tab(), lines.map(marker => marker + 'x'.repeat(22)).join(' ')],
            font,
            size: 22,
          })],
          indent: { left: 720, hanging: 720 },
          spacing: { before: 0, after: 120, line: 276, lineRule: LineRuleType.AUTO },
          widowControl: widow,
          ...(explicit ? { tabStops: [
            { type: TabStopType.LEFT, position: 720 },
            { type: TabStopType.LEFT, position: 2160 },
          ] } : {}),
        }),
      ],
    }],
  }));
}

function measure(root: HTMLElement, printing: boolean) {
  const sheet = root.closest<HTMLElement>('.document-stack');
  const origin = root.getBoundingClientRect();
  const scale = printing ? 1 : origin.width / root.offsetWidth;
  const height = Number(sheet?.dataset.pageHeight);
  const top = Number(sheet?.dataset.margin);
  const bottom = Number(sheet?.dataset.bottomMargin);
  const stride = height + 24;
  const markers: { id: string; page: number; x: number; y: number; inside: boolean }[] = [];
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  while (walker.nextNode()) {
    const text = walker.currentNode as Text;
    if (text.parentElement?.closest('.pagination-gap,.pagination-cell-gap,[data-flow-widget]')) continue;
    for (const hit of text.data.matchAll(/6\.7|LINE[0-3]/g)) {
      const range = document.createRange();
      range.setStart(text, hit.index!);
      range.setEnd(text, hit.index! + hit[0].length);
      const ink = range.getBoundingClientRect();
      const paper = text.parentElement?.closest<HTMLElement>('.print-page');
      const printBody = text.parentElement?.closest('.print-content')?.getBoundingClientRect();
      const editorY = (ink.top - origin.top) / scale;
      const page = paper ? Number(paper.dataset.page) : Math.floor((editorY + .5) / stride) + 1;
      const paperRect = paper?.getBoundingClientRect();
      const y = paperRect ? ink.top - paperRect.top : editorY - (page - 1) * stride;
      markers.push({
        id: hit[0],
        page,
        x: (ink.left - (paperRect?.left ?? origin.left)) / scale,
        y,
        inside: ink.height > 0 && (printBody
          ? ink.top >= printBody.top - 1 && ink.bottom <= printBody.bottom + 1
          : y >= top - 1 && y + ink.height / scale <= height - bottom + 1),
      });
    }
  }
  return { markers, pages: printing ? root.querySelectorAll('.print-page').length : Number(sheet?.dataset.pages) };
}

test('Word automatic line spacing protects widow splits after a manual left tab', async ({ page }) => {
  await page.goto('/');
  const title = page.getByRole('textbox', { name: '문서 제목', exact: true });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  const target = editor.locator('p').filter({ hasText: 'LINE0' });
  const open = async (name: string, bytes: Buffer) => {
    await page.locator('input[type=file]').first().setInputFiles({
      name: name + '.docx',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      buffer: bytes,
    });
    await expect(title).toHaveValue(name);
    await expect(target).toContainText('LINE3');
    await page.evaluate(async () => { await document.fonts.ready; });
    await expect.poll(() => target.evaluate(p => p.style.lineHeight.endsWith('px'))).toBe(true);
  };

  // Measure this four-line paragraph once. Font availability and rounding
  // vary across platforms, so place the boundary halfway between two
  // measured line boxes instead of requiring a fixed 11pt line advance.
  await open('auto-left-tab-calibration', await fixture(300, true, true));
  await expect.poll(async () => (await editor.evaluate(measure, false)).pages).toBe(1);
  const calibration = await editor.evaluate(measure, false);
  expect(calibration.markers.map(marker => marker.id)).toEqual(['6.7', ...lines]);
  const firstField = calibration.markers[1];
  expect(new Set(calibration.markers.slice(1).map(marker => marker.y)).size).toBe(4);
  expect(Math.abs(calibration.markers[0].y - firstField.y)).toBeLessThan(1);
  const geometry = await target.evaluate(paragraph => {
    const editor = paragraph.closest<HTMLElement>('.tiptap')!;
    const sheet = paragraph.closest<HTMLElement>('.document-stack')!;
    return {
      bodyHeight: Number(sheet.dataset.pageHeight) - Number(sheet.dataset.margin) - Number(sheet.dataset.bottomMargin),
      pitch: paragraph.getBoundingClientRect().height / (editor.getBoundingClientRect().width / editor.offsetWidth) / 4,
    };
  });
  expect(geometry.pitch).toBeGreaterThan(0);

  for (const [index, scenario] of cases.entries()) await test.step(scenario.name, async () => {
    const lead = Math.round((geometry.bodyHeight - (scenario.available + .5) * geometry.pitch) * 15);
    await open('auto-left-tab-' + index, await fixture(lead, scenario.explicit, scenario.widow));
    await expect(target).toHaveAttribute('data-word-line-rule', 'auto');
    await expect(target).toHaveAttribute('data-word-line', '1.15');
    await expect(target).toHaveCSS('widows', scenario.widow ? '2' : '1');
    await expect(target).toHaveCSS('orphans', scenario.widow ? '2' : '1');
    await expect(target.locator('[data-docx-tab]')).toHaveCount(1);
    await expect(target.locator('br:not(.ProseMirror-trailingBreak)')).toHaveCount(0);
    await expect.poll(async () => (await editor.evaluate(measure, false)).pages).toBe(2);
    const editing = await editor.evaluate(measure, false);

    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const print = page.locator('.print-layout');
    await expect(print.locator('.print-page')).toHaveCount(2);
    const printing = await print.evaluate(measure, true);
    for (const mode of [editing, printing]) {
      expect(mode.markers.map(marker => marker.id)).toEqual(['6.7', ...lines]);
      const [label, field, ...remaining] = mode.markers;
      expect([field, ...remaining].map(marker => marker.page)).toEqual(
        lines.map((_, index) => index < scenario.retained ? 1 : 2),
      );
      expect(label.page).toBe(field.page);
      expect(Math.abs(label.y - field.y)).toBeLessThan(1);
      expect(Math.abs(field.x - firstField.x)).toBeLessThan(1);
      expect(mode.markers.every(marker => marker.inside)).toBe(true);
    }
    await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
  });
});

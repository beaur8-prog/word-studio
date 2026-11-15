import { expect, test, type Locator, type Page } from '@playwright/test';

const text = 'examplelilix@lili.test';
const paintedText = '[data-word-cell-ink]';

async function dragColumn(page: Page, cell: Locator, offset: number) {
  const box = (await cell.boundingBox())!;
  const x = box.x + box.width - 1;
  const y = box.y + box.height / 2;
  await page.mouse.move(x, y);
  await expect(cell.locator('.column-resize-handle')).toHaveCount(1);
  await page.mouse.down();
  await page.mouse.move(x + offset, y, { steps: 4 });
}

test('an exact fixed cell recalculates fallback spacing after resizing and respects its named face', async ({ page }) => {
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'short-tracked-cell.html',
    mimeType: 'text/html',
    buffer: Buffer.from(`
      <table data-word-table-layout="fixed">
        <tr><td colwidth="81" style="padding:0"><p>Tracking</p></td>
          <td colwidth="120" style="padding:0 1.6px">
            <p data-word-line-rule="exact" data-hwp-cell-line="1 Percent 13.3333 21.3333 13.3 13.3" style="margin:0;font-size:10pt;line-height:10pt;text-align:left">
              <span style="font-family:돋움체;font-size:10pt;letter-spacing:-2.15pt">${text}</span></p></td>
        </tr>
      </table>`.replace(/>\s+</g, '><')),
  });
  const cell = page.locator('.tiptap td').filter({ hasText: text });
  await expect(cell).toHaveText(text);
  const canUseFallback = await page.evaluate(async () => {
    const hasLocalFace = async (name: string) => {
      try {
        await new FontFace('QA cell tracking face', `local(${JSON.stringify(name)})`).load();
        return true;
      } catch { return false; }
    };
    const originals = await Promise.all(['돋움체', 'DotumChe', 'DotumChe-Regular'].map(hasLocalFace));
    return !originals.includes(true) && await hasLocalFace('AppleSDGothicNeo-Regular');
  });

  const paint = cell.locator(paintedText);
  await expect(paint).toHaveCount(canUseFallback ? text.length : 0);
  if (canUseFallback) {
    await dragColumn(page, cell, -47);
    await expect(paint).toHaveCount(0);
    await page.mouse.up();
    await expect(cell).toHaveAttribute('colwidth', '73');
    await expect(paint).toHaveCount(0);

    await dragColumn(page, cell, 47);
    await page.mouse.up();
    await expect(cell).toHaveAttribute('colwidth', '120');
    await expect(paint).toHaveCount(text.length);
    await expect(cell).toHaveText(text);

    // Declare an identical-looking face under the requested name. Matching
    // advance widths alone cannot distinguish it from the installed fallback.
    await page.evaluate(async () => {
      const face = new FontFace('돋움체', 'local("AppleSDGothicNeo-Regular")');
      document.fonts.add(face);
      await face.load();
    });
    await expect(paint).toHaveCount(0);
  }

  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  const printed = page.locator('.print-content td').filter({ hasText: text });
  await expect(printed).toHaveText(text);
  await expect(printed.locator(paintedText)).toHaveCount(0);
});

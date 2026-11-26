import { expect, test } from '@playwright/test';
import { Document, ImageRun, LineRuleType, Packer, Paragraph, TextRun } from 'docx';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aLRsAAAAASUVORK5CYII=', 'base64');
const width = 624, height = 240, x = 96, y = 136;
const marker = 'urn:wordstudio:hwp-body-picture-flow:1';

async function input(protectedFlow: boolean, changedHeight = 0) {
  const image = new ImageRun({
    type: 'png', data: png, transformation: { width, height: height + changedHeight },
    floating: { horizontalPosition: { relative: 'page', offset: x * 9525 },
      verticalPosition: { relative: 'page', offset: y * 9525 }, wrap: { type: 1, side: 'bothSides' }, layoutInCell: false },
  });
  const para = (text: string) => new Paragraph({
    spacing: { before: 0, after: 0, line: 360, lineRule: LineRuleType.EXACT },
    children: [new TextRun({ text, font: 'Arial', size: 24 })],
  });
  const generated = await Packer.toBuffer(new Document({ sections: [{
    properties: { page: { size: { width: 12240, height: 15840 },
      margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    children: [para('Before picture'), new Paragraph({ indent: { left: 450 }, spacing: { before: 0, after: 0, line: 15, lineRule: LineRuleType.EXACT },
      children: [image] }), para(''), para('After blank')],
  }] }));
  if (!protectedFlow) return generated;
  const zip = await JSZip.loadAsync(generated);
  const file = 'word/document.xml', xml = await zip.file(file)!.async('string');
  const proof = JSON.stringify({ anchorY: 120, height, followingEm: 13.333 }).replaceAll('"', '&quot;');
  zip.file(file, xml.replace('<wp:anchor ', `<wp:anchor xmlns:wspf="${marker}" mc:Ignorable="wspf" mc:PreserveAttributes="wspf:clearance" wspf:clearance="${proof}" `));
  return zip.generateAsync({ type: 'nodebuffer' });
}

test('a full-width source picture retains its cleared blank and original drawing through two saves', async ({ page }) => {
  await page.goto('/');
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  const title = page.getByRole('textbox', { name: '문서 제목', exact: true });
  const open = async (bytes: Buffer, name: string) => {
    await page.locator('input[type=file]').first().setInputFiles({ name: name + '.docx',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: bytes });
    await expect(title).toHaveValue(name);
    await expect(editor.locator('p')).toHaveCount(4);
  };
  const printing = async () => {
    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const pageBody = page.locator('.print-layout.preview .print-content');
    await expect(pageBody.locator('p')).toHaveCount(4);
    const geometry = await pageBody.evaluate(root => {
      const img = root.querySelector('img')!, anchor = img.closest('p')!, next = anchor.nextElementSibling!;
      const paper = root.closest('.print-page')!.getBoundingClientRect();
      const box = img.getBoundingClientRect(), before = anchor.getBoundingClientRect(), after = next.getBoundingClientRect();
      return { picture: [box.left - paper.left, box.top - paper.top, box.width, box.height],
        anchor: before.top - paper.top, blank: after.top - paper.top, reserve: parseFloat(anchor.style.minHeight) || 0,
        text: root.textContent!.trim() };
    });
    await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
    return geometry;
  };
  await open(await input(false), 'image-unmeasured');
  const untreated = await printing();
  await open(await input(true, -4), 'image-dimensions-changed');
  const resized = await printing();
  expect(resized.reserve).toBe(untreated.reserve);

  let bytes = await input(true);
  for (let stage = 0; stage < 3; stage++) {
    await open(bytes, `image-clearance-${stage}`);
    const layout = await printing();
    expect(layout.text).toBe(untreated.text);
    expect(untreated.picture[2]).toBeLessThan(width);
    expect(layout.picture).toEqual([x, y, width, height]);
    expect(layout.anchor).toBeCloseTo(120, 1);
    expect(layout.blank).toBeCloseTo(y + height - 40 / 3, 1);
    expect(layout.reserve).toBeCloseTo(y + height - 40 / 3 - 120, 1);
    if (stage === 2) break;
    await page.getByRole('button', { name: '파일', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
    bytes = await readFile((await (await pending).path())!);
    const docx = await JSZip.loadAsync(bytes), xml = await docx.file('word/document.xml')!.async('string');
    expect(xml).toContain('wspf:clearance');
    expect(xml).toContain('<wp:wrapSquare');
    const media = Object.values(docx.files).filter(file => file.name.startsWith('word/media/') && !file.dir);
    expect(media).toHaveLength(1);
    expect(await media[0].async('nodebuffer')).toEqual(png);
  }
});

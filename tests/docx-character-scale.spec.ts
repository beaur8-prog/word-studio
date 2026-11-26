import { expect, test } from '@playwright/test';
import { Document, Packer, Paragraph, TextRun } from 'docx';
import { readFile, writeFile } from 'node:fs/promises';
import JSZip from 'jszip';

const runs = [
  new TextRun({ text: 'Inherited' }),
  new TextRun({ text: ' Reset', scale: 100 }),
  new TextRun({ text: ' Expanded', scale: 120 }),
];
const literals = ['Inherited', ' Reset', ' Expanded'];
const scales = [85, null, 120];

test('Word text scale respects style inheritance and direct resets through two DOCX downloads', async ({ page }, info) => {
  const document = new Document({
    styles: { paragraphStyles: [{ id: 'Compact', name: 'Compact', run: { scale: 85 } }] },
    sections: [{ children: [new Paragraph({ style: 'Compact', children: runs })] }],
  });
  let bytes = await Packer.toBuffer(document);
  await page.goto('/');
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  const title = page.getByRole('textbox', { name: '문서 제목', exact: true });
  const input = page.locator('input[type=file]').first();
  for (let stage = 0; stage <= 2; stage++) {
    const name = `character-scale-${stage}`;
    await input.setInputFiles({ name: name + '.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: bytes });
    await expect(title).toHaveValue(name);
    await expect(editor).toContainText(literals.join(''));
    const rendered = await editor.locator('p').filter({ hasText: 'Inherited' }).evaluate(paragraph => {
      const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
      const result = [];
      while (walker.nextNode()) {
        const text = walker.currentNode;
        const scale = (text.parentElement as HTMLElement).style.getPropertyValue('--word-character-scale');
        result.push([text.nodeValue, scale ? Number(scale) : null]);
      }
      return result;
    });
    expect(rendered).toEqual(literals.map((text, index) => [text, scales[index]]));
    if (stage === 2) break;

    await page.getByRole('button', { name: '파일', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
    bytes = await readFile((await (await pending).path())!);
    await writeFile(info.outputPath(`textbox-save-character-scale-${stage + 1}.docx`), bytes);
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file('word/document.xml')!.async('string');
    const savedRuns = await page.evaluate(xml => {
      const namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      const root = new DOMParser().parseFromString(xml, 'application/xml');
      return [...root.getElementsByTagNameNS(namespace, 'r')].map(run => {
        const scale = run.getElementsByTagNameNS(namespace, 'w')[0]?.getAttributeNS(namespace, 'val');
        return [run.textContent, scale ? Number(scale) : null];
      }).filter(([text]) => text);
    }, xml);
    expect(savedRuns).toEqual(literals.map((text, index) => [text, scales[index]]));
  }
});

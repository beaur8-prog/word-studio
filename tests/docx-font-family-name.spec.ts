import { expect, test } from '@playwright/test';
import { Document, Packer, Paragraph, TextRun } from 'docx';
import { readFile, writeFile } from 'node:fs/promises';
import JSZip from 'jszip';

const numberedFamily = 'Font With Suffix 20';
const literal = '숫자 서체';
const normal = 'ORDINARY';
const stack = 'FALLBACK LIST';
const inert = 'STYLE CONTROL';
const fields = ['ascii', 'hAnsi', 'eastAsia', 'cs'];
const mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

async function fixture() {
  const children = [
    new Paragraph({ children: [new TextRun({ text: literal, font: numberedFamily })] }),
    new Paragraph({ children: [new TextRun({ text: normal, font: 'Arial' })] }),
    new Paragraph({ children: [new TextRun({ text: stack, font: '"Font 20", monospace' })] }),
    new Paragraph({ children: [new TextRun({ text: inert, font: 'Font 20"; color: red' })] }),
  ];
  return Packer.toBuffer(new Document({ sections: [{ children }] }));
}

test('literal numbered font families survive opening and two DOCX downloads', async ({ page }, info) => {
  await page.goto('/');
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  const title = page.getByRole('textbox', { name: '문서 제목', exact: true });
  const input = page.locator('input[type=file]').first();
  const files = [await fixture()];
  const opened = [];
  for (let stage = 0; stage <= 2; stage++) {
    const name = `numbered-font-${stage}`;
    await input.setInputFiles({ name: `${name}.docx`, mimeType, buffer: files[stage] });
    await expect(title).toHaveValue(name);
    await expect(editor).toContainText(literal);
    opened.push(await editor.locator('p').evaluateAll((paragraphs, texts) => texts.map(text => {
      const paragraph = paragraphs.find(node => node.textContent === text)!;
      const glyph = [...paragraph.querySelectorAll('span')].find(node => node.textContent === text) ?? paragraph;
      return { family: getComputedStyle(glyph).fontFamily, color: getComputedStyle(glyph).color };
    }), [literal, normal, stack, inert]));
    if (stage === 2) break;
    await page.getByRole('button', { name: '파일', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
    const download = await pending;
    files.push(await readFile((await download.path())!));
  }
  for (const [stage, file] of files.entries())
    await writeFile(info.outputPath(`textbox-save-numeric-font-${stage}.docx`), file);
  await writeFile(info.outputPath('opened-fonts.json'), JSON.stringify(opened, null, 2));

  expect(opened.map(nodes => nodes[0].family)).toEqual(Array(3).fill('"Font With Suffix 20"'));
  expect(opened.map(nodes => nodes[1].family)).toEqual(Array(3).fill('Arial'));
  expect(opened.map(nodes => nodes[2].family)).toEqual([
    '"Font 20", monospace', '"Font 20"', '"Font 20"',
  ]);
  expect(opened.map(nodes => nodes[3].color)).toEqual(Array(3).fill('rgb(0, 0, 0)'));
  for (const file of files) {
    const zip = await JSZip.loadAsync(file);
    const xml = await zip.file('word/document.xml')!.async('string');
    const fonts = await page.evaluate(({ xml, literal, fields }) => {
      const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      const root = new DOMParser().parseFromString(xml, 'application/xml');
      const run = [...root.getElementsByTagNameNS(w, 'r')].find(node => node.textContent === literal)!;
      const declared = run.getElementsByTagNameNS(w, 'rFonts')[0];
      return fields.map(slot => declared?.getAttributeNS(w, slot));
    }, { xml, literal, fields });
    expect(fonts).toEqual(Array(4).fill(numberedFamily));
  }
});

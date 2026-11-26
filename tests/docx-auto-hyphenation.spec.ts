import { test, expect, type Page } from '@playwright/test';
import { Document, Packer, Paragraph } from 'docx';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

const wordNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const firstSection = 'International cooperation';
const secondSection = 'Document presentation';
const docxMime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

async function source(setting: string) {
  const input = await Packer.toBuffer(new Document({
    sections: [
      { children: [new Paragraph(firstSection)] },
      { children: [new Paragraph(secondSection)] },
    ],
  }));
  if (!setting) return input;
  const zip = await JSZip.loadAsync(input);
  const settings = await zip.file('word/settings.xml')!.async('string');
  if (!settings.includes('<w:compat>')) throw new Error('Expected a settings element after autoHyphenation');
  zip.file('word/settings.xml', settings.replace('<w:compat>', setting + '<w:compat>'));
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, name: string, buffer: Buffer, title = name.replace(/\.[^.]+$/, '')) {
  const mimeType = name.endsWith('.docx') ? docxMime : 'application/json';
  await page.locator('input[type=file]').first().setInputFiles({ name, mimeType, buffer });
  const displayedTitle = page.getByRole('textbox', { name: '문서 제목', exact: true });
  const protect = page.getByRole('dialog', { name: '변경 내용 보관' });
  await expect.poll(async () => await displayedTitle.inputValue() === title || await protect.isVisible()).toBe(true);
  if (await protect.isVisible()) await protect.getByRole('button', { name: '복구본 보관 후 계속' }).click();
  await expect(displayedTitle).toHaveValue(title);
}

async function inspect(page: Page, bytes: Buffer) {
  const zip = await JSZip.loadAsync(bytes);
  const settings = await zip.file('word/settings.xml')!.async('string');
  const document = await zip.file('word/document.xml')!.async('string');
  return page.evaluate(({ settings, document, wordNamespace }) => {
    const parser = new DOMParser();
    const config = parser.parseFromString(settings, 'application/xml');
    const body = parser.parseFromString(document, 'application/xml');
    return {
      hyphenation: Array.from(config.documentElement.children)
        .filter(el => el.namespaceURI === wordNamespace && el.localName === 'autoHyphenation')
        .map(el => el.getAttributeNS(wordNamespace, 'val')),
      sections: body.getElementsByTagNameNS(wordNamespace, 'sectPr').length,
      text: Array.from(body.getElementsByTagNameNS(wordNamespace, 't'), el => el.textContent).join(''),
      softHyphens: body.getElementsByTagNameNS(wordNamespace, 'softHyphen').length,
    };
  }, { settings, document, wordNamespace });
}

const cases = [
  { kind: 'enabled', sourceSetting: '<w:autoHyphenation/>', value: true, exported: [null] },
  { kind: 'disabled', sourceSetting: '<q:autoHyphenation xmlns:q="' + wordNamespace + '" q:val="off"/>', value: false, exported: ['false'] },
  { kind: 'omitted', sourceSetting: '', value: undefined, exported: [] },
];

for (const { kind, sourceSetting, value, exported } of cases) {
  test(`Word auto hyphenation ${kind} survives an editing copy and two DOCX saves`, async ({ page }, info) => {
    await page.goto('/');
    const name = `hyphenation-${kind}`;
    await open(page, name + '.docx', await source(sourceSetting));

    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
    const paragraph = editor.locator('p').filter({ hasText: firstSection });
    await expect(paragraph).toHaveText(firstSection);
    await paragraph.click();
    await page.keyboard.press('End');
    await page.keyboard.insertText(' EDIT');
    await expect(paragraph).toHaveText(firstSection + ' EDIT');

    const fileMenu = page.getByRole('button', { name: '파일', exact: true });
    await fileMenu.click();
    const editingCopy = page.waitForEvent('download');
    await page.locator('.backstage-nav').getByRole('button', { name: '편집본 저장 .wordstudio', exact: true }).click();
    const nativeBytes = await readFile((await (await editingCopy).path())!);
    expect(JSON.parse(nativeBytes.toString()).wordAutoHyphenation).toBe(value);

    await open(page, 'plain.docx', await source(''));
    await expect(paragraph).toHaveText(firstSection);
    await open(page, name + '.wordstudio', nativeBytes);
    await expect(paragraph).toHaveText(firstSection + ' EDIT');

    for (const save of [1, 2]) {
      await fileMenu.click();
      const pending = page.waitForEvent('download');
      await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
      const download = await pending;
      const output = info.outputPath(`textbox-save-auto-hyphenation-${kind}-${save}.docx`);
      await download.saveAs(output);
      const bytes = await readFile(output);
      expect(await inspect(page, bytes)).toEqual({
        hyphenation: exported,
        sections: 2,
        text: firstSection + ' EDIT' + secondSection,
        softHyphens: 0,
      });
      await open(page, `${name}-save-${save}.docx`, bytes);
      await expect(paragraph).toHaveText(firstSection + ' EDIT');
    }
  });
}

test('editing copies only accept a boolean auto hyphenation setting', async ({ page }) => {
  await page.goto('/');
  const accepted = await page.evaluate(async () => {
    const { validateDocument } = await import('/src/document-io.ts');
    const { defaultPage } = await import('/src/types.ts');
    const document = {
      format: 'word-studio', version: 1, title: 'Document', updatedAt: new Date().toISOString(),
      page: defaultPage, comments: [], content: { type: 'doc', content: [{ type: 'paragraph' }] },
    };
    return [undefined, true, false, 'on', 1, null].map(wordAutoHyphenation => {
      try {
        validateDocument({ ...document, wordAutoHyphenation });
        return true;
      } catch {
        return false;
      }
    });
  });
  expect(accepted).toEqual([true, true, true, false, false, false]);
});

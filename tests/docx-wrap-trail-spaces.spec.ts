import { test, expect, type Page } from '@playwright/test';
import { Document, Packer, Paragraph, Tab, TextRun } from 'docx';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

const wordNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const docxMime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const beforeTab = 'Layout  probe\u00a0adjacent\u2003spaced';
const afterTab = 'Field  value';
const afterBreak = 'Second  line';
const lastParagraph = 'End of document';
const sourceParagraphs = [beforeTab + '\t' + afterTab + '\n' + afterBreak, lastParagraph];

const cases = [
  { name: 'enabled', value: true, exported: [null] },
  { name: 'disabled', value: false, exported: ['false'] },
  { name: 'omitted', value: undefined, exported: [] },
] as const;

async function fixture(value: boolean | undefined): Promise<Buffer> {
  const bytes = await Packer.toBuffer(new Document({
    compatibility: value === undefined ? undefined : { wrapTrailSpaces: true },
    sections: [{
      children: [
        new Paragraph({ children: [
          new TextRun({ children: [beforeTab, new Tab(), afterTab] }),
          new TextRun({ break: 1, text: afterBreak }),
        ] }),
        new Paragraph(lastParagraph),
      ],
    }],
  }));
  if (value !== false) return bytes;
  const zip = await JSZip.loadAsync(bytes);
  const settings = await zip.file('word/settings.xml')!.async('string');
  const enabled = '<w:wrapTrailSpaces/>';
  if (!settings.includes(enabled)) throw new Error('Expected the fixture library to supply a wrapTrailSpaces element');
  zip.file('word/settings.xml', settings.replace(enabled, '<q:wrapTrailSpaces xmlns:q="' + wordNamespace + '" q:val="off"/>'));
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, name: string, bytes: Buffer, title = name.replace(/\.[^.]+$/, '')) {
  const mimeType = name.endsWith('.docx') ? docxMime : 'application/json';
  await page.locator('input[type=file]').first().setInputFiles({ name, mimeType, buffer: bytes });
  const displayedTitle = page.getByRole('textbox', { name: '문서 제목', exact: true });
  const protect = page.getByRole('dialog', { name: '변경 내용 보관' });
  await expect.poll(async () => await displayedTitle.inputValue() === title || await protect.isVisible()).toBe(true);
  if (await protect.isVisible()) await protect.getByRole('button', { name: '복구본 보관 후 계속' }).click();
  await expect(displayedTitle).toHaveValue(title);
}

async function inspect(page: Page, bytes: Buffer) {
  const zip = await JSZip.loadAsync(bytes);
  const settings = await zip.file('word/settings.xml')!.async('string');
  const body = await zip.file('word/document.xml')!.async('string');
  return page.evaluate(({ settings, body, wordNamespace }) => {
    const parser = new DOMParser();
    const config = parser.parseFromString(settings, 'application/xml');
    const document = parser.parseFromString(body, 'application/xml');
    const isWord = (element: Element, name: string) => element.namespaceURI === wordNamespace && element.localName === name;
    const compat = Array.from(config.documentElement.children).filter(element => isWord(element, 'compat'));
    const flags = compat.flatMap(element => Array.from(element.children).filter(child => isWord(child, 'wrapTrailSpaces')));
    const bodyElement = document.getElementsByTagNameNS(wordNamespace, 'body')[0];
    const paragraphControls = Array.from(bodyElement.children).filter(element => isWord(element, 'p')).map(paragraph => {
      const runs = Array.from(paragraph.getElementsByTagNameNS(wordNamespace, 'r'));
      return runs.flatMap(run => Array.from(run.children)).filter(element => element.namespaceURI === wordNamespace);
    });
    return {
      settingsRoot: isWord(config.documentElement, 'settings'),
      compatCount: compat.length,
      flagCount: config.getElementsByTagNameNS(wordNamespace, 'wrapTrailSpaces').length,
      flags: flags.map(element => element.getAttributeNS(wordNamespace, 'val')),
      paragraphs: paragraphControls.map(elements => elements.map(element => {
        if (element.localName === 't') return element.textContent ?? '';
        if (element.localName === 'tab') return '\t';
        if (element.localName === 'br') return '\n';
        return '';
      }).join('')),
      tabs: paragraphControls.flat().filter(element => element.localName === 'tab').length,
      breakTypes: paragraphControls.flat().filter(element => element.localName === 'br')
        .map(element => element.getAttributeNS(wordNamespace, 'type') ?? 'textWrapping'),
    };
  }, { settings, body, wordNamespace });
}

for (const { name, value, exported } of cases) {
  test(`Word trailing-space compatibility ${name} survives an editing copy and two DOCX saves`, async ({ page }, info) => {
    await page.goto('/');
    const title = 'wrap-trail-spaces-' + name;
    await open(page, title + '.docx', await fixture(value));

    const fileMenu = page.getByRole('button', { name: '파일', exact: true });
    await fileMenu.click();
    const editingCopy = page.waitForEvent('download');
    await page.locator('.backstage-nav').getByRole('button', { name: '편집본 저장 .wordstudio', exact: true }).click();
    const nativeBytes = await readFile((await (await editingCopy).path())!);
    const native = JSON.parse(nativeBytes.toString());
    if (value === undefined) expect(native).not.toHaveProperty('wordWrapTrailSpaces');
    else expect(native.wordWrapTrailSpaces).toBe(value);

    await open(page, 'other-' + title + '.docx', await fixture(value === true ? undefined : true));
    await open(page, title + '.wordstudio', nativeBytes, title);

    for (const save of [1, 2]) {
      await fileMenu.click();
      const pending = page.waitForEvent('download');
      await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
      const output = info.outputPath(`textbox-save-wrap-trail-spaces-${name}-${save}.docx`);
      await (await pending).saveAs(output);
      const bytes = await readFile(output);
      expect(await inspect(page, bytes)).toEqual({
        settingsRoot: true,
        compatCount: 1,
        flagCount: exported.length,
        flags: [...exported],
        paragraphs: sourceParagraphs,
        tabs: 1,
        breakTypes: ['textWrapping'],
      });
      await open(page, `${title}-save-${save}.docx`, bytes);
    }
  });
}

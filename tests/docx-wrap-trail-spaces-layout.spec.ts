import { test, expect, type Locator, type Page } from '@playwright/test';
import { Document, Packer, Paragraph, Tab, Table, TableCell, TableRow, TextRun } from 'docx';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const docxMime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const head = 'WRAPHEAD';
const tail = 'WRAPTAIL';
const proof = head + ' '.repeat(120) + tail;
const first = 'Keep  two\u00a0NBSP\u2003EM';
const second = 'Tab  field';
const third = 'After  break';
const exactBody = [proof, first + '\t' + second + '\n' + third];
const tableLabel = 'Cell outside the body policy';

type Scenario = { name: string; flag?: boolean; legacy?: boolean; whiteSpace: 'pre-wrap' | 'break-spaces' };
const scenarios: Scenario[] = [
  { name: 'external-omits-flag', whiteSpace: 'pre-wrap' },
  { name: 'external-disabled', flag: false, whiteSpace: 'pre-wrap' },
  { name: 'external-enabled', flag: true, whiteSpace: 'break-spaces' },
  { name: 'prior-studio-omits-flag', legacy: true, whiteSpace: 'break-spaces' },
];

async function fixture({ flag, legacy }: Scenario) {
  const bytes = await Packer.toBuffer(new Document({
    creator: legacy ? 'Word Studio' : 'Independent QA author',
    description: legacy ? 'Created with Word Studio' : 'Independent OOXML layout fixture',
    compatibility: flag === undefined ? undefined : { wrapTrailSpaces: true },
    styles: { default: { document: {
      run: { font: 'Courier New', size: 24 },
      paragraph: { spacing: { after: 0, line: 240 } },
    } } },
    sections: [{
      properties: { page: {
        size: { width: 6000, height: 15840 },
        margin: { left: 1400, right: 1400, top: 1440, bottom: 1440 },
      } },
      children: [
        new Paragraph(proof),
        new Paragraph({ children: [
          new TextRun({ children: [first, new Tab(), second] }),
          new TextRun({ break: 1, text: third }),
        ] }),
        new Table({ rows: [new TableRow({ children: [
          new TableCell({ children: [new Paragraph(tableLabel)] }),
        ] })] }),
      ],
    }],
  }));
  if (flag !== false) return bytes;
  const zip = await JSZip.loadAsync(bytes);
  const settings = await zip.file('word/settings.xml')!.async('string');
  const enabled = '<w:wrapTrailSpaces/>';
  if (!settings.includes(enabled)) throw new Error('The fixture needs the standard flag inside w:compat');
  zip.file('word/settings.xml', settings.replace(enabled, `<q:wrapTrailSpaces xmlns:q="${w}" q:val="off"/>`));
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, name: string, bytes: Buffer) {
  const mimeType = name.endsWith('.docx') ? docxMime : name.endsWith('.wordstudio') ? 'application/json' : 'text/html';
  await page.locator('input[type=file]').first().setInputFiles({ name, mimeType, buffer: bytes });
  const displayedTitle = page.getByRole('textbox', { name: '문서 제목', exact: true });
  const title = name.replace(/\.[^.]+$/, '');
  const protect = page.getByRole('dialog', { name: '변경 내용 보관' });
  await expect.poll(async () => await displayedTitle.inputValue() === title || await protect.isVisible()).toBe(true);
  if (await protect.isVisible()) await protect.getByRole('button', { name: '복구본 보관 후 계속' }).click();
  await expect(displayedTitle).toHaveValue(title);
  await page.evaluate(async () => { await document.fonts.ready; });
}

function evidenceParagraph(root: Locator) {
  return root.locator('p').filter({ hasText: head });
}

async function lineSeparation(paragraph: Locator) {
  return paragraph.evaluate((element, { head, tail }) => {
    const findTop = (marker: string) => {
      const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) {
        const node = walker.currentNode as Text;
        const index = node.data.indexOf(marker);
        if (index < 0) continue;
        const range = document.createRange();
        range.setStart(node, index);
        range.setEnd(node, index + 1);
        const rect = range.getBoundingClientRect();
        if (rect.width > 0 && rect.height > 0) return rect.top;
      }
      throw new Error('Could not measure the visible marker ' + marker);
    };
    const paragraph = element as HTMLElement;
    const scale = paragraph.getBoundingClientRect().width / paragraph.offsetWidth;
    const lineHeight = parseFloat(getComputedStyle(element).lineHeight);
    return (findTop(tail) - findTop(head)) / (scale * lineHeight);
  }, { head, tail });
}

async function assertLayout(root: Locator, scenario: Scenario) {
  const paragraph = evidenceParagraph(root);
  await expect(paragraph).toHaveCount(1);
  await expect(paragraph).toHaveCSS('white-space', scenario.whiteSpace);
  expect(await paragraph.textContent()).toBe(proof);
  // The markers are ordinary printable characters. Their measured line
  // separation is independent of font rounding, editor zoom and preview zoom.
  const separation = await lineSeparation(paragraph);
  if (scenario.whiteSpace === 'pre-wrap') expect(separation).toBeCloseTo(1, 1);
  else expect(separation).toBeGreaterThan(2);
  const cell = root.locator('td p').filter({ hasText: tableLabel });
  await expect(cell).toHaveCount(1);
  await expect(cell).toHaveCSS('white-space', 'break-spaces');
}

async function preview(page: Page) {
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  const root = page.locator('.print-layout');
  await expect(root.locator('.print-page')).toHaveCount(1);
  return root;
}

async function inspectDocx(page: Page, bytes: Buffer) {
  const zip = await JSZip.loadAsync(bytes);
  const settings = await zip.file('word/settings.xml')!.async('string');
  const document = await zip.file('word/document.xml')!.async('string');
  return page.evaluate(({ settings, document, w }) => {
    const parser = new DOMParser();
    const settingsXml = parser.parseFromString(settings, 'application/xml');
    const bodyXml = parser.parseFromString(document, 'application/xml');
    const isWord = (el: Element, name: string) => el.namespaceURI === w && el.localName === name;
    const compat = Array.from(settingsXml.documentElement.children).filter(el => isWord(el, 'compat'));
    const flags = compat.flatMap(el => Array.from(el.children).filter(child => isWord(child, 'wrapTrailSpaces')));
    const body = bodyXml.getElementsByTagNameNS(w, 'body')[0];
    const runs = Array.from(body.children).filter(el => isWord(el, 'p')).map(paragraph =>
      Array.from(paragraph.getElementsByTagNameNS(w, 'r')).flatMap(run =>
        Array.from(run.children).filter(el => el.namespaceURI === w)));
    return {
      compatCount: compat.length,
      flagCount: settingsXml.getElementsByTagNameNS(w, 'wrapTrailSpaces').length,
      flags: flags.map(el => el.getAttributeNS(w, 'val')),
      paragraphs: runs.map(elements => elements.map(el => {
        if (el.localName === 't') return el.textContent ?? '';
        if (el.localName === 'tab') return '\t';
        if (el.localName === 'br') return '\n';
        return '';
      }).join('')).filter(text => text.length > 0),
      tabs: runs.flat().filter(el => el.localName === 'tab').length,
      breaks: runs.flat().filter(el => el.localName === 'br')
        .map(el => el.getAttributeNS(w, 'type') ?? 'textWrapping'),
    };
  }, { settings, document, w });
}

for (const scenario of scenarios) {
  test(`Word body trailing whitespace layout survives two DOCX saves (${scenario.name})`, async ({ page }, info) => {
    await page.goto('/');
    const title = 'trail-space-layout-' + scenario.name;
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
    await open(page, title + '.docx', await fixture(scenario));

    if (scenario.flag === undefined && !scenario.legacy) {
      await page.getByRole('button', { name: '파일', exact: true }).click();
      const pending = page.waitForEvent('download');
      await page.locator('.backstage-nav').getByRole('button', { name: '편집본 저장 .wordstudio', exact: true }).click();
      const editingCopy = await readFile((await (await pending).path())!);
      const native = JSON.parse(editingCopy.toString());
      expect(native.wordDefaultHangingSpaces).toBe(true);
      expect(native).not.toHaveProperty('wordWrapTrailSpaces');
      await open(page, 'native-control.docx', await fixture(scenarios[3]));
      await expect(evidenceParagraph(editor)).toHaveCSS('white-space', 'break-spaces');
      await open(page, title + '.wordstudio', editingCopy);
    }

    for (const stage of [0, 1, 2]) {
      await assertLayout(editor, scenario);
      const print = await preview(page);
      await assertLayout(print, scenario);
      await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
      if (stage === 2) break;

      await page.getByRole('button', { name: '파일', exact: true }).click();
      const pending = page.waitForEvent('download');
      await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
      const output = info.outputPath(`textbox-save-wrap-layout-${scenario.name}-${stage + 1}.docx`);
      await (await pending).saveAs(output);
      const bytes = await readFile(output);
      const flags = scenario.flag === undefined ? [] : [scenario.flag ? null : 'false'];
      expect(await inspectDocx(page, bytes)).toEqual({
        compatCount: 1,
        flagCount: flags.length,
        flags,
        paragraphs: exactBody,
        tabs: 1,
        breaks: ['textWrapping'],
      });
      await open(page, `${title}-save-${stage + 1}.docx`, bytes);
    }
  });
}

test('opening native HTML after a Word DOCX keeps the native trailing-space behavior', async ({ page }) => {
  await page.goto('/');
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  await open(page, 'external-before-native.docx', await fixture(scenarios[0]));
  await expect(evidenceParagraph(editor)).toHaveCSS('white-space', 'pre-wrap');

  await open(page, 'native-whitespace.html', Buffer.from(`<p>${proof}</p>`));
  await expect(evidenceParagraph(editor)).toHaveCSS('white-space', 'break-spaces');
  const print = await preview(page);
  await expect(evidenceParagraph(print)).toHaveCSS('white-space', 'break-spaces');
});

test('Android reading, editing and printing carry the Word whitespace layout', async ({ browser, baseURL }) => {
  const context = await browser.newContext({
    baseURL, viewport: { width: 412, height: 840 }, hasTouch: true,
    userAgent: 'Mozilla/5.0 (Linux; Android 14; Pixel 8) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Mobile Safari/537.36',
  });
  try {
    const page = await context.newPage();
    await page.addInitScript(() => {
      Object.assign(window, { printedWordDocument: '', wordAndroid: {
        printDocument: async ({ data }: { data: Uint8Array }) => {
          Object.assign(window, { printedWordDocument: new TextDecoder().decode(data) });
          return 'submitted';
        },
      } });
    });
    await page.goto('/');
    await page.locator('.android-shell > input[type=file]').setInputFiles({
      name: 'word-whitespace.docx', mimeType: docxMime, buffer: await fixture(scenarios[0]),
    });
    const reading = evidenceParagraph(page.locator('.android-paper-layer'));
    await expect(reading).toHaveCSS('white-space', 'pre-wrap');
    await page.getByRole('button', { name: '문서 메뉴', exact: true }).click();
    await page.getByRole('button', { name: '인쇄·PDF 저장', exact: true }).click();
    await expect.poll(() => page.evaluate(() => {
      const html = (window as Window & { printedWordDocument?: string }).printedWordDocument ?? '';
      const document = new DOMParser().parseFromString(html, 'text/html');
      return document.querySelector('.android-print-source p[data-word-hanging-spaces]')?.textContent;
    })).toBe(proof);
    await page.locator('.android-topbar').getByRole('button', { name: '편집 모드', exact: true }).click();
    const editor = page.getByRole('textbox', { name: 'Android 문서 편집 영역', exact: true });
    await expect(evidenceParagraph(editor)).toHaveCSS('white-space', 'pre-wrap');
  } finally {
    await context.close();
  }
});

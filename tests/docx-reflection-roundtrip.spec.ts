import { expect, test, type Page } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { groupedDrawing, headerGroupedDrawing } from './helpers/grouped-drawing';
import { reflectionAttributes, reflectionFixture, reflectionLabels, word14Namespace, wordNamespace } from './word-reflection-fixture';

const runMark = 'span[data-word-reflection]';
const paragraphMark = '[data-word-paragraph-reflection]';
const { direct, plainSibling, effectOnly, unrelated, unsafe } = reflectionLabels;

async function openReflectionDocx(page: Page, name: string, buffer: Buffer) {
  await page.locator('input[type=file]').first().setInputFiles({
    name: name + '.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer,
  });
  const title = page.getByRole('textbox', { name: '문서 제목', exact: true });
  const recovery = page.getByRole('dialog', { name: '변경 내용 보관' });
  await expect.poll(async () => await title.inputValue() === name || await recovery.isVisible()).toBe(true);
  if (await recovery.isVisible()) await recovery.getByRole('button', { name: '복구본 보관 후 계속' }).click();
  await expect(title).toHaveValue(name);
  await expect(page.locator('.busy-indicator')).toBeHidden();
}

async function downloadReflectionDocx(page: Page, path: string) {
  await page.getByRole('button', { name: '파일', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
  await (await pending).saveAs(path);
  return readFile(path);
}

async function inspectReflectionXml(page: Page, bytes: Buffer) {
  const zip = await JSZip.loadAsync(bytes);
  const xml = await zip.file('word/document.xml')!.async('string');
  return page.evaluate(({ xml, word, word14 }) => {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    if (document.querySelector('parsererror')) throw new Error('Invalid document.xml');
    const children = (el: Element | undefined, name: string) =>
      Array.from(el?.children ?? []).filter(child => child.namespaceURI === word && child.localName === name);
    const reflect = (props: Element | undefined) => Array.from(props?.children ?? [])
      .filter(child => child.namespaceURI === word14 && child.localName === 'reflection')
      .map(element => ({
        attributes: Object.fromEntries(Array.from(element.attributes)
          .filter(attr => attr.namespaceURI !== 'http://www.w3.org/2000/xmlns/')
          .sort((a, b) => a.localName.localeCompare(b.localName)).map(attr => [attr.localName, attr.value])),
        namespaced: Array.from(element.attributes).every(attr => attr.namespaceURI === word14 ||
          attr.namespaceURI === 'http://www.w3.org/2000/xmlns/'),
        children: element.children.length,
      }));
    const body = document.getElementsByTagNameNS(word, 'body')[0];
    const paragraphs = children(body, 'p').map(p => {
      const paragraphEffect = reflect(children(children(p, 'pPr')[0], 'rPr')[0]);
      const runs: { text: string; effects: ReturnType<typeof reflect> }[] = [];
      for (const r of children(p, 'r')) {
        const text = children(r, 't').map(t => t.textContent ?? '').join('');
        const effects = reflect(children(r, 'rPr')[0]);
        if (!text && !effects.length) continue;
        const prior = runs.at(-1);
        // Shape-preserving run splitting is allowed. The first plain sibling
        // must still be a distinct unreflected segment with its leading space.
        if (prior && JSON.stringify(prior.effects) === JSON.stringify(effects)) prior.text += text;
        else runs.push({ text, effects });
      }
      return { paragraphEffect, runs };
    });
    return {
      paragraphs,
      reflectionCount: Array.from(document.getElementsByTagName('*')).filter(el => el.localName === 'reflection').length,
    };
  }, { xml, word: wordNamespace, word14: word14Namespace });
}

test('direct Word reflection and paragraph-mark reflection survive two real DOCX saves', async ({ page }, info) => {
  const leaf = (attributes: Readonly<Record<string, string>>) => [{ attributes, namespaced: true, children: 0 }];
  const expected = { reflectionCount: 4, paragraphs: [
    { paragraphEffect: leaf(reflectionAttributes.paragraphMark), runs: [
      { text: direct, effects: leaf(reflectionAttributes.directRun) }, { text: plainSibling, effects: [] },
    ] },
    { paragraphEffect: [], runs: [{ text: effectOnly, effects: leaf(reflectionAttributes.effectOnlyRun) }] },
    { paragraphEffect: leaf(reflectionAttributes.emptyTitle), runs: [] },
    { paragraphEffect: [], runs: [{ text: unrelated, effects: [] }] },
  ] };
  await page.goto('/');
  let bytes = await reflectionFixture();
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  for (const stage of [0, 1, 2]) {
    const normalized = await inspectReflectionXml(page, bytes);
    expect(normalized).toEqual(expected);
    await openReflectionDocx(page, 'reflection-stage-' + stage, bytes);
    await expect(editor.locator(runMark)).toHaveText([direct, effectOnly]);
    await expect(editor.locator(paragraphMark)).toHaveText([direct + plainSibling, '']);
    await expect(editor.locator(paragraphMark).nth(1)).toBeEmpty();
    const blocks = editor.locator(':scope > :is(p,h1,h2,h3,h4,h5,h6)');
    await expect(blocks).toHaveText([direct + plainSibling, effectOnly, '', unrelated]);
    if (stage < 2) bytes = await downloadReflectionDocx(page, info.outputPath(`textbox-save-reflection-ui-${stage + 1}.docx`));
  }
});

test('unsupported reflection metadata is omitted without losing either plain paragraph', async ({ page }, info) => {
  test.setTimeout(90_000);
  for (const scenario of ['wrong-namespace', 'unknown-child', 'numeric-excess'] as const) await test.step(scenario, async () => {
    await page.goto('/');
    await openReflectionDocx(page, 'reflection-' + scenario, await reflectionFixture(scenario));
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
    await expect(editor.locator(':scope > p')).toHaveText([unsafe, unrelated]);
    await expect(editor.locator(`[data-word-reflection],${paragraphMark}`)).toHaveCount(0);
    const bytes = await downloadReflectionDocx(page, info.outputPath(`textbox-save-reflection-${scenario}.docx`));
    expect(await inspectReflectionXml(page, bytes)).toEqual({ reflectionCount: 0, paragraphs: [unsafe, unrelated]
      .map(text => ({ paragraphEffect: [], runs: [{ text, effects: [] }] })) });
  });
});

test('reflection remains inside a formatted link with a comment anchor', async ({ page }, info) => {
  const zip = await JSZip.loadAsync(await reflectionFixture());
  const initialXml = await zip.file('word/document.xml')!.async('string');
  const relationships = await zip.file('word/_rels/document.xml.rels')!.async('string');
  const types = await zip.file('[Content_Types].xml')!.async('string');
  const officeRelationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const link = 'https://example.invalid/qa';
  const text = 'REFLECTED-LINK';
  const following = 'UNREFLECTED-FOLLOWING';
  const visible = `<w:p><w:commentRangeStart w:id="0"/><w:hyperlink r:id="rLink" xmlns:r="${officeRelationships}">
    <w:r><w:rPr><w:b/><w14:reflection w14:blurRad="1420" w14:sy="-75000"/></w:rPr><w:t>${text}</w:t></w:r>
    </w:hyperlink><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>
    </w:p><w:p><w:r><w:t>${following}</w:t></w:r></w:p>`;
  zip.file('word/document.xml', initialXml.replace(/(<w:body>)[\s\S]*?(<w:sectPr>)/, '$1' + visible + '$2'));
  zip.file('word/_rels/document.xml.rels', relationships.replace('</Relationships>',
    `<Relationship Id="rLink" Type="${officeRelationships}/hyperlink" Target="${link}" TargetMode="External"/>
     <Relationship Id="rComments" Type="${officeRelationships}/comments" Target="comments.xml"/></Relationships>`));
  zip.file('[Content_Types].xml', types.replace('</Types>',
    '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>'));
  zip.file('word/comments.xml', `<w:comments xmlns:w="${wordNamespace}"><w:comment w:author="QA" w:id="0">
    <w:p><w:r><w:t>Keep the comment anchor</w:t></w:r></w:p></w:comment></w:comments>`);
  await page.goto('/');
  await openReflectionDocx(page, 'reflection-link-comment', await zip.generateAsync({ type: 'nodebuffer' }));
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  const marked = editor.getByRole('link', { name: text, exact: true });
  await expect(marked).toHaveAttribute('href', link);
  await expect(editor.locator(runMark)).toHaveText([text]);
  await expect(editor.locator('[data-comment-ids]')).toHaveText([text]);
  await expect(editor.locator(':scope > p').last()).toHaveText(following);
  const saved = await downloadReflectionDocx(page, info.outputPath('textbox-save-reflection-link-comment.docx'));
  const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
  expect(await page.evaluate(({ xml, word, word14 }) => {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const leaf = doc.getElementsByTagNameNS(word14, 'reflection');
    const run = leaf[0]?.parentElement?.parentElement;
    return { effectCount: leaf.length, grandparent: run?.parentElement?.localName,
      linkedText: run?.getElementsByTagNameNS(word, 't')[0]?.textContent,
      isBold: Boolean(run?.getElementsByTagNameNS(word, 'b').length),
      comments: doc.getElementsByTagNameNS(word, 'commentRangeStart').length };
  }, { xml, word: wordNamespace, word14: word14Namespace })).toEqual({
    effectCount: 1, grandparent: 'hyperlink', linkedText: text, isBold: true, comments: 1,
  });
});

test('Word reflection survives both sides of a native tab without styling the tab', async ({ page }, info) => {
  const afterTab = ' REFLECTED-AFTER-TAB';
  const before = `<w:t xml:space="preserve">${direct}</w:t>`;
  const fixture = await JSZip.loadAsync(await reflectionFixture());
  const xml = await fixture.file('word/document.xml')!.async('string');
  expect(xml.split(before)).toHaveLength(2);
  fixture.file('word/document.xml', xml.replace(before, before + '<w:tab/>' +
    `<w:t xml:space="preserve">${afterTab}</w:t>`));
  let bytes = await fixture.generateAsync({ type: 'nodebuffer' });
  await page.goto('/');
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  for (const stage of [0, 1, 2]) {
    await openReflectionDocx(page, 'reflection-tab-stage-' + stage, bytes);
    await expect(editor.locator(runMark)).toHaveText([direct, afterTab, effectOnly]);
    await expect(editor.locator('[data-docx-tab]')).toHaveCount(1);
    await expect(editor.locator(`${runMark} [data-docx-tab]`)).toHaveCount(0);
    if (stage === 2) break;
    bytes = await downloadReflectionDocx(page, info.outputPath(`textbox-save-reflection-tab-${stage + 1}.docx`));
    const saved = await JSZip.loadAsync(bytes);
    const savedXml = await saved.file('word/document.xml')!.async('string');
    expect(await page.evaluate(({ xml, w, w14 }) => {
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      const first = doc.getElementsByTagNameNS(w, 'body')[0].getElementsByTagNameNS(w, 'p')[0];
      const tabs = [...doc.getElementsByTagNameNS(w, 'tab')].filter(tab => tab.parentElement?.localName === 'r');
      const tokens = [...first.getElementsByTagNameNS(w, 'r')].flatMap(run => {
        const effect = run.getElementsByTagNameNS(w14, 'reflection')[0];
        const attributes = effect ? Object.fromEntries([...effect.attributes]
          .filter(a => a.namespaceURI === w14).map(a => [a.localName, a.value])) : null;
        return [...run.children].filter(child => child.namespaceURI === w && ['t', 'tab'].includes(child.localName))
          .map(child => ({ text: child.localName === 'tab' ? '\t' : child.textContent, attributes,
            tabAttributeCount: child.localName === 'tab' ? child.attributes.length : null }));
      });
      return { runTabs: tabs.length, tokens };
    }, { xml: savedXml, w: wordNamespace, w14: word14Namespace })).toEqual({
      runTabs: 1, tokens: [
        { text: direct, attributes: reflectionAttributes.directRun, tabAttributeCount: null },
        { text: '\t', attributes: null, tabAttributeCount: 0 },
        { text: afterTab, attributes: reflectionAttributes.directRun, tabAttributeCount: null },
        { text: plainSibling, attributes: null, tabAttributeCount: null },
      ],
    });
  }
});

test('a textbox retains its reflection only within the main document scope', async ({ page }) => {
  await page.goto('/');
  for (const story of ['body', 'header'] as const) {
    const source = story === 'body' ? await groupedDrawing('mixed') : await headerGroupedDrawing();
    const zip = await JSZip.loadAsync(source);
    const part = story === 'body' ? 'word/document.xml' : 'word/header1.xml';
    const xml = await zip.file(part)!.async('string');
    zip.file(part, xml.replace('mc:Ignorable="wpg wps"',
      'xmlns:w14="' + word14Namespace + '" mc:Ignorable="wpg wps w14"')
      .replace('<w:txbxContent><w:p>', '<w:txbxContent><w:p><w:pPr><w:rPr><w14:reflection w14:dist="1524"/></w:rPr></w:pPr>')
      .replace('<w:color w:val="FFFFFF"/></w:rPr><w:t>EDITABLE GROUP STORY',
        '<w:color w:val="FFFFFF"/><w14:reflection w14:dist="1270"/></w:rPr><w:t>EDITABLE GROUP STORY'));
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    const imported = await page.evaluate(async ({ bytes, story }) => {
      const { importFile } = await import('/src/document-io.ts');
      const opened = await importFile(new File([new Uint8Array(bytes)], 'reflected-box.docx'));
      const template = document.createElement('template');
      template.innerHTML = (story === 'body' ? opened.html : opened.document.page.headerHtml) ?? '';
      const box = template.content.querySelector('[data-document-textbox]');
      return { text: box?.textContent, effects: box?.querySelectorAll('[data-word-reflection],[data-word-paragraph-reflection]').length };
    }, { bytes: [...bytes], story });
    expect(imported).toEqual({ text: 'EDITABLE GROUP STORY', effects: story === 'body' ? 2 : 0 });
  }
});

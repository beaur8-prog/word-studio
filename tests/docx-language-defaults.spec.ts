import { test, expect } from '@playwright/test';
import { Document, Packer, Paragraph } from 'docx';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

const cases = [
  { name: 'unified', language: { value: 'en-US', eastAsia: 'en-US', bidirectional: 'ar-SA' }, rendered: 'en-US' },
  { name: 'mixed', language: { value: 'en-US', eastAsia: 'ja-JP', bidirectional: 'ar-SA' }, rendered: undefined },
  { name: 'default', language: undefined, rendered: undefined },
];

for (const { name, language, rendered } of cases) test('DOCX default text language stays with its ' + name + ' document', async ({ page }, info) => {
  await page.goto('/');
  const source = await Packer.toBuffer(new Document({
    styles: { default: { document: { run: { language } } } },
    sections: [{ children: [new Paragraph('English text with 日本語')] }],
  }));
  const input = page.locator('input[type=file]').first();
  await input.setInputFiles({ name: name + '.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: source });
  const stack = page.locator('.document-stack'), editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  const title = page.getByRole('textbox', { name: '문서 제목', exact: true });
  const show = async stage => {
    await expect(title).toHaveValue(stage);
    await expect(editor).toHaveText('English text with 日本語');
    await expect(page.locator('html')).toHaveAttribute('lang', 'ko');
    if (rendered) await expect(stack).toHaveAttribute('lang', rendered);
    else await expect(stack).not.toHaveAttribute('lang');
    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const preview = page.locator('.print-layout.preview');
    await expect(preview.locator('.print-content')).toHaveText('English text with 日本語');
    if (rendered) await expect(preview).toHaveAttribute('lang', rendered);
    else await expect(preview).not.toHaveAttribute('lang');
    await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
  };
  await show(name);
  for (let i = 1; i <= 2; i++) {
    await page.getByRole('button', { name: '파일', exact: true }).click();
    const download = page.waitForEvent('download');
    await page.getByRole('button', { name: 'Word 문서로 내보내기 DOCX', exact: true }).click();
    const saved = info.outputPath('textbox-save-language-' + name + '-' + i + '.docx');
    await (await download).saveAs(saved);
    const archive = await JSZip.loadAsync(await readFile(saved));
    const styles = await archive.file('word/styles.xml')!.async('string');
    const defaults = styles.match(/<w:docDefaults>[\s\S]*?<\/w:docDefaults>/)?.[0];
    expect(defaults).toBeDefined();
    const lang = defaults!.match(/<w:lang\b[^>]*\/>/)?.[0];
    if (language) {
      expect(lang).toContain('w:val="' + language.value + '"');
      expect(lang).toContain('w:eastAsia="' + language.eastAsia + '"');
      expect(lang).toContain('w:bidi="' + language.bidirectional + '"');
    } else expect(lang).toBeUndefined();
    await input.setInputFiles(saved);
    await show('textbox-save-language-' + name + '-' + i);
  }
});

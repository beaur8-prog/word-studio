import { test, expect } from '@playwright/test';
import { mkdir, writeFile } from 'node:fs/promises';

test('missing native Korean serif uses bundled glyph advances without changing the requested face or point size', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).waitFor();
  const result = await page.evaluate(async () => {
    let installed = false;
    try { await new FontFace('Native requested face', 'local("휴먼명조")').load(); installed = true; } catch {}
    const root = document.createElement('div'); root.className = 'tiptap'; document.body.append(root);
    try {
      const span = document.createElement('span'); span.style.cssText = 'font-family:휴먼명조;font-size:15pt;letter-spacing:0;white-space:pre'; span.textContent = '가나다라마바사아자차'; root.append(span);
      await document.fonts.load('15pt "Noto Serif KR"', span.textContent);
      await document.fonts.ready;
      // Exercise the portable no-native-face branch, including on a Mac or
      // Windows runner where the optional verified native alias is present.
      if (!installed) for (const face of [...document.fonts])
        if (face.family.replace(/["']/g, '') === 'Studio Korean Serif') document.fonts.delete(face);
      const range = document.createRange(); range.selectNodeContents(span);
      const reference = document.createElement('span'); reference.style.cssText = 'font-family:"Noto Serif KR";font-size:15pt;letter-spacing:0;white-space:pre'; reference.textContent = span.textContent; root.append(reference);
      const referenceRange = document.createRange(); referenceRange.selectNodeContents(reference);
      const control = document.createElement('span'); control.style.fontFamily = 'BatangChe'; root.append(control);
      const authored = document.createElement('span'); authored.style.fontFamily = '휴먼명조, Arial'; root.append(authored);
      return { installed, width: range.getBoundingClientRect().width, referenceWidth: referenceRange.getBoundingClientRect().width, size: getComputedStyle(span).fontSize, source: span.style.fontFamily, control: getComputedStyle(control).fontFamily, authored: getComputedStyle(authored).fontFamily };
    } finally { root.remove(); }
  });
  if (!result.installed) expect(result.width).toBeCloseTo(result.referenceWidth, 2);
  expect(result.size).toBe('20px'); expect(result.source).toBe('휴먼명조');
  expect(result.control).toBe('BatangChe'); expect(result.authored).toBe('휴먼명조, Arial');
});

test('Korean serif fallback survives editing, undo, redo, cold print and two DOCX saves with original font names', async ({ page }) => {
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({ name: 'serif.html', mimeType: 'text/html', buffer: Buffer.from('<p style="font-family:휴먼명조;font-size:15pt;line-height:20pt">정부 문서의 ‘본문’과 “표” 내용·123</p>') });
  const p = page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).locator('p').first();
  await p.click();
  const documentEnd = await page.evaluate(() => navigator.platform.startsWith('Mac') ? 'Meta+ArrowDown' : 'Control+End');
  await page.keyboard.press(documentEnd); await page.keyboard.type(' EDIT');
  await page.keyboard.press('ControlOrMeta+z'); await expect(p).not.toContainText('EDIT');
  await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(p).toContainText('EDIT');
  const states = await page.evaluate(async () => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    let model: any = { format: 'word-studio', version: 1, id: 'serif', title: 'Serif', updatedAt: new Date().toISOString(), page: defaultPage, comments: [], content: parseDocumentHtml(document.querySelector('.tiptap')!.innerHTML) };
    const rows = [];
    for (let save = 0; save < 3; save++) {
      const html = generateHTML(model.content, editorExtensions), layout = await createPrintLayout(html, model.page);
      try { const el = layout.element.querySelector('.print-content p')!; rows.push({ text: el.textContent, pages: layout.pages, family: getComputedStyle(el).fontFamily, source: html, xml: '', bytes: [] as number[] }); } finally { layout.dispose(); }
      const blob = await exportDocx(model), zip = await JSZip.loadAsync(await blob.arrayBuffer()); rows.at(-1)!.xml = await zip.file('word/document.xml')!.async('string'); rows.at(-1)!.bytes = [...new Uint8Array(await blob.arrayBuffer())];
      const loaded = await importFile(new File([blob], 'serif.docx')); model = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
    }
    return rows;
  });
  await mkdir('test-results/korean-font-schema', { recursive: true });
  for (const [index, row] of states.entries()) {
    await writeFile(`test-results/korean-font-schema/textbox-save-korean-serif-${index}.docx`, Buffer.from(row.bytes));
    expect(row.text).toBe('정부 문서의 ‘본문’과 “표” 내용·123 EDIT'); expect(row.pages).toBe(1);
    expect(row.family).toContain('휴먼명조'); expect(row.source).not.toContain('Noto Serif');
    expect(row.xml).toContain('휴먼명조'); expect(row.xml).not.toMatch(/Noto Serif|AppleMyungjo|Times New Roman|Studio Korean Serif/);
  }
});


test('native Korean serif fallback keeps real Batang and proportional Mac punctuation without changing Hangul advances', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).waitFor();
  const result = await page.evaluate(async () => {
    const available = async (name: string) => {
      try { await new FontFace('Native serif availability', `local(${JSON.stringify(name)})`).load(); return true; }
      catch { return false; }
    };
    const batang = await available('Batang') || await available('바탕');
    const apple = await available('AppleMyungjo'), times = await available('Times New Roman');
    await document.fonts.load('20px "Studio Korean Serif"', '가나다‘’“”·');
    const faces = [...document.fonts].filter(face => face.family.replace(/["']/g, '') === 'Studio Korean Serif');
    const canvas = document.createElement('canvas'), ctx = canvas.getContext('2d')!;
    const width = (family: string, text: string) => { ctx.font = `20px "${family}"`; return ctx.measureText(text).width; };
    return { batang, apple, times, ranges: faces.map(f => f.unicodeRange),
      hangul: width('Studio Korean Serif', '가나다'), appleHangul: width('AppleMyungjo', '가나다'),
      punctuation: width('Studio Korean Serif', '‘’“”·'), timesPunctuation: width('Times New Roman', '‘’“”·'),
      batangPunctuation: width('Batang', '‘’“”·') };
  });
  if (result.batang) {
    expect(result.ranges).toEqual(['U+0-10FFFF']);
    expect(result.punctuation).toBeCloseTo(result.batangPunctuation, 2);
  } else if (result.apple && result.times) {
    expect(result.ranges).toHaveLength(2);
    expect(result.hangul).toBeCloseTo(result.appleHangul, 2);
    expect(result.punctuation).toBeCloseTo(result.timesPunctuation, 2);
    expect(result.punctuation).toBeLessThan(60); // Full-em Apple punctuation is100px.
  } else if (!result.apple) expect(result.ranges).toHaveLength(0);
});

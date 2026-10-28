import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
const createFixture = createRequire(import.meta.url)('./hwp-layout-fixture.cjs');
let directory: string, files: Record<string, string>;
test.beforeAll(async () => {
  await fs.mkdir('.cache', { recursive: true });
  directory = await fs.mkdtemp(path.resolve('.cache/hwp-layout-'));
  files = await createFixture(directory);
});
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

for (const format of ['hwp', 'hwpx']) test(`${format}: custom paper, two columns, breaks, equation appearance and background image survive conversion and DOCX reopening`, async ({ page }) => {
  await page.goto('/'); await page.locator('input[type=file]').first().setInputFiles(files[format]);
  const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
  await expect(reader.getByTestId('hwp-page')).toBeVisible();
  await reader.getByRole('button', { name: 'Word 화면에서 편집', exact: true }).click();
  const notice = reader.getByRole('region', { name: '한글 편집 변환 안내' });
  await expect(notice).toBeVisible();
  await expect(notice).toContainText('수식은 원래 모양');
  await notice.getByRole('button', { name: '편집본 만들기', exact: true }).click();
  await expect(reader).not.toBeVisible();
  const editor = page.locator('.tiptap');
  await expect(editor).toContainText('FIRST 한글'); await expect(editor).toContainText('SECOND 둘째 단'); await expect(editor).toContainText('THIRD 다음 쪽');
  await expect(editor.locator('img[alt="수식: a over b"]')).toHaveCount(1);
  await expect(editor.locator('img[data-image-floating]')).toHaveCount(1);
  await expect(editor.locator('[data-column-break]')).toHaveCount(1);
  await expect(editor.locator('[data-page-break]')).toHaveCount(1);
  const saved = page.waitForEvent('download');
  await page.getByRole('button', { name: '파일', exact: true }).click();
  await page.getByRole('button', { name: 'Word 문서로 내보내기 DOCX', exact: true }).click();
  const file = path.join(directory, `converted-${format}.docx`); await (await saved).saveAs(file);
  const result = await page.evaluate(async base64 => {
    const JSZip = (await import('/node_modules/.vite/deps/jszip.js')).default;
    const { importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const bytes = Uint8Array.from(atob(base64), c => c.charCodeAt(0)), zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file('word/document.xml')!.async('string');
    const reopened = await importFile(new File([bytes], 'again.docx'));
    const nodes: Record<string, number> = {};
    const walk = (n: any) => { nodes[n.type] = (nodes[n.type] || 0) + 1; n.content?.forEach(walk); };
    walk(parseDocumentHtml(reopened.html!));
    return { paper: reopened.document.page, xml, nodes };
  }, (await fs.readFile(file)).toString('base64'));
  expect(result.paper.size).toBe('Custom'); expect(result.paper.width).toBeCloseTo(297, 1); expect(result.paper.height).toBeCloseTo(420, 1);
  expect(result.paper.columns?.count).toBe(2); expect(result.nodes.columnBreak).toBe(1); expect(result.nodes.pageBreak).toBe(1);
  expect(result.xml).toContain('wp:anchor'); expect(result.xml).toContain('behindDoc="1"');
  expect(result.xml).toContain('수식: a over b'); expect(result.xml).toContain('w:lineRule="atLeast"');
});

test('percentage spacing is written as OOXML line multiples, and short numeric cells keep nowrap after sanitizing and reopening', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const JSZip = (await import('/node_modules/.vite/deps/jszip.js')).default;
    const source = '<p style="line-height:160%">ONE</p><p style="line-height:80%">TWO</p><table><tr><td style="white-space:nowrap"><p>20.</p></td></tr></table>';
    const loaded = await importFile(new File([source], 'spacing.html'));
    const doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
    const zip = await JSZip.loadAsync(await (await exportDocx(doc)).arrayBuffer());
    const body = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
    const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    return { lines: [...body.getElementsByTagNameNS(w, 'spacing')].map(e => e.getAttributeNS(w, 'line')).filter(Boolean),
      nowrap: body.getElementsByTagNameNS(w, 'noWrap').length, content: doc.content };
  });
  expect(result.lines).toContain('384'); expect(result.lines).toContain('192'); expect(result.nowrap).toBe(1);
});

test('equation SVG preparation rejects external resources and leaves readable scripts on invalid previews', async ({ page }) => {
  await page.goto('/');
  const requests: string[] = []; page.on('request', req => { if (req.url().startsWith('https://equation.invalid')) requests.push(req.url()); });
  const result = await page.evaluate(async () => {
    const { prepareEditableHwp } = await import('/src/hwp/svg-images.ts');
    const { defaultPage } = await import('/src/types.ts');
    const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24"><script>alert(1)</script><image href="https://equation.invalid/a.png"/><rect width="24" height="24" fill="red"/></svg>';
    const root = document.createElement('div'), image = document.createElement('img'); image.width = image.height = 24;
    image.dataset.hwpSvg = svg; image.alt = '수식: ORIGINAL'; root.append(image);
    const safe = await prepareEditableHwp({ html: root.innerHTML, page: defaultPage, warnings: [] });
    image.width = 0; image.height = 0;
    const invalid = await prepareEditableHwp({ html: root.innerHTML, page: defaultPage, warnings: [] });
    const controller = new AbortController(); let aborted = false;
    image.width = image.height = 24;
    try { await prepareEditableHwp({ html: root.innerHTML + root.innerHTML, page: defaultPage, warnings: [] }, () => controller.abort(), controller.signal); }
    catch { aborted = controller.signal.aborted; }
    return { safe: safe.html, invalid: invalid.html, warnings: invalid.warnings, aborted };
  });
  expect(result.safe).toContain('data:image/png;base64,'); expect(result.safe).not.toContain('equation.invalid');
  expect(result.invalid).toContain('수식: ORIGINAL'); expect(result.warnings).toHaveLength(1); expect(requests).toEqual([]); expect(result.aborted).toBe(true);
});

test('column pagination keeps page-relative pictures out of overflow measurements and consumes a boundary break once', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { defaultPage } = await import('/src/types.ts');
    const floating = JSON.stringify({ horizontalPosition: { relative: 'page', offset: 720 * 9525 }, verticalPosition: { relative: 'page', offset: 100 * 9525 }, wrap: { type: 0 } });
    const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==';
    const wrapper = document.createElement('div'), img = document.createElement('img');
    img.src = png; img.width = img.height = 40; img.dataset.imageFloating = floating; wrapper.append(img);
    const html = `<div data-column-count="2" data-column-space="24" class="document-columns" style="column-count:2;column-gap:24px;column-fill:auto"><p style="height:300px;margin:0">FIRST${wrapper.innerHTML}</p><div data-column-break="true" class="column-break"></div><p style="height:300px;margin:0">SECOND</p><div data-column-break="true" class="column-break"></div><p style="height:300px;margin:0">THIRD</p></div>`;
    const layout = await createPrintLayout(html, defaultPage);
    const pages = [...layout.element.querySelectorAll('.print-content')].map(p => ({ text: p.textContent, left: (p.querySelector('p') as HTMLElement)?.offsetLeft }));
    const count = layout.pages; layout.dispose(); return { count, pages };
  });
  expect(result.count).toBe(2);
  expect(result.pages[0].text).toContain('FIRST'); expect(result.pages[0].text).toContain('SECOND');
  expect(result.pages[1].text).toBe('THIRD'); expect(result.pages[1].left).toBe(0);
});

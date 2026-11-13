import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-script-fonts-fixture.cjs');
const text = '한글 문단은 원래 지정한 글꼴과 크기로 줄바꿈을 보존해야 합니다. '.repeat(14) + 'Latin 123, Next line.';
let directory: string, files: Record<string, string>;
test.beforeAll(async () => {
  directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-script-fallback-'));
  files = await fixture(directory, { hangulFace: '휴먼명조', body: text });
});
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

for (const format of ['hwp', 'hwpx']) test(`${format}: selected Korean face and line layout survive two real DOCX saves`, async ({ page }) => {
  await page.goto('/');
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  await expect(editor).toBeVisible();
  const result = await page.evaluate(async ({ bytes, format, text }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const reader = new HwpReaderSession(); let doc;
    try { await reader.open(new File([new Uint8Array(bytes)], 'fallback.' + format)); doc = importEditableHwp(await reader.editable(), 'fallback.' + format); }
    finally { reader.close(); }
    const states = [], saves = [];
    for (let pass = 0; pass < 3; pass++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try {
        const p = [...layout.element.querySelectorAll('p')].find(p => p.textContent === text);
        if (!p) throw Error('Source paragraph missing or split unexpectedly');
        const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT), lines = new Set<number>();
        let koreanFamily = '', latinFamily = '';
        while (walker.nextNode()) {
          const node = walker.currentNode, value = node.textContent || '', style = getComputedStyle(node.parentElement!);
          if (/[가-힣]/.test(value)) koreanFamily = style.fontFamily;
          if (/[A-Za-z]/.test(value)) latinFamily = style.fontFamily;
          const range = document.createRange(); range.selectNodeContents(node);
          for (const box of range.getClientRects()) if (box.width > 0) lines.add(Math.round(box.y * 32) / 32);
        }
        const style = getComputedStyle(p), box = p.getBoundingClientRect();
        states.push({ text: p.textContent, pages: layout.pages, height: box.height, width: box.width,
          lineCount: lines.size, koreanFamily, latinFamily, size: style.fontSize, line: style.lineHeight });
      } finally { layout.dispose(); }
      if (pass < 2) {
        const blob = await exportDocx(doc); saves.push([...new Uint8Array(await blob.arrayBuffer())]);
        const loaded = await importFile(new File([blob], 'again.docx'));
        doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
      }
    }
    return { states, saves };
  }, { bytes: [...await fs.readFile(files[format])], format, text });
  await fs.writeFile(test.info().outputPath('script-font-layout.json'), JSON.stringify(result.states, null, 2));
  for (const [i, bytes] of result.saves.entries()) await fs.writeFile(test.info().outputPath(`textbox-save-script-fallback-${format}-${i + 1}.docx`), Buffer.from(bytes));
  for (const state of result.states) {
    expect(state.text).toBe(text);
    expect(state.koreanFamily).toContain('Studio Korean Serif');
    expect(state.latinFamily).toBe('"Courier New"');
    expect(state.size).toBe(result.states[0].size);
    expect(state.line).toBe(result.states[0].line);
    expect(state.lineCount).toBe(result.states[0].lineCount);
    expect(state.height).toBeCloseTo(result.states[0].height, 1);
    expect(state.width).toBeCloseTo(result.states[0].width, 1);
    expect(state.pages).toBe(result.states[0].pages);
  }
});

import { test, expect } from '@playwright/test';
import { Bookmark, Document, HeadingLevel, LineRuleType, Packer, Paragraph, TextRun } from 'docx';
import { writeFile } from 'node:fs/promises';

// Independent source fixtures. These are regression cases, not Word-oracle
// additions or adjustments to the frozen 398/140 paragraph sets.
async function fixture() {
  const children = [];
  for (const [i, rule] of [LineRuleType.AUTO, LineRuleType.AT_LEAST, LineRuleType.EXACT].entries()) {
    const spacing = { before: 0, after: 0, line: rule === LineRuleType.AUTO ? 259 : 300, lineRule: rule };
    for (const heading of [false, true]) for (const marked of [false, true]) {
      const runs = [new TextRun('Before '), new TextRun({ text: 'bold', bold: true }), new TextRun(' after')];
      children.push(new Paragraph({
        ...(heading ? { heading: HeadingLevel.HEADING_2 } : {}), spacing,
        children: marked ? [new Bookmark({ id: `anchor_${i}_${heading}`, children: runs })] : runs,
      }));
    }
  }
  return [...await Packer.toBuffer(new Document({
    styles: { default: { document: { run: { font: 'Arial', size: 26 } }, heading2: { run: { font: 'Arial', size: 26, bold: false } } } },
    sections: [{ children }],
  }))];
}

for (const mode of ['editor', 'print'] as const) test(`empty bookmark anchors do not enlarge ${mode} paragraph or heading lines`, async ({ page }) => {
  await page.goto('/');
  const bytes = await fixture();
  await page.locator('input[type=file]').first().setInputFiles({ name: 'bookmark-lines.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(bytes) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  await expect(editor.locator('[data-docx-bookmark]')).toHaveCount(6);
  await page.evaluate(() => document.fonts.ready);
  if (mode === 'print') {
    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    await expect(page.locator('.print-content [data-docx-bookmark]')).toHaveCount(6);
  }
  const root = mode === 'editor' ? editor : page.locator('.print-content');
  await expect.poll(() => root.evaluateAll(roots => roots.flatMap(root => [...root.querySelectorAll('p,h2')])
    .filter(p => p.textContent === 'Before bold after').map(p => p.getBoundingClientRect().height))).toHaveLength(12);
  const rows = await root.evaluateAll(roots => roots.flatMap(root => [...root.querySelectorAll<HTMLElement>('p,h2')])
    .filter(p => p.textContent === 'Before bold after').map(p => ({ height: p.getBoundingClientRect().height, line: p.dataset.wordLineRule, tag: p.tagName })));
  for (let i = 0; i < rows.length; i += 2) {
    expect(rows[i + 1].line).toBe(rows[i].line);
    expect(rows[i + 1].tag).toBe(rows[i].tag);
    expect(Math.abs(rows[i + 1].height - rows[i].height), `${mode} ${rows[i].tag} ${rows[i].line}`).toBeLessThan(.2);
  }
});

test('bookmarks retain text positions through two DOCX saves without changing print line heights', async ({ page }, info) => {
  await page.goto('/');
  const result = await page.evaluate(async bytes => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const states = [], saves = [];
    let file = new File([new Uint8Array(bytes)], 'bookmark-lines.docx');
    for (let stage = 0; stage < 3; stage++) {
      const imported = await importFile(file), content = parseDocumentHtml(imported.html!);
      const layout = await createPrintLayout(generateHTML(content, editorExtensions), imported.document.page);
      try {
        const blocks = [...layout.element.querySelectorAll<HTMLElement>('.print-content p,.print-content h2')];
        states.push({ text: blocks.map(p => p.textContent), heights: blocks.map(p => p.getBoundingClientRect().height), bookmarks: [...layout.element.querySelectorAll('[data-docx-bookmark]')].map(b => {
          const range = document.createRange(); range.selectNodeContents(b.parentElement!); range.setEndBefore(b);
          return { name: b.getAttribute('data-docx-bookmark'), before: range.toString() };
        }) });
      } finally { layout.dispose(); }
      if (stage < 2) { const blob = await exportDocx({ ...imported.document, content }); saves.push([...new Uint8Array(await blob.arrayBuffer())]); file = new File([blob], 'again.docx'); }
    }
    return { states, saves };
  }, await fixture());
  for (const state of result.states) {
    expect(state.text).toEqual(result.states[0].text);
    expect(state.bookmarks).toEqual(result.states[0].bookmarks);
    expect(state.bookmarks).toHaveLength(6);
    for (let i = 0; i < state.heights.length; i++) expect(Math.abs(state.heights[i] - result.states[0].heights[i])).toBeLessThan(.2);
    for (let i = 0; i < state.heights.length; i += 2) expect(Math.abs(state.heights[i + 1] - state.heights[i])).toBeLessThan(.2);
  }
  for (const [i, bytes] of result.saves.entries()) await writeFile(info.outputPath(`textbox-save-bookmark-line-${i + 1}.docx`), Buffer.from(bytes));
});

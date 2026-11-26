import { expect, test } from '@playwright/test';

test('native Korean serif struts keep both script faces ahead of the view fallback', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).waitFor();
  const found = await page.evaluate(async () => {
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
    const pair = '휴먼명조, &quot;Palatino Linotype&quot;';
    const paragraph = (text: string, extra = '', font = pair) =>
      `<p style="line-height:17.9px;font-family:${font}" ${extra}>${text}</p>`;
    const content = parseDocumentHtml([
      paragraph('혼용 첫 줄'),
      paragraph('문서의 확정 줄', 'data-word-line-rule="exact"'),
      paragraph('다른 대체 서체', '', pair + ', serif'),
      '<table><tr><td>' + paragraph('일반 셀') + '</td></tr></table>',
    ].join(''));
    const root = document.createElement('div');
    root.className = 'tiptap';
    root.innerHTML = generateHTML(content, editorExtensions);
    document.body.append(root);
    try {
      return [...root.querySelectorAll('p')].map(p => ({
        requested: p.style.fontFamily,
        displayed: getComputedStyle(p).fontFamily,
        marked: p.hasAttribute('data-hwp-korean-serif-strut'),
      }));
    } finally { root.remove(); }
  });
  expect(found).toHaveLength(4);
  expect(found[0].requested).toBe('휴먼명조, "Palatino Linotype"');
  expect(found[0].marked).toBe(true);
  expect(found[0].displayed).toContain('Palatino Linotype", "Studio Korean Serif"');
  expect(found.slice(1).every(p => !p.displayed.includes('Studio Korean Serif'))).toBe(true);
});

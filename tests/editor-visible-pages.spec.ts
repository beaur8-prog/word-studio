import { test, expect } from '@playwright/test';

for (const scale of [1, .75]) test(`paper coverage follows real content at ${scale} scale, including a reflowed last table`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async scale => {
    const { visibleEditorPages } = await import('/src/editor-visible-pages.ts');
    const outer = document.createElement('div'); outer.style.transform = `scale(${scale})`; outer.style.transformOrigin = 'top left'; document.body.append(outer);
    const root = document.createElement('div'); root.style.cssText = 'position:relative;width:600px;min-height:5000px'; outer.append(root);
    try {
      root.innerHTML = '<p style="margin:0;line-height:20px">FIRST</p><table style="position:absolute;top:1160px;width:400px"><tr><td>LAST TABLE</td></tr></table>';
      const undercount = visibleEditorPages(root, 1122, 1);
      root.querySelector('table')!.remove();
      const afterDeletion = visibleEditorPages(root, 1122, 1);
      const authoredBlankSheets = visibleEditorPages(root, 1122, 4);
      root.innerHTML = '<div data-section-content="{}" style="height:3000px">MIXED PAPER SIZES</div>';
      return { undercount, afterDeletion, authoredBlankSheets, sections: visibleEditorPages(root, 1122, 3) };
    } finally { outer.remove(); }
  }, scale);
  expect(result).toEqual({ undercount: 2, afterDeletion: 1, authoredBlankSheets: 4, sections: 3 });
});

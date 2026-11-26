import { expect, test } from '@playwright/test';

test('only an automatic break after the short empty line can combine two native table fragments', async ({page}) => {
  await page.goto('/');
  await page.locator('#root .tiptap[role="textbox"]').waitFor();
  const result = await page.evaluate(async () => {
    const { measuredAdjacentNativeTables } = await import('/src/hwp/export-adjacent-tables.ts');
    const { hiddenEmptyLineClass } = await import('/src/page-empty-lines.ts');
    const first = document.createElement('section'), second = document.createElement('section');
    const cells = '<colgroup><col style="width:200px"><col style="width:200px"></colgroup>';
    first.innerHTML = '<div class="print-content" style="position:static;width:400px;height:600px">' +
      '<p>Earlier source material</p><table data-hwp-inline-table-top="141" style="width:400px;height:32px">' +
      '<colgroup><col style="width:400px"></colgroup><tbody><tr><td>Table title</td></tr></tbody></table>' +
      '<p class="' + hiddenEmptyLineClass + '" style="font-size:5pt;line-height:8pt;margin:0"></p></div>';
    second.innerHTML = '<div class="print-content" style="position:static;width:400px;height:600px">' +
      '<table style="width:400px;height:500px">' + cells +
      '<tbody><tr><td>One</td><td>Two</td></tr><tr><td>Three</td><td>Four</td></tr></tbody></table></div>';
    second.querySelector('table')!.dataset.tableAnchor = JSON.stringify({x:0,y:1.88,bottom:1.88,reference:'paragraph',horizontal:'margin'});
    document.body.append(first, second);
    try {
      const state = () => {
        const proof = measuredAdjacentNativeTables(first,second);
        return proof && [proof.label.rows, proof.main.rows, proof.main.columns, proof.spacerFont];
      };
      const automatic = state();
      second.dataset.keepBlank = 'true';
      const explicit = state();
      delete second.dataset.keepBlank;
      second.dataset.sectionFirst = 'true';
      const section = state();
      delete second.dataset.sectionFirst;
      first.querySelector('p:last-child')!.append(document.createElement('br'));
      const hardBreak = state();
      return {automatic,explicit,section,hardBreak};
    } finally { first.remove(); second.remove(); }
  });
  expect(result.automatic).toEqual([1,2,2,500]);
  expect(result.explicit).toBeUndefined();
  expect(result.section).toBeUndefined();
  expect(result.hardBreak).toBeUndefined();
});

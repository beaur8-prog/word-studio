import { expect, test } from '@playwright/test';

test('native fixed table measurement retains the same line boxes and inline prefixes', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { layoutHwpCellLines } = await import('/src/hwp-cell-line-metrics');
    const line = (text: string, size = '20px') =>
      `<p data-hwp-cell-line="1 Percent 20 32 20 20" style="font:${size}/32px monospace;margin:0">${text}</p>`;
    const grid = `<table style="width:204px;table-layout:fixed;border-collapse:collapse"><colgroup>
      <col style="width:102px"><col style="width:102px"></colgroup><tbody><tr>
      <td rowspan="2" style="vertical-align:middle">${line('one two three four five six')}${line('seven eight')}</td>
      <td>${line('inline<br>break retained')}${line('<span style="font-size:28px">large</span>')}</td></tr>
      <tr><td>${line('two plain lines in one cell')}</td></tr></tbody></table>`;
    const make = () => {
      const root = document.createElement('div');
      root.className = 'tiptap';
      root.style.cssText = 'position:absolute;left:20px;top:20px;width:210px;visibility:hidden';
      root.innerHTML = grid;
      document.body.append(root);
      return root;
    };
    const ordinary = make(), batched = make();
    const read = (root: HTMLElement) => [...root.querySelectorAll('td,p,br,[data-hwp-cell-prefix]')].map(el => {
      const { x,y,width,height } = el.getBoundingClientRect();
      return [el.tagName, x,y,width,height, el.getAttribute('data-hwp-cell-terminal'), el.textContent];
    });
    try {
      layoutHwpCellLines(ordinary);
      layoutHwpCellLines(batched, true);
      const expected = read(ordinary), actual = read(batched);
      layoutHwpCellLines(batched, true);
      return { expected, actual, repeated: read(batched) };
    } finally { ordinary.remove(); batched.remove(); }
  });

  expect(result.expected.filter(line => line[0] === 'SPAN').length).toBeGreaterThan(0);
  expect(result.actual).toEqual(result.expected);
  expect(result.repeated).toEqual(result.expected);
});

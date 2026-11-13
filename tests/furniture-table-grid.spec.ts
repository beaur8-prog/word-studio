import { test, expect } from '@playwright/test';

test('authored fixed header grids retain their widths in flat, framed and print stories', async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async () => {
    const { layoutFurnitureImages } = await import('/src/furniture-layout.ts');
    const cell = (width: number, text: string, attrs = '') => `<td colwidth="${width}" style="padding:0;border:0" ${attrs}><p style="font-family:Arial;font-size:10pt;line-height:16px;margin:0">${text}</p></td>`;
    const html = `<table data-word-table-layout="fixed"><tbody><tr>${cell(350, 'AUTHORED HEADER COLUMN WIDTH SHOULD STAY WIDE', 'rowspan="2"')}${cell(50, 'FIELD')}${cell(80, 'PAGE')}</tr><tr>${cell(50, 'A')}${cell(80, 'B')}</tr></tbody></table>`;
    const results = [];
    for (const kind of ['flat', 'framed', 'print']) {
      const paper = document.createElement('div'), story = document.createElement('div');
      paper.className = kind === 'framed' ? 'tiptap section-paper' : kind === 'print' ? 'print-page' : 'document-stack';
      paper.style.cssText = 'position:fixed;left:-100000px;top:0;width:800px;height:1100px;padding:0;visibility:hidden';
      story.className = 'document-furniture ' + (kind === 'framed' ? 'section-header' : kind === 'print' ? 'print-header' : 'paper-header');
      story.style.cssText = 'position:absolute;left:50px;right:50px;top:20px'; story.innerHTML = html; paper.append(story); document.body.append(paper);
      try {
        await document.fonts.ready; layoutFurnitureImages(paper);
        const table = story.querySelector('table')!, first = table.rows[0], snapshot = () => ({ width: table.getBoundingClientRect().width,
          height: table.getBoundingClientRect().height, cells: [...first.cells].map(cell => cell.getBoundingClientRect().width),
          grid: [...table.querySelectorAll(':scope > colgroup > col')].map(col => (col as HTMLElement).style.width), text: table.textContent, groups: table.querySelectorAll(':scope > colgroup').length });
        const before = snapshot(); layoutFurnitureImages(paper); results.push({ kind, before, after: snapshot() });
      } finally { paper.remove(); }
    }
    return results;
  });
  for (const state of states) {
    expect(state.before.grid).toEqual(['350px', '50px', '80px']); expect(state.before.groups).toBe(1);
    expect(state.before.width).toBeCloseTo(480, 0); expect(state.before.cells).toEqual([350, 50, 80]);
    expect(state.after).toEqual(state.before); expect(state.before.height).toBeCloseTo(states[0].before.height, 0);
    expect(state.before.text).toBe(states[0].before.text);
  }
});

test('ambiguous, relative and autofit furniture table widths retain their existing policy', async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async () => {
    const { layoutFurnitureImages } = await import('/src/furniture-layout.ts');
    const paper = document.createElement('div'); paper.className = 'print-page'; paper.style.cssText = 'position:fixed;left:-100000px;top:0;width:800px;height:1000px';
    const story = document.createElement('div'); story.className = 'document-furniture';
    story.innerHTML = '<table data-word-table-layout="fixed" style="width:100%"><tr><td colwidth="100"><p>RELATIVE</p></td></tr></table>' +
      '<table data-word-table-layout="autofit"><tr><td colwidth="100"><p>AUTOFIT</p></td></tr></table>' +
      '<table data-word-table-layout="fixed"><tr><td colspan="2" colwidth="100"><p>INCOMPLETE</p></td></tr></table>' +
      '<table data-word-table-layout="fixed" style="width:400px"><tr><td colwidth="100"><p>OTHER WIDTH</p></td></tr></table>';
    paper.append(story); document.body.append(paper);
    try { const before = story.innerHTML; layoutFurnitureImages(paper); return { before, after: story.innerHTML, groups: story.querySelectorAll('colgroup').length }; }
    finally { paper.remove(); }
  });
  expect(states.groups).toBe(0); expect(states.after).toBe(states.before);
});

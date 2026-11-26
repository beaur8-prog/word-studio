import { expect, test } from '@playwright/test';

test('a nested print head supplies reserve only when both editable cell cuts agree', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { measuredNestedHeadSpace } = await import('/src/editor-nested-row-projection.ts');
    const host = document.createElement('div');
    host.className = 'tiptap';
    const paragraph = (text: string) => `<p style="margin:0;padding:0;font:12px/20px Arial">${text}</p>`;
    const cell = (html: string) => `<td style="border:0;padding:0;vertical-align:top">${html}</td>`;
    const inner = `<div class="tableWrapper" style="margin:0"><table style="margin:0;border:0"><tbody>` +
      `<tr data-row-split="allow">${cell(['N1', 'N2', 'N3'].map(paragraph).join(''))}</tr></tbody></table></div>`;
    host.innerHTML = `<table style="width:200px;margin:0;border:0"><colgroup><col style="width:100px"><col style="width:100px"></colgroup>
      <tbody><tr>${cell(paragraph('Previous left'))}${cell(paragraph('Previous right'))}</tr>
      <tr data-row-split="allow" style="height:150px">${cell(paragraph('Lead') + inner + paragraph('Tail'))}
      ${cell(['R1', 'R2', 'R3', 'R4', 'R5'].map(paragraph).join(''))}</tr></tbody></table>`;
    document.body.append(host);
    try {
      const row = host.querySelector<HTMLTableElement>('table')!.rows[1];
      const cuts = [{ head: 'LeadN1N2', tail: 'N3Tail' }, { head: 'R1R2R3', tail: 'R4R5' }];
      const pieces = [...row.cells].map(c => c.textContent);
      const head = measuredNestedHeadSpace(row, cuts, 65, 125, 1);
      const stale = measuredNestedHeadSpace(row, [cuts[0], { head: 'R1R3R2', tail: 'R4R5' }], 65, 125, 1);
      const earlier = measuredNestedHeadSpace(row, cuts, 45, 125, 1);
      return { head, stale, earlier, before: pieces, after: [...row.cells].map(c => c.textContent),
        height: row.getBoundingClientRect().height, probeRemoved: document.querySelectorAll('.print-page[aria-hidden="true"]').length === 0 };
    } finally { host.remove(); }
  });
  expect(result.head).toBeCloseTo(5, 2);
  expect(result.stale).toBeUndefined();
  expect(result.earlier).toBeUndefined();
  expect(result.after).toEqual(result.before);
  expect(result.height).toBe(150);
  expect(result.probeRemoved).toBe(true);
});

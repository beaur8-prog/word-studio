import { expect, test } from '@playwright/test';
import { encodeHwpCellLine, parseHwpCellLine, type HwpCellLine } from '../src/hwp-cell-line';

test('the authored terminal policy is bounded and round-trips with its original height', () => {
  const quantity: HwpCellLine = { rule:'Percent', font:14.7, advance:19.1, end:14.7, base:14.7 };
  const terminal = { ...quantity, authoredTerminal:{ text:'a789ab89532bbb23',width:191,row:26.5 } };
  expect(parseHwpCellLine(encodeHwpCellLine(quantity))).toEqual(quantity);
  expect(parseHwpCellLine(encodeHwpCellLine(terminal))).toEqual(terminal);
  const centered: HwpCellLine = { ...terminal, authoredTerminal:{ ...terminal.authoredTerminal,middle:true } };
  expect(parseHwpCellLine(encodeHwpCellLine(centered))).toEqual(centered);
  expect(parseHwpCellLine(encodeHwpCellLine(terminal) + ' center')).toBeNull();
  expect(parseHwpCellLine(encodeHwpCellLine({ ...terminal, base:19.1 }))).toBeNull();
  expect(parseHwpCellLine(encodeHwpCellLine({ ...terminal, authoredTerminal:{ ...terminal.authoredTerminal, width:Infinity } }))).toBeNull();
});

test('a witnessed terminal reserves one line without changing the painted characters', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const { layoutHwpAuthoredCellTerminals } = await import('/src/hwp-authored-cell-terminal.ts');
    const body = document.createElement('div');
    body.className = 'tiptap';
    document.body.append(body);
    try {
      const source = 'Sample body text', fingerprint = hwpWrapFingerprint(source);
      return ['match','changed-text','changed-width','changed-height','two-lines','legacy','moved-top','moved-middle','middle'].map(mode => {
        const centered = mode === 'middle' || mode === 'moved-top';
        const provedCenter = mode === 'middle' || mode === 'moved-middle';
        const minimum = centered ? 25 : 26;
        body.innerHTML = `<table style="table-layout:fixed;width:200px;border-collapse:collapse"><colgroup><col style="width:200px"></colgroup>
          <tbody><tr style="height:${minimum}px"><td style="padding:5px 0 6px;vertical-align:${centered ? 'middle' : 'top'};border:0"><p style="font-size:14px;line-height:14px;margin:0"><span style="line-height:21px">${mode === 'changed-text' ? 'Sample edited text' : source}${mode === 'two-lines' ? '<br>next' : ''}</span></p></td></tr></tbody></table>`;
        const table = body.querySelector('table')!, p = table.querySelector('p')!;
        p.dataset.hwpCellLine = mode === 'legacy' ? '1 Percent 14 20 14 14' :
          `2 Percent 14 20 14 14 ${fingerprint} ${mode === 'changed-width' ? 170 : 200} ${mode === 'changed-height' ? 42 : minimum}${provedCenter ? ' middle' : ''}`;
        const range = document.createRange();
        range.selectNodeContents(p);
        const original = range.getBoundingClientRect();
        layoutHwpAuthoredCellTerminals(table);
        const painted = range.getBoundingClientRect();
        return { hasTerminal:p.style.height === '14px', text:p.textContent, inkChanged:Math.abs(original.width - painted.width) > .12 };
      });
    } finally { body.remove(); }
  });
  expect(result.map(row => row.hasTerminal)).toEqual([true,false,false,false,false,false,false,false,true]);
  expect(result[0]).toEqual({ hasTerminal:true,text:'Sample body text',inkChanged:false });
});

import { expect, test } from '@playwright/test';

test('a restored HWP title uses visible ink to check that its native table fits', async ({ page }) => {
  await page.goto('/');
  const measured = await page.evaluate(async () => {
    const { measuredTableLead } = await import('/src/hwp/export-table-lead.ts');
    const { measuredHwpLineInk } = await import('/src/hwp-line-ink.ts');
    const { hwpWrapFingerprint, normalizeHwpWrapText } = await import('/src/hwp-text-advances.ts');
    const title = '원래 한 줄 제목과 이어지는 표';
    await document.fonts.load('400 18.67px "Noto Sans KR"', title);
    const frame = document.createElement('section');
    frame.lang = 'ko';
    frame.style.cssText = 'position:absolute;top:32px;left:48px;width:320px;height:400px;';
    frame.innerHTML = '<p></p><table style="height:220px;width:300px;border-collapse:collapse;"><tbody><tr><td>표의 본문</td></tr></tbody></table>';
    document.body.append(frame);
    try {
      const p = frame.querySelector('p')!, table = frame.querySelector('table')!;
      p.style.cssText = 'margin:0;font:400 18.67px/24.27px "Noto Sans KR";';
      p.textContent = title;
      const proof = {
        fingerprint: hwpWrapFingerprint(title), ends: [normalizeHwpWrapText(title).length],
        origin: { page: 3, x: 48, y: 32 }, bodyWidth: 320, singleBodyFallback: true,
      };
      p.dataset.hwpSourceWrap = JSON.stringify(proof);
      const range = document.createRange();
      range.selectNodeContents(p);
      const emBottom = range.getBoundingClientRect().bottom;
      const ink = measuredHwpLineInk(p, 1, 18.67, true);
      if (!ink || !(emBottom - ink.bottom > .5)) throw Error('The fixture must separate font box and painted ink');
      const setTableTop = (top: number) => { table.style.marginTop = (top - p.getBoundingClientRect().bottom) + 'px'; };
      setTableTop((emBottom + ink.bottom) / 2);
      const ordinary = measuredTableLead(frame, 2);
      p.dataset.hwpPrintSourceFlow = 'true';
      const withProof = measuredTableLead(frame, 2);
      const wrongPage = measuredTableLead(frame, 1);
      p.dataset.hwpSourceWrap = JSON.stringify({ ...proof, fingerprint: hwpWrapFingerprint(title + ' 수정') });
      const changed = measuredTableLead(frame, 2);
      p.dataset.hwpSourceWrap = JSON.stringify(proof);
      setTableTop(ink.bottom - 1);
      const overlapping = measuredTableLead(frame, 2);
      return { ordinary, withProof, wrongPage, changed, overlapping };
    } finally { frame.remove(); }
  });
  expect(measured.withProof).toEqual({ sheet: 2, literal: '원래 한 줄 제목과 이어지는 표', width: 300, height: 220, rows: 1 });
  expect(measured.ordinary).toBeUndefined();
  expect(measured.wrongPage).toBeUndefined();
  expect(measured.changed).toBeUndefined();
  expect(measured.overlapping).toBeUndefined();
});

test('native empty paragraph sizes use the terminal text boxes across the print edge', async ({ page }) => {
  await page.goto('/');
  const measured = await page.evaluate(async () => {
    const { measuredTerminalEmptyFonts } = await import('/src/hwp/export-empty-paragraph-format.ts');
    const frame = document.createElement('section');
    frame.style.cssText = 'position:absolute;top:32px;left:48px;width:320px;height:100px;';
    frame.innerHTML = '<p data-exact-line-height style="margin:0;font-size:3pt;line-height:6.4px;height:6.4px"></p>' +
      '<p style="margin:0;line-height:20px">A heading</p>' +
      '<p data-exact-line-height style="margin:0;font-size:14pt;line-height:29.8667px;height:29.8667px"></p>' +
      '<p style="margin:0;font-size:15pt;line-height:26px;height:48px"><span>The final text</span></p>';
    document.body.append(frame);
    try {
      const terminal = frame.lastElementChild as HTMLElement;
      const extendsPastFrame = terminal.getBoundingClientRect().bottom > frame.getBoundingClientRect().bottom;
      const fits = measuredTerminalEmptyFonts(frame,0);
      terminal.style.paddingTop = '60px';
      const overflows = measuredTerminalEmptyFonts(frame,0);
      return { extendsPastFrame, fits, overflows };
    } finally { frame.remove(); }
  });
  expect(measured.extendsPastFrame).toBe(true);
  expect(measured.fits).toEqual({page:0,heading:'A heading',terminal:'The final text',beforeHeading:300,beforeTerminal:1400});
  expect(measured.overflows).toBeUndefined();
});

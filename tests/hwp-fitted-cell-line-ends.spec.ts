import { expect, test } from '@playwright/test';

test('verified same-count HWP cell lines keep their painted row and authored character width', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { cellCharacterScaleDimensions } = await import('/src/cell-character-scale.ts');
    const { layoutHwpFittedCellLineEnds, layoutHwpSourceCellLines } = await import('/src/hwp-source-cell-lines.ts');
    const { normalizeHwpWrapText, hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
    const literal = 'AAAA BBBB CCCC DDDD EEEE FFFF';
    const paper = document.createElement('div');
    const observePeer = document.createElement('style');
    observePeer.textContent = 'table[data-test-peer] td:has(wbr) + td p { transform: translateX(1px) }';
    document.head.append(observePeer);
    paper.className = 'print-page'; paper.dataset.page = '1';
    paper.style.cssText = 'position:fixed;left:0;top:0;width:420px;height:310px';
    paper.innerHTML = '<div class="tiptap print-content" style="position:absolute;left:10px;top:10px;width:400px;height:280px;padding:0"></div>';
    document.body.append(paper);
    try {
      const body = paper.firstElementChild!;
      return ['original', 'edited space', 'different width', 'different scale', 'shifted peer'].map((mode) => {
        body.innerHTML = '<table data-hwp-centered-borders="continuation" style="width:332px;margin:0;table-layout:fixed;border-collapse:collapse">' +
          '<colgroup><col style="width:222px"><col style="width:110px"></colgroup><tbody><tr><td style="border:0;padding:1px">' +
          '<p style="font-family:monospace;font-size:16px;line-height:24px;margin:0;text-align:justify"></p></td>' +
          '<td><p>PEER</p></td></tr></tbody></table><p>FOLLOWING</p>';
        const table = body.querySelector('table')!;
        table.toggleAttribute('data-test-peer', mode === 'shifted peer');
        const p = table.querySelector('p')!, td = p.parentElement!;
        const peer = table.querySelector('td:nth-child(2) p')!;
        const peerX = peer.getBoundingClientRect().x;
        const span = document.createElement('span');
        span.style.setProperty('--word-character-scale', mode === 'different scale' ? '92' : '95');
        span.textContent = mode === 'edited space' ? literal.replace('CCCC DDDD', 'CCCC  DDDD') : literal;
        p.append(span);
        p.style.cssText += ';' + cellCharacterScaleDimensions(95);
        const width = p.getBoundingClientRect().width;
        const outer = td.getBoundingClientRect().width;
        const following = table.nextElementSibling!.getBoundingClientRect().top;
        const originalText = p.textContent!;
        const ends = () => paragraphLineOffsets(p).map(at => normalizeHwpWrapText(p.textContent!.slice(0, at)).length);
        const before = ends();
        p.dataset.hwpSourceWrap = JSON.stringify({ fingerprint: hwpWrapFingerprint(literal), ends: [12, 24],
          cellWidth: { outer: mode === 'different width' ? outer + 6 : outer, text: width } });
        layoutHwpSourceCellLines(paper);
        const attempts = new MutationObserver(() => {});
        attempts.observe(p, { childList: true, subtree: true });
        layoutHwpFittedCellLineEnds(paper);
        const triedWrap = attempts.takeRecords().some(record => [...record.addedNodes].some(node =>
          node instanceof Element && (node.matches('wbr') || node.querySelector('wbr'))));
        attempts.disconnect();
        return { mode, before, after: ends(), sameText: p.textContent === originalText,
          scale: getComputedStyle(p).transform, hard: p.querySelectorAll('br').length,
          triedWrap, peerMoves: peer.getBoundingClientRect().x !== peerX,
          followingMoves: table.nextElementSibling!.getBoundingClientRect().top !== following };
      });
    } finally { paper.remove(); observePeer.remove(); }
  });
  const { before, ...original } = result[0];
  expect(before).toHaveLength(1);
  expect(before[0]).toBeGreaterThan(12);
  expect(original).toEqual({ mode: 'original', after: [12], sameText: true,
    scale: 'matrix(0.95, 0, 0, 1, 0, 0)', hard: 0, triedWrap: true, peerMoves: false, followingMoves: false });
  for (const checked of result.slice(1)) {
    expect(checked.after).toEqual(checked.before);
    expect(checked.sameText).toBe(true);
    expect(checked.hard).toBe(0);
    expect(checked.peerMoves).toBe(false);
    expect(checked.followingMoves).toBe(false);
  }
  expect(result.find(checked => checked.mode === 'shifted peer')?.triedWrap).toBe(true);
});

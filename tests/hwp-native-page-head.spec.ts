import { expect, test } from '@playwright/test';

test('rejoin a verified source page heading and body only across an automatic boundary', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { measuredNativePageHead, restoreNativePageHead } = await import('/src/hwp/export-page-head');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances');
    const { default: init, HwpDocument } = await import('/node_modules/.vite/deps/@rhwp_core.js');
    const { default: wasm } = await import('/node_modules/@rhwp/core/rhwp_bg.wasm?url');
    await init({ module_or_path: wasm });
    const doc = HwpDocument.createEmpty();
    let repaired;
    try {
      const head = 'First line on a new sheet', body = ' The body follows its heading';
      doc.createBlankDocument();
      doc.pasteHtml(0, 0, 0, '<p>Previous text</p><p>' + head + '</p><p>' + body.trimStart() + '</p>');
      doc.insertParagraph(0, 2);
      const table = [...Array(doc.getParagraphCount(0))].map((_, i) => doc.getTextRange(0, i, 0, 0xffffffff));
      const leadIndex = table.indexOf(head), following = table.indexOf(body.trimStart());
      if (following !== leadIndex + 2) throw Error('Synthetic spacer is not between the two paragraphs');
      doc.insertText(0, following, 0, ' ');
      doc.applyCharFormat(0, following, 0, 1, JSON.stringify({ bold: true }));
      doc.applyCharFormat(0, leadIndex + 1, 0, 0, JSON.stringify({ fontSize: 400 }));
      doc.applyParaFormat(0, leadIndex, JSON.stringify({ pageBreakBefore: true }));
      const y = (owner: number, page: number) =>
        JSON.parse(doc.getPageTextLayout(page)).runs.find(run => run.paraIdx === owner && run.charStart === 0)?.y;
      const geometry = { page: 2, y: y(leadIndex, 1), nextY: y(following, 1), spacer: 400, next: hwpWrapFingerprint(body) };
      const a = document.createElement('div'), b = document.createElement('div');
      a.dataset.page = '1';
      b.dataset.page = '2';
      a.innerHTML = '<div class="print-content"><p></p><p style="font-size:4pt"></p></div>';
      b.innerHTML = '<div class="print-content"><p></p></div>';
      const title = a.querySelector('p')!, text = b.querySelector('p')!;
      title.textContent = head;
      text.textContent = body;
      title.dataset.hwpSourceWrap = JSON.stringify({ fingerprint: hwpWrapFingerprint(head), ends: [head.replace(/\s/gu, '').length], authoredPageHead: geometry });
      const automatic = measuredNativePageHead(a, b);
      b.dataset.keepBlank = 'true';
      const explicit = measuredNativePageHead(a, b);
      delete b.dataset.keepBlank;
      text.textContent = body + ' edit';
      const changed = measuredNativePageHead(a, b);
      doc.applyParaFormat(0, following, JSON.stringify({ pageBreakBefore: true }));
      const startingRun = JSON.parse(doc.getPageTextLayout(2)).runs.find(run => run.paraIdx === following);
      const leadingSpaceSeparate = startingRun?.charStart === 0 && startingRun.text === ' ';
      repaired = restoreNativePageHead(HwpDocument, doc, automatic ? [automatic] : [], 2);
      const restored = (owner: number) => JSON.parse(repaired.getPageTextLayout(1)).runs.find(run => run.paraIdx === owner && run.charStart === 0)?.y;
      return { leadingSpaceSeparate, preserved: repaired.getTextRange(0, following, 0, 0xffffffff), geometry,
        explicit, changed, before: doc.pageCount(), after: repaired.pageCount(), positions: [restored(leadIndex), restored(following)] };
    } finally {
      if (repaired && repaired !== doc) repaired.free();
      doc.free();
    }
  });

  expect(result.explicit).toBeUndefined();
  expect(result.changed).toBeUndefined();
  expect(result.before).toBe(3);
  expect(result.after).toBe(2);
  expect(result.leadingSpaceSeparate).toBe(true);
  expect(result.preserved).toBe(' The body follows its heading');
  expect(result.positions).toEqual([result.geometry.y, result.geometry.nextY]);
});

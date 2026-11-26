'use strict';

const { createHash } = require('node:crypto');

module.exports = async function wordPrintLineBreaks(page) {
  const captures = await page.evaluate(() => {
    const results = [];
    const sheets = document.querySelectorAll('.print-layout.preview .print-page');
    const range = document.createRange();
    const graphemes = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
    for (const [index, sheet] of [...sheets].entries()) {
      const paper = sheet.getBoundingClientRect();
      for (const p of sheet.querySelectorAll('.print-content > :is(p,h1,h2,h3,h4,h5,h6)')) {
        if (p.hasAttribute('data-print-paragraph-fragment') ||
            p.querySelector('br,svg,img,sup,sub,[data-note-reference],[data-document-field],[data-word-overlap],[data-flow-widget]')) continue;
        const bands = [];
        let seen = 0;
        const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
        const style = getComputedStyle(p);
        const writing = style.direction === 'ltr' && style.writingMode === 'horizontal-tb';
        if (!writing) continue;
        while (walker.nextNode()) {
          const text = walker.currentNode;
          if (text.parentElement.closest('[data-docx-tab]')) continue;
          for (const { segment, index: offset } of graphemes.segment(text.data)) {
            if (!segment.trim()) continue;
            range.setStart(text, offset);
            range.setEnd(text, offset + segment.length);
            const rect = range.getBoundingClientRect();
            if (!(rect.width > 0 && rect.height > 0)) continue;
            const last = bands.at(-1);
            const overlap = last ? Math.min(last.bottom, rect.bottom) - Math.max(last.top, rect.top) : 0;
            if (last && overlap >= Math.min(last.bottom - last.top, rect.height) * .5) {
              last.top = Math.min(last.top, rect.top);
              last.bottom = Math.max(last.bottom, rect.bottom);
            } else {
              bands.push({ start: seen, top: rect.top, bottom: rect.bottom });
            }
            seen++;
          }
        }
        if (!seen) continue;
        const position = p.getBoundingClientRect();
        results.push({ literal: p.textContent.replace(/\s/gu, ''), sheet: index + 1,
          starts: bands.map(band => band.start), glyphs: seen, lineY: bands.map(band => band.top - paper.top),
          y: position.top - paper.top, height: position.height, width: position.width,
          tabs: p.querySelectorAll('[data-docx-tab]').length, font: style.fontFamily,
          size: style.fontSize, rule: p.dataset.wordLineRule ?? '', auto: parseFloat(style.lineHeight) });
      }
    }
    return results;
  });
  return captures.map(({ literal, ...values }) => ({
    key: createHash('sha256').update(literal).digest('hex'), ...values,
  }));
};

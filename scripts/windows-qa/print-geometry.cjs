'use strict';

const { createHash } = require('node:crypto');

module.exports = async function printGeometry(page) {
  const values = await page.evaluate(() => {
    const result = [];
    const print = document.querySelector('.print-layout.preview');
    for (const [index, sheet] of [...(print?.querySelectorAll('.print-page') ?? [])].entries()) {
      const paper = sheet.getBoundingClientRect();
      const scale = Number.parseFloat(getComputedStyle(sheet).width) / paper.width;
      if (!(paper.width > 0 && paper.height > 0 && scale > 0)) throw Error('Invalid physical print paper');
      const relative = box => [(box.x-paper.x)*scale, (box.y-paper.y)*scale, box.width*scale, box.height*scale];
      const ink = [];
      const walker = document.createTreeWalker(sheet, NodeFilter.SHOW_TEXT);
      const range = document.createRange();
      while (walker.nextNode()) {
        const text = walker.currentNode;
        if (!text.data.trim() || text.parentElement?.closest('.print-page-label, [data-flow-widget]')) continue;
        range.selectNodeContents(text);
        const rects = [...range.getClientRects()].filter(box => box.width > 0 && box.height > 0);
        if (rects.length) ink.push([text.data, rects.map(relative)]);
      }
      const objects = [...sheet.querySelectorAll('table,img,svg,[data-document-textbox]')].flatMap(element => {
        const box = element.getBoundingClientRect();
        return box.width > 0 && box.height > 0 ? [[element.tagName, relative(box)]] : [];
      });
      result.push({ index: index+1, paper: [paper.width*scale,paper.height*scale], ink, objects });
    }
    if (!result.length || !result.some(sheet => sheet.ink.length)) throw Error('No printed glyph geometry');
    return result;
  });
  for (const item of values)
    for (const entry of item.ink)
      entry[0] = createHash('sha256').update(entry[0]).digest('hex');
  return values;
};

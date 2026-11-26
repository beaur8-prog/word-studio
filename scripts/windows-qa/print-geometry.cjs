'use strict';

const { createHash } = require('node:crypto');

module.exports = async function printGeometry(page) {
  const values = await page.evaluate(async includeImageIdentity => {
    const imageKeys = new Map();
    const imageKey = source => {
      if (!imageKeys.has(source)) imageKeys.set(source, crypto.subtle.digest('SHA-256', new TextEncoder().encode(source))
        .then(digest => Array.from(new Uint8Array(digest), byte => byte.toString(16).padStart(2, '0')).join('')));
      return imageKeys.get(source);
    };
    const result = [];
    const norm = value => value.replace(/\s/gu, '');
    const print = document.querySelector('.print-layout.preview');
    for (const [index, sheet] of [...(print?.querySelectorAll('.print-page') ?? [])].entries()) {
      const paper = sheet.getBoundingClientRect();
      const scale = Number.parseFloat(getComputedStyle(sheet).width) / paper.width;
      if (!(paper.width > 0 && paper.height > 0 && scale > 0)) throw Error('Invalid physical print paper');
      const relative = box => [(box.x-paper.x)*scale, (box.y-paper.y)*scale, box.width*scale, box.height*scale];
      const ink = [], paragraphs = [];
      const walker = document.createTreeWalker(sheet, NodeFilter.SHOW_TEXT);
      const range = document.createRange();
      while (walker.nextNode()) {
        const text = walker.currentNode;
        if (!text.data.trim() || text.parentElement?.closest('.print-page-label, [data-flow-widget]')) continue;
        range.selectNodeContents(text);
        const rects = [...range.getClientRects()].filter(box => box.width > 0 && box.height > 0);
        if (rects.length) ink.push([text.data, rects.map(relative)]);
      }
      for (const paragraph of sheet.querySelectorAll('.print-content p,.print-content h1,.print-content h2,.print-content h3')) {
        if (paragraph.closest('[data-flow-widget]')) continue;
        const text = norm(paragraph.textContent);
        if (!text) continue;
        range.selectNodeContents(paragraph);
        const lines = [...range.getClientRects()].filter(box => box.width > 0 && box.height > 0).map(relative);
        if (!lines.length) continue;
        const sorted = [...new Set(lines.map(box => Math.round(box[1]*10)/10))].sort((a,b)=>a-b);
        const style = getComputedStyle(paragraph), box = paragraph.getBoundingClientRect();
        paragraphs.push({ key: text, length:text.length, p:relative(box), y:[sorted[0],sorted.at(-1)], lines:sorted.length,
          font:style.fontFamily, size:style.fontSize, lineHeight:style.lineHeight });
      }
      const objects = [...sheet.querySelectorAll('table,img,svg,[data-document-textbox]')].flatMap(element => {
        const box = element.getBoundingClientRect();
        return box.width > 0 && box.height > 0 ? [[element.tagName, relative(box)]] : [];
      });
      const images = includeImageIdentity ? await Promise.all([...sheet.querySelectorAll('img')].flatMap(image => {
        const source = image.getAttribute('src') || '';
        const box = image.getBoundingClientRect();
        if (!source.startsWith('data:image/') || box.width <= 0 || box.height <= 0) return [];
        return [imageKey(source).then(key => [key, relative(box)])];
      })) : undefined;
      result.push({ index: index+1, paper: [paper.width*scale,paper.height*scale], ink, paragraphs, objects, images });
    }
    if (!result.length || !result.some(sheet => sheet.ink.length)) throw Error('No printed glyph geometry');
    return result;
  }, process.env.WINDOWS_CAPTURE_PRINT_IMAGE_IDENTITY === '1');
  for (const item of values) {
    for (const entry of item.ink) entry[0] = createHash('sha256').update(entry[0]).digest('hex');
    for (const paragraph of item.paragraphs) paragraph.key = createHash('sha256').update(paragraph.key).digest('hex');
  }
  return values;
};

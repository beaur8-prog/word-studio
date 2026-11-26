'use strict';

const { createHash } = require('node:crypto');
const { readFileSync, writeFileSync } = require('node:fs');

const digest = value => createHash('sha256').update(value).digest('hex');

async function capture(page) {
  const pieces = await page.evaluate(() => {
    const sheets = [...document.querySelectorAll('.print-layout.preview .print-page')];
    const result = { body:[], cells:[], repeatedHeaders:[], images:[], outside:0 };
    const range = document.createRange();
    for (const sheet of sheets) {
      const paper = sheet.getBoundingClientRect();
      const content = sheet.querySelector('.print-content');
      if (!content) continue;
      const walk = document.createTreeWalker(content, NodeFilter.SHOW_TEXT);
      while (walk.nextNode()) {
        const node = walk.currentNode;
        const owner = node.parentElement;
        if (!owner || owner.closest('[data-flow-widget],[data-page-number],[data-note-reference]')) continue;
        const value = node.data.replace(/\s/gu, '');
        if (!value) continue;
        range.selectNodeContents(node);
        const boxes = [...range.getClientRects()].filter(box => box.width > 0 && box.height > 0);
        if (!boxes.length) continue;
        if (boxes.some(box => box.right > paper.right + .5 || box.left < paper.left - .5 ||
            box.bottom > paper.bottom + .5 || box.top < paper.top - .5)) result.outside++;
        const kind = owner.closest('thead') ? 'repeatedHeaders' : owner.closest('td,th') ? 'cells' : 'body';
        result[kind].push(value);
      }
      for (const image of content.querySelectorAll('img')) {
        const source = image.getAttribute('src');
        const rect = image.getBoundingClientRect();
        if (source?.startsWith('data:image/') && rect.width > 0 && rect.height > 0)
          result.images.push(source);
      }
    }
    return result;
  });
  const files = { outside: pieces.outside, images: pieces.images.map(digest) };
  for (const scope of ['body','cells','repeatedHeaders']) {
    const value = pieces[scope].join('');
    const counts = new Map();
    for (const character of value) {
      const key = digest(character);
      counts.set(key, (counts.get(key) ?? 0) + 1);
    }
    files[scope] = { length: [...value].length, sequence: digest(value), counts: Object.fromEntries(counts) };
  }
  writeFileSync(process.env.WINDOWS_FIXED_PRINT_CONTENT_AUDIT, JSON.stringify(files), 'utf8');
}

function compare(before, after) {
  const summary = {};
  for (const scope of ['body','cells','repeatedHeaders']) {
    const [a,b] = [before[scope], after[scope]];
    const keys = new Set([...Object.keys(a.counts), ...Object.keys(b.counts)]);
    let fewer = 0, more = 0;
    for (const key of keys) {
      const difference = (b.counts[key] ?? 0) - (a.counts[key] ?? 0);
      if (difference < 0) fewer -= difference;
      else more += difference;
    }
    summary[scope] = { codepoints:[a.length,b.length], sameSequence:a.sequence===b.sequence, fewer, more };
  }
  summary.outside = [before.outside,after.outside];
  summary.inlineImages = { count:[before.images.length,after.images.length],
    sameBytesAndOrder:JSON.stringify(before.images)===JSON.stringify(after.images) };
  return summary;
}

if (require.main === module) {
  const [beforePath, afterPath] = process.argv.slice(2);
  if (!beforePath || !afterPath) throw Error('Expected both print audits');
  const read = path => JSON.parse(readFileSync(path, 'utf8'));
  console.log('WINDOWS_FIXED_PRINT_CONTENT ' + JSON.stringify(compare(read(beforePath),read(afterPath))));
} else module.exports = capture;

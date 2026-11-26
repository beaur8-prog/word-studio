'use strict';

function location(doc, literal) {
  const found = [];
  for (let page = 0; page < Math.min(doc.pageCount(), 16); page++) {
    const owners = new Map();
    for (const run of JSON.parse(doc.getPageTextLayout(page)).runs) {
      if (!Number.isInteger(run.paraIdx) || run.paraIdx === 0xffffffff || run.cellPath?.length || run.headerFooter) continue;
      const key = `${run.secIdx}:${run.paraIdx}`;
      const para = owners.get(key) ?? { page: page + 1, y: run.y, paragraph: run.paraIdx, literal: '' };
      para.literal += run.text;
      owners.set(key, para);
    }
    for (const para of owners.values())
      if (para.literal === literal) found.push({ page: para.page, y: para.y, paragraph: para.paragraph, glyphs: para.literal.replace(/\s/gu, '').length });
  }
  if (found.length !== 1) throw Error('The page-head text is missing or ambiguous');
  return found[0];
}

module.exports = (source, saved) => {
  const texts = [136, 138].map(index => source.getTextRange(0, index, 0, 0xffffffff));
  const original = texts.map(text => location(source, text));
  const output = texts.map(text => location(saved, text));
  const matches = original.every((point, index) => point.page === output[index].page && Math.abs(point.y - output[index].y) <= .2);
  return { original, output, matches };
};

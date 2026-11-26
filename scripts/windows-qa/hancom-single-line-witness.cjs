'use strict';

function placements(doc, literal, firstPage, lastPage) {
  const all = new Map();
  for (let page = firstPage - 1; page < Math.min(lastPage, doc.pageCount()); page++) {
    const paper = JSON.parse(doc.getPageInfo(page));
    for (const run of JSON.parse(doc.getPageTextLayout(page)).runs) {
      if (run.headerFooter || run.cellPath?.length || run.paraIdx === 0xffffffff || !run.text.length) continue;
      const key = `${run.secIdx}:${run.paraIdx}`;
      const item = all.get(key) ?? { paragraph: run.paraIdx, literal: '', lines: new Map(), withinBody: true };
      item.literal += run.text;
      const line = `${page + 1}:${run.y}`;
      if (!item.lines.has(line)) item.lines.set(line, { page: page + 1, y: run.y });
      item.withinBody &&= run.x + run.w <= paper.bodyRight + .16;
      all.set(key, item);
    }
  }
  const matches = [...all.values()].filter(item => item.literal === literal);
  if (matches.length !== 1) throw Error('The authored paragraph witness is missing or ambiguous');
  const { paragraph, literal: payload, lines, withinBody } = matches[0];
  return { paragraph, length: payload.length, lines: [...lines.values()], withinBody };
}

module.exports = (original, saved) => {
  const literal = original.getTextRange(0, 162, 0, 0xffffffff);
  const source = placements(original, literal, 9, 12);
  const output = placements(saved, literal, 9, 13);
  return { source, saved: output, matches: source.lines.length === 1 && output.lines.length === 1 &&
    output.withinBody && output.lines[0].page === source.lines[0].page &&
    Math.abs(output.lines[0].y - source.lines[0].y) <= .16 };
};

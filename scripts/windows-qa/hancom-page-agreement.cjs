'use strict';

const compact = value => value.normalize('NFC').replace(/\s/gu, '');

function owners(doc) {
  const body = new Map(), cell = new Map();
  for (let page = 0; page < doc.pageCount(); page++) {
    const { runs } = JSON.parse(doc.getPageTextLayout(page));
    for (const run of runs) {
      const text = compact(run.text);
      if (!text || run.headerFooter || !Number.isInteger(run.paraIdx) || run.paraIdx === 0xffffffff) continue;
      const map = run.cellPath?.length ? cell : body;
      const owner = JSON.stringify([run.secIdx, run.parentParaIdx, run.paraIdx, run.cellPath]);
      const entry = map.get(owner) ?? { text:'', page:page+1, x:run.x, y:run.y, lines:new Map() };
      const line = page + ':' + run.y;
      entry.lines.set(line, (entry.lines.get(line) ?? '') + text);
      entry.text += text;
      map.set(owner, entry);
    }
  }
  const index = map => {
    const unique = new Map();
    for (const value of map.values()) {
      if (value.text.length < 20) continue;
      unique.set(value.text, unique.has(value.text) ? null : value);
    }
    return unique;
  };
  return { body:index(body), cell:index(cell) };
}

function tally(reference, saved) {
  const counts = { fixed:0, matched:0, missing:0, ambiguous:0, sameStartPage:0, sameLineBreaks:0,
    sameStartPageAndLines:0, sameStartXYWithin1px:0 };
  for (const [literal, before] of reference) {
    if (!before) continue;
    counts.fixed++;
    const after = saved.get(literal);
    if (after === null) { counts.ambiguous++; continue; }
    if (!after) { counts.missing++; continue; }
    counts.matched++;
    const samePage = before.page === after.page;
    const a = [...before.lines.values()], b = [...after.lines.values()];
    const sameBreaks = a.length === b.length && a.every((line, i) => line === b[i]);
    counts.sameStartPage += Number(samePage);
    counts.sameLineBreaks += Number(sameBreaks);
    counts.sameStartPageAndLines += Number(samePage && sameBreaks);
    counts.sameStartXYWithin1px += Number(samePage && Math.abs(before.x-after.x) <= 1 && Math.abs(before.y-after.y) <= 1);
  }
  return counts;
}

module.exports = function pageAgreement(original, converted) {
  const a = owners(original), b = owners(converted);
  return { originalPages:original.pageCount(), convertedPages:converted.pageCount(),
    body:tally(a.body,b.body), cell:tally(a.cell,b.cell) };
};

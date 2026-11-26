'use strict';

function measure(doc, lastPage) {
  const owners = new Map();
  const pages = [];
  for (let page = 0; page < Math.min(lastPage, doc.pageCount()); page++) {
    const runs = JSON.parse(doc.getPageTextLayout(page)).runs;
    const controls = JSON.parse(doc.getPageControlLayout(page)).controls;
    let body = 0;
    let cells = 0;
    for (const run of runs) {
      if (run.headerFooter || run.paraIdx === 0xffffffff) continue;
      const literal = run.text.replace(/\s/gu, '');
      if (!literal) continue;
      const inCell = run.cellPath?.length > 0;
      const key = JSON.stringify([run.secIdx, run.parentParaIdx, run.paraIdx, run.cellPath]);
      const item = owners.get(key) ?? { paragraph: run.paraIdx, page: page + 1, x: run.x, y: run.y, literal: '', inCell };
      item.literal += literal;
      owners.set(key, item);
      if (inCell) cells += literal.length;
      else body += literal.length;
    }
    pages.push({ page: page + 1, body, cells, tables: controls.filter(c => c.type === 'table' && !c.headerFooter).length });
  }
  const unique = new Map();
  for (const item of owners.values()) {
    if (item.literal.length < 12) continue;
    const prior = unique.get(item.literal);
    unique.set(item.literal, prior === undefined ? item : null);
  }
  return { pages, unique };
}

module.exports = (original, saved) => {
  const source = measure(original, 14);
  const output = measure(saved, 19);
  const offsets = new Map();
  const boundaries = [];
  let previous = 0;
  for (const [literal, item] of source.unique) {
    if (!item || item.page > 12) continue;
    const after = output.unique.get(literal);
    if (!after || item.inCell !== after.inCell) continue;
    const bucket = offsets.get(item.page) ?? {};
    const delta = after.page - item.page;
    bucket[delta] = (bucket[delta] ?? 0) + 1;
    offsets.set(item.page, bucket);
    if (!item.inCell && delta !== previous && boundaries.length < 24) {
      previous = delta;
      const at = ({ paragraph, page, x, y }) => ({ paragraph, page, x, y });
      boundaries.push({ source: at(item), saved: at(after), length: literal.length });
    }
  }
  return { offsets: [...offsets], boundaries, sourcePages: source.pages, savedPages: output.pages };
};

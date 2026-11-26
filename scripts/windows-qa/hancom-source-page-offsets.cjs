'use strict';

function measure(doc, lastPage) {
  const owners = new Map();
  const tables = [];
  const pages = [];
  for (let page = 0; page < Math.min(lastPage, doc.pageCount()); page++) {
    const runs = JSON.parse(doc.getPageTextLayout(page)).runs;
    const controls = JSON.parse(doc.getPageControlLayout(page)).controls;
    let body = 0;
    let cells = 0;
    for (const run of runs) {
      if (run.headerFooter || !Number.isInteger(run.paraIdx) || run.paraIdx === 0xffffffff) continue;
      const literal = run.text.replace(/\s/gu, '');
      if (!literal) continue;
      const inCell = run.cellPath?.length > 0;
      const key = JSON.stringify([run.secIdx, run.parentParaIdx, run.paraIdx, run.cellPath]);
      const item = owners.get(key) ?? { section: run.secIdx, paragraph: run.paraIdx,
        page: page + 1, x: run.x, y: run.y, literal: '', inCell, lines: new Map() };
      item.literal += literal;
      if (!inCell) {
        const key = `${page + 1}:${run.y}`;
        const line = item.lines.get(key) ?? { page: page + 1, y: run.y, h: 0, glyphs: 0 };
        line.glyphs += literal.length;
        line.h = Math.max(line.h, run.h);
        item.lines.set(key, line);
      }
      owners.set(key, item);
      if (inCell) cells += literal.length;
      else body += literal.length;
    }
    const grid = controls.filter(c => c.type === 'table' && !c.headerFooter && !c.cellPath?.length);
    for (const table of grid) {
      const literal = runs.filter(run => run.parentParaIdx === table.paraIdx && run.secIdx === table.secIdx &&
        run.cellPath?.length === 1 && run.cellPath[0].controlIndex === table.controlIdx)
        .map(run => run.text.replace(/\s/gu, '')).join('');
      if (!literal) continue;
      const { outerTop, outerBottom, pageBreak } = JSON.parse(doc.getTableProperties(table.secIdx, table.paraIdx, table.controlIdx));
      tables.push({ paragraph: table.paraIdx, page: page + 1, rows: table.rowCount, columns: table.colCount,
        x: table.x, y: table.y, h: table.h, outerTop, outerBottom, pageBreak, literal });
    }
    pages.push({ page: page + 1, body, cells, tables: grid.length });
  }
  const unique = new Map();
  for (const item of owners.values()) {
    if (item.literal.length < 12) continue;
    const prior = unique.get(item.literal);
    unique.set(item.literal, prior === undefined ? item : null);
  }
  return { pages, unique, tables };
}

module.exports = (original, saved) => {
  const source = measure(original, 14);
  const output = measure(saved, 19);
  const offsets = new Map();
  const boundaries = [];
  const endLines = [];
  let previous = 0;
  for (const [literal, item] of source.unique) {
    if (!item || item.page > 12) continue;
    const after = output.unique.get(literal);
    if (!after || item.inCell !== after.inCell) continue;
    const bucket = offsets.get(item.page) ?? {};
    const delta = after.page - item.page;
    bucket[delta] = (bucket[delta] ?? 0) + 1;
    offsets.set(item.page, bucket);
    if (!item.inCell) {
      const at = (doc, { section, paragraph, page, x, y, lines }) => {
        const { pageBreakBefore, lineSpacing } = JSON.parse(doc.getParaPropertiesAt(section, paragraph));
        return { paragraph, page, x, y, lines: [...lines.values()], pageBreakBefore, lineSpacing };
      };
      if (delta !== previous && boundaries.length < 20) {
        previous = delta;
        boundaries.push({ source: at(original, item), saved: at(saved, after), length: literal.length });
      }
      if (item.page >= 7 && endLines.length < 24 &&
          (item.lines.size !== after.lines.size || new Set([...after.lines.values()].map(line => line.page)).size > 1)) {
        endLines.push({ source: at(original, item), saved: at(saved, after), length: literal.length });
      }
    }
  }
  const tablePairs = [];
  for (const table of source.tables.filter(value => value.page >= 7 && value.page <= 11)) {
    const candidates = output.tables.filter(value => value.literal === table.literal && value.columns === table.columns);
    if (candidates.length !== 1) continue;
    const cleaned = ({ literal, ...value }) => ({ ...value, glyphs: literal.length });
    tablePairs.push({ source: cleaned(table), saved: cleaned(candidates[0]) });
  }
  const originalTail = [...source.unique.values()].find(value => value && !value.inCell &&
    value.paragraph === 80 && value.page === 7 && value.literal.length === 60);
  const savedTail = originalTail && output.unique.get(originalTail.literal);
  const bodyTail = originalTail ? { source: [...originalTail.lines.values()],
    saved: savedTail ? [...savedTail.lines.values()] : [] } : undefined;
  return { offsets: [...offsets], boundaries, endLines, tablePairs, bodyTail, sourcePages: source.pages, savedPages: output.pages };
};

'use strict';

function locateBodyLines(document, literal) {
  let owner;
  const lines = [];
  for (let page = 0; page < document.pageCount(); page++) {
    const checked = new Map();
    for (const run of JSON.parse(document.getPageTextLayout(page)).runs) {
      if (run.cellPath?.length || !run.text.trim()) continue;
      const key = [run.secIdx, run.paraIdx].join(':');
      if (!checked.has(key)) checked.set(key,
        document.getTextRange(run.secIdx, run.paraIdx, 0, 0xffffffff) === literal);
      if (!checked.get(key)) continue;
      if (owner && owner !== key) throw Error('Ambiguous body text');
      owner = key;
      const last = lines.at(-1);
      if (!last || last.page !== page + 1 || Math.abs(last.y - run.y) > .25)
        lines.push({ page: page + 1, charStart: run.charStart, x: run.x, y: run.y });
    }
  }
  if (!lines.length) throw Error('Body text missing');
  return lines;
}

module.exports = (original, saved) => {
  const terminal = original.getTextRange(0, 62, 0, 0xffffffff);
  const next = original.getTextRange(0, 64, 0, 0xffffffff);
  const sourceLines = locateBodyLines(original, terminal);
  const savedLines = locateBodyLines(saved, terminal);
  const sourceNext = locateBodyLines(original, next)[0].page;
  const savedNext = locateBodyLines(saved, next)[0].page;
  if (sourceLines.length !== 2 || sourceLines[0].charStart !== 0 || sourceLines[1].charStart !== 79)
    throw Error('Source line-start witness differs from the expected original');
  return {
    source: { lines: sourceLines, nextPage: sourceNext },
    saved: { lines: savedLines, nextPage: savedNext },
    matches: sourceNext === savedNext && sourceLines.length === savedLines.length &&
      sourceLines.every((line, index) => {
        const actual = savedLines[index];
        return line.page === actual.page && line.charStart === actual.charStart &&
          Math.abs(line.x - actual.x) < .15 && Math.abs(line.y - actual.y) < .15;
      }),
  };
};

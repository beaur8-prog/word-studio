'use strict';

const compact = value => value.replace(/\s/gu, '');
const address = (section, paragraph, control, path = []) => JSON.stringify([section, paragraph, control, path]);

function snapshot(doc) {
  const tables = new Map();
  const body = [];
  let tablePlacements = 0;
  for (let page = 0; page < doc.pageCount(); page++) {
    const controls = JSON.parse(doc.getPageControlLayout(page)).controls;
    for (const table of controls) {
      if (table.type !== 'table' || table.headerFooter) continue;
      const key = address(table.secIdx, table.paraIdx, table.controlIdx, table.cellPath);
      const value = tables.get(key) ?? { section: table.secIdx, paragraph: table.paraIdx, control: table.controlIdx,
        path: table.cellPath, rows: table.rowCount, columns: table.colCount, pages: [], ink: new Map() };
      value.pages.push(page + 1);
      tables.set(key, value);
      tablePlacements++;
    }
    for (const run of JSON.parse(doc.getPageTextLayout(page)).runs) {
      if (run.headerFooter || !Number.isInteger(run.paraIdx) || run.paraIdx === 0xffffffff) continue;
      const text = compact(run.text);
      if (!text) continue;
      if (!run.cellPath?.length) {
        body.push(text);
        continue;
      }
      const path = run.cellPath, owner = path.at(-1);
      const key = address(run.secIdx, run.parentParaIdx, owner.controlIndex, path.slice(0, -1));
      const table = tables.get(key);
      if (!table) continue;
      const cell = JSON.stringify([owner.cellIndex, run.paraIdx]);
      table.ink.set(cell, (table.ink.get(cell) ?? '') + text);
    }
  }
  const groups = [...tables.values()].map(({ ink, ...table }) => ({
    ...table, texts: [...ink.values()], visible: [...ink.values()].join(''),
  }));
  return { pages: doc.pageCount(), body: body.join(''), groups, tablePlacements };
}

function count(values) {
  const result = new Map();
  for (const value of values) result.set(value, (result.get(value) ?? 0) + 1);
  return result;
}

function difference(before, after) {
  const removed = [], added = [];
  for (const [key, value] of before) if ((after.get(key) ?? 0) < value) removed.push([key, value - (after.get(key) ?? 0)]);
  for (const [key, value] of after) if ((before.get(key) ?? 0) < value) added.push([key, value - (before.get(key) ?? 0)]);
  return { removed, added };
}

module.exports = (baseline, candidate) => {
  const before = snapshot(baseline);
  const after = snapshot(candidate);
  const shape = table => JSON.stringify([table.rows, table.columns, table.visible]);
  const tableDiff = difference(count(before.groups.map(shape)), count(after.groups.map(shape)));
  const cellDiff = difference(count(before.groups.flatMap(table => table.texts)), count(after.groups.flatMap(table => table.texts)));
  const describe = (key, groups) => {
    const match = groups.filter(table => shape(table) === key);
    if (!match.length) return {};
    const { rows, columns, pages, path, paragraph, visible } = match[0];
    return { rows, columns, pages, nesting: path?.length ?? 0, paragraph, glyphs: visible.length };
  };
  const cells = diff => diff.map(([text, occurrences]) => ({ glyphs: text.length, occurrences }));
  return {
    totalPages: [before.pages, after.pages], placements: [before.tablePlacements, after.tablePlacements],
    physicalOwners: [before.groups.length, after.groups.length], bodySequenceExact: before.body === after.body,
    tableGlyphSequenceExact: before.groups.map(table => table.visible).join('') === after.groups.map(table => table.visible).join(''),
    sameCellMultiset: cellDiff.removed.length === 0 && cellDiff.added.length === 0,
    cellDiff: { removed: cells(cellDiff.removed).slice(0,20), added: cells(cellDiff.added).slice(0,20) },
    tableDiff: { removed: tableDiff.removed.map(([key, occurrences]) => ({ ...describe(key, before.groups), occurrences })),
      added: tableDiff.added.map(([key, occurrences]) => ({ ...describe(key, after.groups), occurrences })) },
  };
};

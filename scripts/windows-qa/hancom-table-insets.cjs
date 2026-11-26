'use strict';

const normalized = value => value.replace(/\s/gu, '');
const nonempty = values => values.filter(Boolean);

function tables(document, shapes) {
  const result = new Map();
  for (let index = 0; index < document.pageCount(); index++) {
    for (const table of JSON.parse(document.getPageControlLayout(index)).controls) {
      if (table.type !== 'table' || table.cellPath?.length ||
          !shapes.some(([rows, columns]) => table.rowCount === rows && table.colCount === columns)) continue;
      const key = [table.secIdx, table.paraIdx, table.controlIdx].join(':');
      const entry = result.get(key) ?? { table, pages: [], left: [] };
      entry.pages.push(index + 1);
      entry.left.push(table.x);
      result.set(key, entry);
    }
  }
  const hasAllCells = ({ table }) => {
    const cells = JSON.parse(document.getTableCellBboxes(table.secIdx, table.paraIdx, table.controlIdx));
    return new Set(cells.map(cell => cell.cellIdx)).size === table.rowCount * table.colCount;
  };
  return [...result.values()].filter(hasAllCells);
}
function paragraphs(document, { secIdx, paraIdx, controlIdx }, cell) {
  return nonempty(Array.from({ length: document.getCellParagraphCount(secIdx, paraIdx, controlIdx, cell) }, (_, index) =>
    normalized(document.getTextInCell(secIdx, paraIdx, controlIdx, cell, index, 0, 0xffffffff))));
}

module.exports = function hancomTableInsets(original, saved) {
  const source = tables(original, [[2, 2]]).filter(({ table }) => table.secIdx === 0 && table.paraIdx === 52);
  if (source.length !== 1 || source[0].pages.length !== 2) throw Error('Missing the source table spanning two sheets');
  const sourceCells = Array.from({ length: 4 }, (_, cell) => paragraphs(original, source[0].table, cell));
  const grid = tables(saved, [[2, 2], [1, 2]]);
  const lead = grid.filter(({ table }) => table.rowCount === 2 && [0, 1, 2].every(cell =>
    paragraphs(saved, table, cell).join('') === sourceCells[cell].join('')));
  if (lead.length !== 1 || lead[0].pages.length !== 1) throw Error('The first table segment was missing or ambiguous');
  const leadingCell = paragraphs(saved, lead[0].table, 3).join('');
  const sourceFullCell = sourceCells[3].join('');
  if (!leadingCell || !sourceFullCell.startsWith(leadingCell) || leadingCell === sourceFullCell)
    throw Error('The first segment changed the original cell order');
  const trailingCell = sourceFullCell.slice(leadingCell.length);
  const tail = grid.filter(({ table }) => table.rowCount === 1 &&
    !paragraphs(saved, table, 0).length && paragraphs(saved, table, 1).join('') === trailingCell);
  if (tail.length !== 1 || tail[0].pages.length !== 1) throw Error('The original table continuation was missing or ambiguous');
  const { secIdx, paraIdx, controlIdx } = lead[0].table;
  const { outerTop: top, outerBottom: bottom } = JSON.parse(saved.getTableProperties(secIdx, paraIdx, controlIdx));
  const blankIndex = (doc, table, cell, before, after) => {
    const { secIdx, paraIdx, controlIdx } = table;
    const key = [secIdx, paraIdx, controlIdx, cell];
    const literals = Array.from({ length:doc.getCellParagraphCount(...key) }, (_, i) => normalized(doc.getTextInCell(...key, i, 0, 0xffffffff)));
    const index = literals.findIndex((item, i) => !item && literals[i-1]===before && literals[i+1]===after);
    if (index < 0 || doc.getCellParagraphLength(...key, index)) throw Error('Missing original empty cell paragraph');
    return JSON.parse(doc.getCellCharPropertiesAt(...key, index, 0)).fontSize;
  };
  const sourceKey = [source[0].table.secIdx, source[0].table.paraIdx, source[0].table.controlIdx, 3];
  const witnesses = [10,12].map(i => normalized(original.getTextInCell(...sourceKey, i, 0, 0xffffffff)));
  const sourceBlank = blankIndex(original, source[0].table, 3, ...witnesses);
  const blank = blankIndex(saved, tail[0].table, 1, ...witnesses);
  return { blank, sourceBlank, continuationHeight:tail[0].table.h, sourcePage: source[0].pages[0], sourceContinuationPage: source[0].pages[1],
    page: lead[0].pages[0], continuationPage: tail[0].pages[0], top, bottom,
    sourceX: source[0].left[0], sourceContinuationX: source[0].left[1],
    x: lead[0].left[0], continuationX: tail[0].left[0],
    textLength: sourceCells.flat().reduce((total, text) => total + text.length, 0) };
};

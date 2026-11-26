'use strict';

const compact = text => text.replace(/\s/gu, '');

function cellLiteral(doc, key, cell) {
  const owner = [...key, cell];
  return compact(Array.from({length:doc.getCellParagraphCount(...owner)}, (_, p) =>
    doc.getTextInCell(...owner, p, 0, 0xffffffff)).join(''));
}

function firstFragment(doc, key, targetCells) {
  const pieces = JSON.parse(doc.getTableCellBboxes(...key));
  const page = Math.min(...pieces.map(cell => cell.pageIndex));
  const rows = new Map(pieces.filter(cell => targetCells.includes(cell.cellIdx) && cell.pageIndex === page)
    .map(cell => [cell.cellIdx, cell]));
  if (rows.size !== targetCells.length) throw Error('The two physical cells did not share the first table slice');
  const runs = JSON.parse(doc.getPageTextLayout(page)).runs.filter(run => run.secIdx === key[0] &&
    run.parentParaIdx === key[1] && run.cellPath?.length === 1 && run.cellPath[0].controlIndex === key[2]);
  const visible = targetCells.map(cell => compact(runs
    .filter(run => run.cellPath[0].cellIndex === cell).map(run => run.text).join('')));
  const table = JSON.parse(doc.getPageControlLayout(page)).controls.find(c =>
    c.type === 'table' && c.secIdx === key[0] && c.paraIdx === key[1] && c.controlIdx === key[2] && !c.cellPath?.length);
  if (!table) throw Error('The source table did not paint');
  return { page: page + 1, x: table.x, y: table.y, height: rows.get(targetCells[0]).h, visible };
}

module.exports = (original, saved) => {
  const reference = [0, 69, 0], cells = [9, 10];
  const prefix = [0,1,2,3].map(index => cellLiteral(original, reference, index));
  const source = firstFragment(original, reference, cells);
  if (source.page !== 5 || source.visible[0].length !== 4 || !source.visible[1])
    throw Error('The pinned source compact row does not match its verified layout');
  const matches = [];
  const seen = new Set();
  for (let page = 0; page < saved.pageCount(); page++) for (const table of JSON.parse(saved.getPageControlLayout(page)).controls) {
    const key = [table.secIdx, table.paraIdx, table.controlIdx];
    const identity = key.join(':');
    if (seen.has(identity) || table.type !== 'table' || table.cellPath?.length ||
        table.rowCount !== 5 || table.colCount !== 3) continue;
    seen.add(identity);
    try {
      if (prefix.every((text,index) => text === cellLiteral(saved, key, index))) matches.push(key);
    } catch { /* Reported child controls can share a table's visual shape. */ }
  }
  if (matches.length !== 1) throw Error('The saved first table slice was missing or ambiguous');
  const actual = firstFragment(saved, matches[0], cells);
  const sourceTail = JSON.parse(original.getTableCellBboxes(...reference))
    .filter(box => box.pageIndex === source.page && box.col === 2 && box.rowSpan === 1)
    .sort((a,b) => a.row - b.row);
  if (sourceTail.length !== 4) throw Error('The original tail grid changed');
  const lastLiteral = cellLiteral(original, reference, sourceTail.at(-1).cellIdx);
  const tails = [];
  for (let page = 0; page < saved.pageCount(); page++) for (const item of JSON.parse(saved.getPageControlLayout(page)).controls) {
    if (item.type !== 'table' || item.rowCount !== 4 || item.colCount !== 3 || item.cellPath?.length) continue;
    const key = [item.secIdx, item.paraIdx, item.controlIdx];
    try {
      const boxes = JSON.parse(saved.getTableCellBboxes(...key)).filter(box => box.pageIndex === page &&
        box.col === 2 && box.rowSpan === 1).sort((a,b) => a.row - b.row);
      if (boxes.length === 4 && cellLiteral(saved, key, boxes.at(-1).cellIdx) === lastLiteral)
        tails.push({page:page + 1, heights:boxes.map(box => box.h)});
    } catch { /* Ignore child controls differing from the pinned table. */ }
  }
  if (tails.length !== 1) throw Error('The saved tail grid was missing or ambiguous');
  const originalTail = sourceTail.map(box => box.h);
  const tail = tails[0];
  const geometry = ({visible, ...position}) => ({...position, glyphs:visible.map(value=>value.length)});
  return {source:geometry(source), saved:geometry(actual), tail:{ original:originalTail, ...tail },
    tailMatches:tail.page === source.page + 1 && tail.heights.every((height,i) => Math.abs(height - originalTail[i]) < .3),
    matches:actual.page === source.page && Math.abs(actual.y - source.y) <= .15 &&
      Math.abs(actual.height - source.height) <= .15 &&
      actual.visible.every((literal, index) => literal === source.visible[index])};
};

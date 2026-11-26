'use strict';

const compact = value => value.replace(/\s/gu,'');
const size = (doc, ref) => {
  const {rowCount,colCount} = JSON.parse(doc.getTableDimensions(...ref));
  return [rowCount,colCount];
};
function printedTables(doc) {
  const tables = [];
  for (let page = 0; page < doc.pageCount(); page++) {
    const shapes = JSON.parse(doc.getPageControlLayout(page)).controls;
    let ink;
    for (const table of shapes) {
      if (table.type !== 'table' || table.cellPath?.length || table.headerFooter) continue;
      const address = [table.secIdx, table.paraIdx, table.controlIdx];
      if (!((table.rowCount === 1 && table.colCount === 1) ||
           (table.rowCount === 12 && table.colCount === 4))) continue;
      ink ??= JSON.parse(doc.getPageTextLayout(page)).runs;
      const literal = compact(ink.filter(run => run.secIdx === table.secIdx && run.parentParaIdx === table.paraIdx &&
        run.cellPath?.length === 1 && run.cellPath[0].controlIndex === table.controlIdx)
        .map(run => run.text).join(''));
      if (literal) tables.push({address,page:page+1,x:table.x,y:table.y,h:table.h,literal});
    }
  }
  return tables;
}
module.exports = (original, saved) => {
  const prior = printedTables(original), next = printedTables(saved);
  const label = prior.find(item => item.page === 10 && item.address[1] === 109);
  const target = prior.find(item => item.page === 10 && item.address[1] === 111);
  if (!label || !target || label.literal.length !== 39 || target.literal.length !== 1149)
    throw Error('The original table test signature changed');
  const find = item => {
    const same = next.filter(candidate => candidate.literal === item.literal);
    if (same.length !== 1 || String(size(saved,same[0].address)) !== String(size(original,item.address)))
      throw Error('Missing or ambiguous saved table');
    return same[0];
  };
  const title = find(label), body = find(target);
  const info = JSON.parse(saved.getPageInfo(body.page - 1));
  const cleaned = ({address,literal,...item}) => ({...item,glyphs:literal.length});
  return {source:cleaned(target), saved:cleaned(body), title:cleaned(title),
    matches: title.page === label.page && body.page === target.page && Math.abs(body.y-target.y) <= 2 &&
      body.y > title.y + title.h && body.y + body.h <= info.footerArea.y + .15};
};

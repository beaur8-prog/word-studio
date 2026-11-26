'use strict';

function measure(doc, pages) {
  const matches = [];
  for (const page of pages) {
    const paper = JSON.parse(doc.getPageInfo(page - 1));
    const runs = JSON.parse(doc.getPageTextLayout(page - 1)).runs;
    const tables = JSON.parse(doc.getPageControlLayout(page - 1)).controls.filter(table =>
      table.type === 'table' && !table.headerFooter && !table.cellPath?.length);
    for (const table of tables) {
      const section = table.secIdx, para = table.paraIdx, control = table.controlIdx;
      const tableRuns = runs.filter(run => run.secIdx === section && run.parentParaIdx === para &&
        run.cellPath?.[0]?.controlIndex === control);
      const text = tableRuns.map(run => run.text.replace(/\s/gu, '')).join('');
      const properties = JSON.parse(doc.getTableProperties(section, para, control));
      const paragraph = JSON.parse(doc.getParaPropertiesAt(section, para));
      const line = JSON.parse(doc.getParagraphLineMetrics(section, para, '[]'));
      matches.push({ section, para, control, page, width: table.w, height: table.h, x: table.x, y: table.y,
        shape: [table.rowCount, table.colCount], text, bottomRoom: paper.footerArea.y - table.y - table.h,
        table: { treatAsChar: properties.treatAsChar, horzRelTo: properties.horzRelTo, vertRelTo: properties.vertRelTo,
          horzOffset: properties.horzOffset, vertOffset: properties.vertOffset, left: properties.outerLeft,
          right: properties.outerRight, top: properties.outerTop, bottom: properties.outerBottom,
          split: properties.pageBreak, headers: properties.repeatHeader },
        rows: [...new Set((table.cells ?? []).map(cell => cell.row))].length,
        paragraph: { break: paragraph.pageBreakBefore, marginLeft: paragraph.marginLeft, marginRight: paragraph.marginRight,
          before: paragraph.spacingBefore, after: paragraph.spacingAfter, indent: paragraph.indent,
          pitch: paragraph.lineSpacing, height: paragraph.fontLineHeight, stored: !!line.stored,
          first: line.lines?.slice(0,1) },
      });
    }
  }
  return matches;
}

module.exports = (original, saved, hwpReencoded) => {
  const source = measure(original, [12]);
  if (source.length !== 3 || !source.some(table => table.shape[0] === 8 && table.text.length === 1080))
    throw Error('The known source table fragment did not match');
  const output = measure(saved, [12,13,14,15]);
  const roundTrip = hwpReencoded && measure(hwpReencoded, [12,13,14,15]);
  const safe = ({text, ...table}) => ({...table, glyphs: text.length});
  const locate = (table, group) => {
    const equal = group.filter(item => item.text === table.text && item.shape[1] === table.shape[1]);
    return equal.length === 1 ? safe(equal[0]) : { differentFragments: group.filter(item =>
      item.shape[1] === table.shape[1] && item.text.length > 0).map(safe) };
  };
  return source.map(table => ({source: safe(table), saved: locate(table, output),
    binaryReencodedAsPackage: roundTrip && locate(table, roundTrip)}));
};

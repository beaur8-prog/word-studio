'use strict';

const { unzipSync, zipSync, strFromU8, strToU8 } = require('fflate');
const measure = require('./hancom-table-fragment-position.cjs');
const prefix = /^<hp:sz\b[^>]*\/><hp:pos\b[^>]*\/><hp:outMargin\b[^>]*\/>/;

function painted(doc) {
  const paragraphs = new Map();
  const shapes = [];
  for (let page = 0; page < doc.pageCount(); page++) {
    for (const run of JSON.parse(doc.getPageTextLayout(page)).runs) {
      if (run.headerFooter || !Number.isInteger(run.paraIdx) || run.paraIdx === 0xffffffff) continue;
      const text = run.text.replace(/\s/gu, '');
      if (!text) continue;
      const owner = JSON.stringify([run.secIdx, run.paraIdx, run.parentParaIdx, run.cellPath]);
      paragraphs.set(owner, (paragraphs.get(owner) ?? '') + text);
    }
    for (const shape of JSON.parse(doc.getPageControlLayout(page)).controls) if (!shape.headerFooter)
      shapes.push(JSON.stringify([shape.type, shape.secIdx, shape.paraIdx, shape.controlIdx, shape.w, shape.h, shape.cellPath]));
  }
  return { paragraphs: JSON.stringify([...paragraphs].sort()), shapes: JSON.stringify(shapes.sort()) };
}

module.exports = (Core, original, saved, elementRanges) => {
  const source = measure(original, saved);
  const prior = source.filter(item => item.source.shape[1] === 1 && item.saved.page === 12);
  const long = source.find(item => item.source.glyphs === 1080);
  if (prior.length !== 2 || long?.saved?.page !== 13) throw Error('Expected two labels and the displaced native table');
  const baseline = painted(saved);
  const archive = unzipSync(saved.exportHwpx());
  const cases = [prior.slice(0, 1), prior.slice(1), prior];
  return cases.map(items => {
    const copies = { ...archive };
    const bySection = new Map();
    for (const record of items) {
      const { section } = record.saved;
      if (!bySection.has(section)) bySection.set(section, []);
      bySection.get(section).push(record);
    }
    for (const [section, items] of bySection) {
      const name = 'Contents/section' + section + '.xml';
      let text = strFromU8(copies[name]);
      const paragraphs = elementRanges(text, 'hp:p');
      for (const record of [...items].sort((a,b) => b.saved.para - a.saved.para)) {
        const p = paragraphs[record.saved.para];
        const tables = elementRanges(p.inner, 'hp:tbl');
        if (tables.length !== 1 || !prefix.test(tables[0].inner)) throw Error('Ambiguous label table markup');
        const t = tables[0], tag = prefix.exec(t.inner)[0];
        const before = 'bottom="' + record.saved.table.bottom + '"';
        const after = 'bottom="' + record.source.table.bottom + '"';
        const changed = tag.replace(/<hp:outMargin\b[^>]*\/>/, margin => margin.replace(before, after));
        if (changed === tag) throw Error('The authenticated outer margin was not in the package');
        const at = p.start + p.open.length + t.start + t.open.length;
        text = text.slice(0, at) + changed + text.slice(at + tag.length);
      }
      copies[name] = strToU8(text);
    }
    const candidate = new Core(zipSync(copies,{level:1}));
    try {
      const results = measure(original, candidate);
      const savedText = painted(candidate);
      return { margins: items.map(item => [item.saved.para,item.saved.table.bottom,item.source.table.bottom]),
        pages: candidate.pageCount(), tables: results.map(item => {
          const { page, x, y, width, height, glyphs, bottomRoom } = item.saved;
          return { page, x, y, width, height, glyphs, bottomRoom };
        }), sameOwnerText: baseline.paragraphs === savedText.paragraphs, sameShapes: baseline.shapes === savedText.shapes };
    } finally { candidate.free(); }
  });
};

module.exports.placement = (Core, original, saved, elementRanges) => {
  const prior = measure(original, saved);
  const target = prior.find(item => item.source.glyphs === 1080 && item.saved.page === 13);
  if (!target) throw Error('Expected native displaced table');
  const archive = unzipSync(saved.exportHwpx());
  const position = {
    paperExact: { vertRelTo: 'PAPER', horzRelTo: 'PAPER',
      vertOffset: Math.round(target.source.y * 75) - target.source.table.top,
      horzOffset: Math.round(target.source.x * 75), flowWithText: 0 },
  };
  const name = 'Contents/section' + target.saved.section + '.xml';
  const section = strFromU8(archive[name]);
  const para = elementRanges(section, 'hp:p')[target.saved.para];
  const tables = elementRanges(para.inner, 'hp:tbl');
  if (tables.length !== 1 || !prefix.test(tables[0].inner)) throw Error('Ambiguous displaced native table markup');
  const table = tables[0], before = prefix.exec(table.inner)[0];
  const offset = para.start + para.open.length + table.start + table.open.length;
  const baseline = painted(saved);
  const set = (tag, values) => Object.entries(values).reduce((result,[name,value]) => {
    const field = new RegExp('\\b' + name + '="[^"]*"');
    if (!field.test(result)) throw Error('A required native placement field is missing');
    return result.replace(field, name + '="' + value + '"');
  }, tag);
  return Object.entries(position).map(([mode, anchor]) => {
    const attributes = { treatAsChar: 0, affectLSpacing: 0, allowOverlap: 0, holdAnchorAndSO: 0,
      vertAlign: 'TOP', horzAlign: 'LEFT', ...anchor };
    const marked = before.replace(/<hp:pos\b[^>]*\/>/, value => set(value, attributes))
      .replace(/<hp:outMargin\b[^>]*\/>/, value => set(value, Object.fromEntries(['top','bottom','left','right'].map(side => [side,target.source.table[side]]))));
    const candidate = new Core(zipSync({ ...archive,
      [name]: strToU8(section.slice(0, offset) + marked + section.slice(offset + before.length)) }, { level: 1 }));
    try {

      const group = measure(original, candidate);
      const next = painted(candidate);
      const observed = group.find(item => item.source.glyphs === 1080)?.saved;
      const pageIndex = observed?.page - 1;
      const paper = Number.isInteger(pageIndex) ? JSON.parse(candidate.getPageInfo(pageIndex)) : null;
      const runs = paper ? JSON.parse(candidate.getPageTextLayout(pageIndex)).runs.filter(run => !run.headerFooter &&
        run.secIdx === observed.section && run.parentParaIdx === observed.para && run.cellPath?.[0]?.controlIndex === observed.control) : [];
      const outside = [], outsideVisible = [];
      const content = runs.filter(run => run.text.replace(/\s/gu, ''));
      for (const run of runs) {
        const spills = [paper.bodyLeft - run.x, run.x + run.w - paper.bodyRight,
          paper.marginTop + paper.marginHeader - run.y, run.y + run.h - paper.footerArea.y].map(value => Math.round(value * 100) / 100);
        const end = run.text.trimEnd().length;
        const advance = run.charX?.[end];
        const rightVisible = run.x + (end < run.text.length && Number.isFinite(advance) ? advance : run.w) - paper.bodyRight;
        if (end && Math.max(spills[0], rightVisible, spills[2], spills[3]) > .15) {
          outsideVisible.push({rightVisible, cell:run.cellPath[0].cellIndex});
        }
        if (spills.some(value => value > .15)) outside.push({spills, nonspace:run.text.replace(/\s/gu,'').length,
          trailing:run.text.length-end,rightVisible,y:run.y,h:run.h,font:run.fontSize,cell:run.cellPath[0].cellIndex});
      }
      return { mode, pages: candidate.pageCount(), runCount:runs.length, bodyGlyphRuns:content.length, outside, outsideVisible, tableOrdinals: group.map(item => {
        const { page, x, y, height, glyphs, bottomRoom } = item.saved;
        return { page, x, y, height, glyphs, bottomRoom };
      }), sameOwnerText: baseline.paragraphs === next.paragraphs, sameShapes: baseline.shapes === next.shapes };
    } finally { candidate.free(); }
  });
};

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

const fs = require('node:fs/promises');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
let pending;
module.exports = () => pending ??= (async () => {
  const dir = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(dir + '/rhwp.js'));
  await init({ module_or_path: await fs.readFile(dir + '/rhwp_bg.wasm') });
  const samples = {};
  for (const nested of [false, true]) {
    const source = HwpDocument.createEmpty();
    try {
      source.createBlankDocument();
      const paragraph = (text, margin) => `<p style="font-family:Arial;font-size:12pt;line-height:20px;margin-top:${margin}px;margin-bottom:${margin ? 4 : 0}px">${text}</p>`;
      const table = `<table style="width:400px"><tr><td style="padding:0;vertical-align:top">${paragraph('FIRST', 0)}${paragraph('SPACED 한글', 12)}${paragraph('LAST', 0)}</td></tr></table>`;
      source.pasteHtml(0, 0, 0, '<p>LEAD</p>' + (nested ? `<table><tr><td>${table}</td></tr></table>` : table));
      for (const format of ['hwp', 'hwpx']) {
        const bytes = format === 'hwp' ? source.exportHwp() : source.exportHwpx();
        const doc = new HwpDocument(bytes);
        try {
          let run;
          for (let i = 0; i < doc.pageCount() && !run; i++) run = JSON.parse(doc.getPageTextLayout(i)).runs.find(r => r.text?.includes('SPACED'));
          if (!run?.cellPath?.length) throw Error('Fixture must expose an unambiguous native cell path');
          // The fixture has one source body table, and (when nested) one
          // inner table in its outer cell's first control-only paragraph.
          // Native paint hints omit that ancestor; query the source address.
          const outer = JSON.parse(doc.getControls()).find(c => c.list === 0 && c.ctrlId === 'tbl');
          const cellPath = nested ? [{ controlIndex: outer.controlIndex, cellIndex: 0, cellParaIndex: 0 }, { controlIndex: 0, cellIndex: 0, cellParaIndex: 1 }]
            : [{ controlIndex: outer.controlIndex, cellIndex: 0, cellParaIndex: 1 }];
          const properties = JSON.parse(doc.getCellParaPropertiesAtByPath(0, outer.para, JSON.stringify(cellPath)));
          if (Math.abs(properties.spacingBefore / 2 - 12) > .01 || Math.abs(properties.spacingAfter / 2 - 4) > .01)
            throw Error('Native source must retain the 12px/4px paragraph spacing: ' + JSON.stringify(properties));
          samples[`${format}-${nested}`] = { bytes, properties, path: cellPath };
        } finally { doc.free(); }
      }
    } finally { source.free(); }
  }
  return samples;
})();

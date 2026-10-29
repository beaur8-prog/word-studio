const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url');
const { unzipSync, zipSync, strFromU8, strToU8 } = require('fflate');
module.exports = async function createFixture(directory) {
  const root = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(root, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(root, 'rhwp_bg.wasm')) });
  const result = {};
  for (const mode of ['rows', 'merged', 'deep', 'grow', 'gap-zero', 'gap-large']) {
    let doc = HwpDocument.createEmpty();
    try {
      doc.createBlankDocument();
      const cell = (text, height, extra = '') => `<td style="width:200px;height:${height}px;padding:0;border:none" ${extra}><p style="margin:0;line-height:20px">${text}</p></td>`;
      const first = mode === 'grow' ? 'TALL A' + '<br>GROW'.repeat(8) : 'TALL A';
      const table = `<table style="width:400px"><tr>${mode === 'merged' ? cell('MERGED', 168, 'rowspan="2"') : cell(first, mode === 'grow' ? 12 : 96)}${cell('TALL B', 96)}</tr><tr>${cell('SECOND', 72, mode === 'merged' ? '' : 'colspan="2"')}</tr></table>`;
      const inner = mode === 'deep' ? `<table style="width:440px"><tr><td style="height:260px;padding:0;border:none"><p>MIDDLE</p>${table}<p>MIDDLE END</p></td></tr></table>` : table;
      doc.pasteHtml(0, 0, 0, `<table style="width:500px"><tr><td style="height:400px;padding:0;border:none"><p>BEFORE</p>${inner}<p>AFTER</p></td></tr></table>`);
      if (mode.startsWith('gap-')) {
        const entries = unzipSync(doc.exportHwpx()), key = 'Contents/section0.xml', xml = strFromU8(entries[key]);
        const start = [...xml.matchAll(/<hp:tbl\b/g)][1].index;
        const tail = xml.slice(start).replace(/<hp:outMargin\b[^>]*\/>/, `<hp:outMargin left="0" right="0" top="0" bottom="${mode === 'gap-large' ? 1800 : 0}"/>`);
        entries[key] = strToU8(xml.slice(0, start) + tail);
        doc.free(); doc = new HwpDocument(zipSync(entries));
      }
      const outer = JSON.parse(doc.getControls()).find(c => c.list === 0 && c.ctrlId === 'tbl');
      const tablePath = [{ controlIndex: outer.controlIndex, cellIndex: 0, cellParaIndex: 1 },
        ...(mode === 'deep' ? [{ controlIndex: 0, cellIndex: 0, cellParaIndex: 1 }] : []),
        { controlIndex: 0, cellIndex: 0, cellParaIndex: 0 }];
      for (const format of ['hwp', 'hwpx']) {
        const file = path.join(directory, mode + '.' + format), bytes = format === 'hwp' ? doc.exportHwp() : doc.exportHwpx();
        await fs.writeFile(file, bytes);
        // Query the serialized input, including each format's native defaults.
        const loaded = new HwpDocument(bytes);
        try {
          const native = JSON.parse(loaded.getTableCellBboxesByPath(0, 0, JSON.stringify(tablePath)));
          if (native.length !== 3 || native.some(c => c.pageIndex !== 0)) throw Error('Invalid one-page nested-height source');
          const host = JSON.parse(loaded.getParagraphLineMetrics(0, 0, JSON.stringify(tablePath.slice(0, -1))));
          const spacing = host.tables[0].outerBottom + host.lines[0].spacing;
          result[mode + '-' + format] = { file, native, spacing };
        } finally { loaded.free(); }
      }
    } finally { doc.free(); }
  }
  return result;
};

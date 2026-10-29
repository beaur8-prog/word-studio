const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url');
module.exports = async function fixture(directory) {
  await fs.mkdir(directory, { recursive: true });
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  const result = {};
  const tail = ' small words make a second line with smaller letters';
  for (const columns of [1, 2]) for (const rule of ['percent', 'fixed']) {
    const doc = HwpDocument.createEmpty();
    try {
      doc.createBlankDocument();
      const html = `<p style="font-family:Arial;font-size:18pt;line-height:${rule === 'percent' ? '150%' : '36px'};margin:0"><span style="font-family:Arial;font-size:18pt">1.</span><span style="font-family:Arial;font-size:10pt">${tail}</span></p><p style="font-family:Arial;font-size:10pt;line-height:20px;margin:0">AFTER</p>`;
      doc.pasteHtml(0, 0, 0, '<p></p>' + html + '<p></p>');
      doc.deleteParagraph(0, doc.getParagraphCount(0) - 1); doc.deleteParagraph(0, 0);
      const width = columns === 1 ? 22500 : 43500;
      doc.setPageDef(0, JSON.stringify({ width, height: 45000, landscape: false, marginTop: 1500, marginBottom: 1500, marginLeft: 1500, marginRight: 1500, marginHeader: 0, marginFooter: 0, marginGutter: 0 }));
      if (columns === 2) doc.setColumnDef(0, 2, 0, 1, 1500);
      for (const format of ['hwp', 'hwpx']) {
        const bytes = format === 'hwp' ? doc.exportHwp() : doc.exportHwpx(), loaded = new HwpDocument(bytes);
        try {
          const metrics = JSON.parse(loaded.getParagraphLineMetrics(0, 0, '[]'));
          const runs = JSON.parse(loaded.getPageTextLayout(0)).runs.filter(r => r.parentParaIdx === undefined);
          const first = runs.find(r => r.paraIdx === 0), after = runs.find(r => r.paraIdx === 1 && r.text.includes('AFTER'));
          const key = `${columns}-${rule}-${format}`, file = path.join(directory, key + '.' + format);
          const lineTops = [...new Set(runs.filter(r => r.paraIdx === 0 && r.text.trim()).map(r => r.y))];
          if (!first || !after || lineTops.length !== 2) throw Error('Mixed body fixture requires two native visual lines on one page: ' + key);
          result[key] = { file, format, columns, rule, text: '1.' + tail, nativeAdvance: after.y - first.y,
            nativeLines: metrics.stored ? metrics.lines.map(l => ({ height: l.height, spacing: l.spacing })) : lineTops.map((y, i) => ({ top: y, advance: (lineTops[i + 1] ?? after.y) - y })), properties: JSON.parse(loaded.getParaPropertiesAt(0, 0)) };
          await fs.writeFile(file, bytes);
        } finally { loaded.free(); }
      }
    } finally { doc.free(); }
  }
  for (const edge of ['Page', 'Paper']) {
    const doc = HwpDocument.createEmpty();
    try {
      doc.createBlankDocument();
      doc.pasteHtml(0, 0, 0, '<p style="font-size:10pt;line-height:20px">BEFORE</p><table style="width:220px"><tr><td style="width:220px;height:72px;padding:0"><p style="font-size:10pt;line-height:20px">BOTTOM OVERLAY</p></td></tr></table><p style="font-size:10pt;line-height:20px">AFTER</p>');
      doc.setPageDef(0, JSON.stringify({ width: 42520, height: 51024, landscape: false, marginTop: 4500, marginBottom: 4500, marginLeft: 4500, marginRight: 4500, marginHeader: 0, marginFooter: 0, marginGutter: 0 }));
      const t = JSON.parse(doc.getControls()).find(c => c.list === 0 && c.ctrlId === 'tbl');
      doc.setTableProperties(0, t.para, t.controlIndex, JSON.stringify({ repeatHeader: false, pageBreak: 2, treatAsChar: false,
        textWrap: 'InFrontOfText', vertRelTo: edge, vertAlign: 'Bottom', vertOffset: 1500, horzRelTo: 'Paper', horzAlign: 'Left', horzOffset: 4500, outerTop: 0, outerBottom: 0 }));
      for (const format of ['hwp', 'hwpx']) {
        const bytes = format === 'hwp' ? doc.exportHwp() : doc.exportHwpx(), loaded = new HwpDocument(bytes);
        try {
          const position = JSON.parse(loaded.getPageControlLayout(0)).controls.find(c => c.type === 'table');
          const key = `${edge}-${format}`, file = path.join(directory, key + '.' + format);
          result[key] = { file, format, edge, nativePosition: position }; await fs.writeFile(file, bytes);
        } finally { loaded.free(); }
      }
    } finally { doc.free(); }
  }
  return result;
};

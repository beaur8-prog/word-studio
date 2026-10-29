const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url');
module.exports = async function createFixture(directory) {
  const root = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(root, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(root, 'rhwp_bg.wasm')) });
  const result = {};
  for (const nested of [false, true]) for (const mode of ['short', 'wide', 'mixed']) {
    const source = HwpDocument.createEmpty();
    try {
      source.createBlankDocument();
      const text = mode === 'wide' ? 'Hgyp TEST '.repeat(6) : 'Hgyp TEST';
      const runs = mode === 'mixed' ? `<span style="font-family:Arial;font-size:16pt">BIG </span><span style="font-family:Arial;font-size:10pt">${text}</span>` : `<span style="font-family:Arial;font-size:10pt">${text}</span>`;
      const table = `<table style="width:480px"><tr><td style="width:480px;height:64px;padding:0;border:none"><p style="font-family:Arial;font-size:10pt;line-height:160%">${runs}</p></td></tr></table>`;
      source.pasteHtml(0, 0, 0, nested ? `<table style="width:540px"><tr><td style="height:650px;padding:0;border:none"><p>OUTER</p>${table}</td></tr></table>` : table);
      for (const format of ['hwp', 'hwpx']) {
        const key = `${nested ? 'nested' : 'ordinary'}-${mode}-${format}`;
        const file = path.join(directory, key + '.' + format), bytes = format === 'hwp' ? source.exportHwp() : source.exportHwpx();
        await fs.writeFile(file, bytes);
        const native = new HwpDocument(bytes);
        try {
          const outer = JSON.parse(native.getControls()).find(c => c.list === 0 && c.ctrlId === 'tbl');
          const cells = [{ controlIndex: outer.controlIndex, cellIndex: 0, cellParaIndex: nested ? 1 : 0 },
            ...(nested ? [{ controlIndex: 0, cellIndex: 0, cellParaIndex: 0 }] : [])];
          const metrics = JSON.parse(native.getParagraphLineMetrics(0, 0, JSON.stringify(cells)));
          const char = JSON.parse(native.getCellCharPropertiesAtByPath(0, 0, JSON.stringify(cells), 0));
          const properties = JSON.parse(native.getCellParaPropertiesAtByPath(0, 0, JSON.stringify(cells)));
          if (metrics.lines.length !== 1 || properties.lineSpacingType !== 'Percent' || properties.lineSpacing !== 160)
            throw Error('Fixture must be one native line with 160% spacing');
          const runs = JSON.parse(native.getPageTextLayout(0)).runs.filter(run => /BIG|Hgyp TEST/.test(run.text));
          result[key] = { file, metrics, char, properties, inkHeight: Math.max(...runs.map(run => run.h)), fontPx: char.fontSize / 75,
            advance: metrics.lines[0].height + metrics.lines[0].spacing };
        } finally { native.free(); }
      }
    } finally { source.free(); }
  }
  return result;
};

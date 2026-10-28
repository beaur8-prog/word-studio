const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url');
module.exports = async function createFixture(directory) {
  const root = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(root, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(root, 'rhwp_bg.wasm')) });
  const doc = HwpDocument.createEmpty(), files = {};
  try {
    doc.createBlankDocument();
    const text = 'Long explanation should occupy the wide final column without pushing the form onto another page. '.repeat(4);
    doc.pasteHtml(0, 0, 0, `<table style="width:560px"><tr><td style="padding:0"><p>OUTER</p>
      <table style="width:500px"><tr><td colspan="2" style="width:100px;padding:0;vertical-align:middle"><p>MERGED LABEL</p></td><td style="width:400px;padding:0"><p>${text}</p></td></tr>
      <tr><td style="width:40px;padding:0"><p>A</p></td><td style="width:60px;padding:0"><p>B</p></td><td style="width:400px;padding:0;vertical-align:bottom"><p>BOTTOM</p>
      <table style="width:300px"><tr><td style="width:60px;padding:0"><p>DEEP</p></td><td style="width:240px;padding:0"><p>DEEP VALUE</p></td></tr></table></td></tr></table><p>TAIL</p></td></tr></table>`);
    const outer = JSON.parse(doc.getControls()).find(c => c.list === 0 && c.ctrlId === 'tbl');
    const geometry = [];
    for (const cellPath of [
      [{ controlIndex: outer.controlIndex, cellIndex: 0, cellParaIndex: 1 }, { controlIndex: 0, cellIndex: 0, cellParaIndex: 0 }],
      [{ controlIndex: outer.controlIndex, cellIndex: 0, cellParaIndex: 1 }, { controlIndex: 0, cellIndex: 4, cellParaIndex: 1 }, { controlIndex: 0, cellIndex: 0, cellParaIndex: 0 }],
    ]) {
      const json = JSON.stringify(cellPath), dims = JSON.parse(doc.getTableDimensionsByPath(0, 0, json)), cells = [];
      for (let index = 0; index < dims.cellCount; index++) cells.push({
        ...JSON.parse(doc.getCellInfoByPath(0, 0, JSON.stringify([...cellPath.slice(0, -1), { ...cellPath.at(-1), cellIndex: index }]))),
        ...JSON.parse(doc.getCellPropertiesByPath(0, 0, json, index)),
      });
      geometry.push(cells);
    }
    for (const format of ['hwp', 'hwpx']) {
      files[format] = path.join(directory, 'nested-grid.' + format);
      await fs.writeFile(files[format], format === 'hwp' ? doc.exportHwp() : doc.exportHwpx());
    }
    return { files, geometry };
  } finally { doc.free(); }
};

const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url'), { createRequire } = require('node:module'), ts = require('typescript');
const { unzipSync, zipSync, strToU8, strFromU8 } = require('fflate');
function load(file, cache = new Map()) {
  if (cache.has(file)) return cache.get(file).exports;
  const module = { exports: {} }; cache.set(file, module);
  const native = createRequire(file), code = ts.transpileModule(require('node:fs').readFileSync(file, 'utf8'), { compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.CommonJS, esModuleInterop: true } }).outputText;
  new Function('require', 'module', 'exports', code)(name => name.startsWith('.') && !path.extname(name) ? load(path.resolve(path.dirname(file), name + '.ts'), cache) : native(name), module, module.exports);
  return module.exports;
}
module.exports = async function fixture() {
  const root = path.dirname(require.resolve('@rhwp/core')), { default: init, HwpDocument } = await import(pathToFileURL(path.join(root, 'rhwp.js')));
  await init({ module_or_path: await fs.readFile(path.join(root, 'rhwp_bg.wasm')) });
  const { editableHwp } = load(path.join(process.env.HWP_FIXTURE_SOURCE_ROOT || path.resolve(__dirname, '..'), 'src/hwp/conversion.ts'));
  const { elementRanges } = load(path.join(__dirname, '../src/hwp/markup-ranges.ts'));
  const doc = HwpDocument.createEmpty(), tail = ' TAIL START ' + 'This editable text must wrap beside the label and regain the full paragraph width below it. '.repeat(5) + ' TAIL END';
  try {
    doc.createBlankDocument();
    doc.pasteHtml(0, 0, 0, '<p>BEFORE</p><table style="width:160px"><tr><td style="width:160px;height:28px;padding:0;border:1px solid black"><p style="font-family:Arial;font-size:12pt;line-height:100%;margin:0">INLINE LABEL</p></td></tr></table><p>AFTER</p>');
    const files = unzipSync(doc.exportHwpx());
    files['Contents/section0.xml'] = strToU8(strFromU8(files['Contents/section0.xml']).replace('</hp:tbl>', '</hp:tbl><hp:t>' + tail + '</hp:t>'));
    const source = new HwpDocument(zipSync(files));
    try {
      const c = JSON.parse(source.getControls()).find(c => c.ctrlId === 'tbl' && c.list === 0);
      source.setTableProperties(0, c.para, c.controlIndex, JSON.stringify({ treatAsChar: true }));
      source.applyCharFormat(0, c.para, 0, source.getParagraphLength(0, c.para), JSON.stringify({ fontSize: 1200 }));
      source.applyParaFormat(0, c.para, JSON.stringify({ lineSpacing: 160, lineSpacingType: 'Percent', indent: -2700, marginLeft: 0 }));
      source.reflowLinesegs();
      // The core's synthetic HWPX writer leaves newly edited host line caches
      // empty. Derive this fixture's explicit cache from the same binary HWP
      // layout; do not let a stale one-line fixture masquerade as multiline.
      const normalized = new HwpDocument(source.exportHwp()), hwpx = unzipSync(normalized.exportHwpx());
      const metrics = JSON.parse(normalized.getParagraphLineMetrics(0, c.para, '[]')).lines;
      const units = value => Math.round(value * 75);
      const cache = '<hp:linesegarray>' + metrics.map(line => `<hp:lineseg textpos="${line.textStart}" vertpos="${units(line.verticalPosition)}" vertsize="${units(line.height)}" textheight="${units(line.textHeight)}" baseline="${units(line.baseline)}" spacing="${units(line.spacing)}" horzpos="0" horzsize="${units(line.width)}" flags="${line.flags}"/>`).join('') + '</hp:linesegarray>';
      let xml = strFromU8(hwpx['Contents/section0.xml']);
      const host = elementRanges(xml, 'hp:p').find(p => p.inner.includes('<hp:tbl'));
      xml = xml.slice(0, host.start) + host.open + host.inner + cache + '</hp:p>' + xml.slice(host.end);
      hwpx['Contents/section0.xml'] = strToU8(xml); normalized.free();
      const result = {};
      for (const format of ['hwp', 'hwpx']) {
        const bytes = format === 'hwp' ? source.exportHwp() : zipSync(hwpx), reopened = new HwpDocument(bytes);
        try { result[format] = { bytes, tail, editable: editableHwp(reopened), lines: JSON.parse(reopened.getLineInfo(0, c.para, 0)), properties: JSON.parse(reopened.getParaPropertiesAt(0, c.para)), char: JSON.parse(reopened.getCharPropertiesAt(0, c.para, 0)) }; }
        finally { reopened.free(); }
      }
      return result;
    } finally { source.free(); }
  } finally { doc.free(); }
};

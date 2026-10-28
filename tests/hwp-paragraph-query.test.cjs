'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url');
let HwpDocument;
test.before(async () => {
  const root = process.env.RHWP_CORE_DIRECTORY || path.dirname(require.resolve('@rhwp/core'));
  const core = await import(pathToFileURL(path.join(root, 'rhwp.js')));
  await core.default({ module_or_path: await fs.readFile(path.join(root, 'rhwp_bg.wasm')) }); HwpDocument = core.HwpDocument;
});
for (const format of ['hwp', 'hwpx']) test(`${format}: nested paragraph queries preserve formatting without changing the document`, () => {
  const original = HwpDocument.createEmpty(); let doc;
  try {
    original.createBlankDocument();
    original.pasteHtml(0, 0, 0, '<table style="width:500px"><tr><td><p style="text-align:center;line-height:25px">OUTER</p><table style="width:300px"><tr><td><p style="text-align:right;line-height:37px;margin-left:12px;text-indent:-5px"><span style="font-size:11pt">NESTED</span></p><table><tr><td><p style="text-align:center;line-height:19px">DEEP</p></td></tr></table></td></tr></table></td></tr></table><p style="text-align:right;line-height:37px;margin-left:12px;text-indent:-5px">REFERENCE</p>');
    doc = new HwpDocument(format === 'hwp' ? original.exportHwp() : original.exportHwpx());
    const control = JSON.parse(doc.getControls()).find(c => c.list === 0 && c.ctrlId === 'tbl');
    const outer = [{ controlIndex: control.controlIndex, cellIndex: 0, cellParaIndex: 0 }];
    const nested = [{ ...outer[0], cellParaIndex: 1 }, { controlIndex: 0, cellIndex: 0, cellParaIndex: 0 }];
    const deep = [{ ...outer[0], cellParaIndex: 1 }, { controlIndex: 0, cellIndex: 0, cellParaIndex: 1 }, { controlIndex: 0, cellIndex: 0, cellParaIndex: 0 }];
    const before = Buffer.from(doc.exportHwpx());
    const bodyMetrics = JSON.parse(doc.getParagraphLineMetrics(0, control.para, '[]'));
    assert.equal(bodyMetrics.tables.length, 1); assert.equal(bodyMetrics.tables[0].inline, true);
    const nestedMetrics = JSON.parse(doc.getParagraphLineMetrics(0, control.para, JSON.stringify(nested)));
    assert(nestedMetrics.lines.length > 0); assert(nestedMetrics.lines.every(line => line.height > 0 && Number.isFinite(line.spacing)));
    assert.throws(() => doc.getParagraphLineMetrics(9999, 0, '[]'), /문단/);

    assert.deepEqual(JSON.parse(doc.getCellParaPropertiesAtByPath(0, control.para, JSON.stringify(outer))), JSON.parse(doc.getCellParaPropertiesAt(0, control.para, control.controlIndex, 0, 0)));
    const actual = JSON.parse(doc.getCellParaPropertiesAtByPath(0, control.para, JSON.stringify(nested)));
    assert.equal(actual.alignment, 'right'); assert.equal(actual.lineSpacingType, 'Fixed'); assert.equal(actual.lineSpacing, 37);
    const ref = Array.from({ length: doc.getParagraphCount(0) }, (_, i) => i).find(i => doc.getTextRange(0, i, 0, doc.getParagraphLength(0, i)) === 'REFERENCE');
    const expected = JSON.parse(doc.getParaPropertiesAt(0, ref));
    for (const key of ['alignment', 'lineSpacing', 'lineSpacingType', 'marginLeft', 'marginRight', 'indent']) assert.equal(actual[key], expected[key], key);
    const last = JSON.parse(doc.getCellParaPropertiesAtByPath(0, control.para, JSON.stringify(deep)));
    assert.equal(last.alignment, 'center'); assert.equal(last.lineSpacing, 19);
    assert.throws(() => doc.getCellParaPropertiesAtByPath(0, control.para, '[]'), /경로/);
    assert.throws(() => doc.getCellParaPropertiesAtByPath(0, control.para, JSON.stringify([{ ...outer[0], cellIndex: 9999 }])), /셀/);
    assert.deepEqual(Buffer.from(doc.exportHwpx()), before);
  } finally { doc?.free(); original.free(); }
});

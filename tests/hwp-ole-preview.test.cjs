const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { createRequire } = require('node:module');
const CFB = require('cfb'), ts = require('typescript');
const file = path.resolve(__dirname, '../src/hwp/ole-preview.ts'), compiled = ts.transpileModule(fs.readFileSync(file, 'utf8'), {
  compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.CommonJS, esModuleInterop: true },
}).outputText;
const loaded = { exports: {} }; new Function('require', 'module', 'exports', compiled)(createRequire(file), loaded, loaded.exports);
const { olePreview } = loaded.exports;
function dib(color, width = 2, height = 2) {
  const b = Buffer.alloc(40 + 16); b.writeUInt32LE(40); b.writeInt32LE(width, 4); b.writeInt32LE(height, 8);
  b.writeUInt16LE(1, 12); b.writeUInt16LE(24, 14);
  for (const at of [40, 43, 48, 51]) b.set(color, at); return b;
}
function presentation(data, kind = 3) {
  const header = Buffer.alloc(40); header.writeUInt32LE(0xffffffff); header.writeUInt32LE(kind, 4); header.writeUInt32LE(4, 8);
  header.writeUInt32LE(1, 12); header.writeUInt32LE(data.length, 36);
  const c = CFB.utils.cfb_new(); CFB.utils.cfb_add(c, '\x02OlePres000', Buffer.concat([header, data]));
  CFB.utils.cfb_add(c, 'OOXMLChartContents', Buffer.from('<c:chartSpace><a:t>CHART</a:t><c:v>65.4</c:v></c:chartSpace>'));
  return new Uint8Array(Buffer.concat([Buffer.alloc(4), CFB.write(c, { type: 'buffer' })]));
}
function record(op, data = Buffer.alloc(0)) { const header = Buffer.alloc(6); header.writeUInt32LE((data.length + 6) / 2); header.writeUInt16LE(op, 4); return Buffer.concat([header, data]); }
function wmf(extra = [], raster = 0x00cc0020) {
  const extent = Buffer.alloc(4); extent.writeInt16LE(2); extent.writeInt16LE(2, 2);
  const map = Buffer.from([8,0]), drawing = Buffer.alloc(20); drawing.writeUInt32LE(raster);
  for (const at of [4,6,12,14]) drawing.writeInt16LE(2, at);
  const body = Buffer.concat([record(0x103,map),record(0x20b,Buffer.alloc(4)),record(0x20c,extent),...extra,record(0xb41,Buffer.concat([drawing,dib([11,22,33])])),record(0)]);
  const header = Buffer.alloc(18); header.writeUInt16LE(1); header.writeUInt16LE(9,2); header.writeUInt16LE(0x300,4); header.writeUInt32LE((body.length+18)/2,6);
  return Buffer.concat([header,body]);
}
test('OLE raster presentation retains the original pixels and accessible chart values, with either compound header form', () => {
  const source = presentation(wmf());
  for (const input of [source,source.subarray(4)]) {
    const result = olePreview(input); assert.ok(result); assert.equal(result.description,'CHART · 65.4');
    assert.equal(Buffer.from(result.bytes).subarray(54,60).toString('hex'),'0b16210b1621');
  }
  assert.ok(olePreview(presentation(dib([0,0,255]),8)));
});
test('OLE rejects malformed or unsupported presentations instead of delivering a partial picture', () => {
  assert.equal(olePreview(presentation(wmf([record(0x214,Buffer.alloc(4))]))),undefined); // lineTo
  assert.equal(olePreview(presentation(wmf([],0))),undefined); // unsupported raster op
  assert.equal(olePreview(presentation(dib([0,0,0],100000,100000),8)),undefined);
  assert.equal(olePreview(presentation(wmf().subarray(0,40))),undefined);
  assert.equal(olePreview(new Uint8Array(32)),undefined);
});

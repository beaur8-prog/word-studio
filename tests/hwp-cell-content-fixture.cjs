const fs = require('node:fs/promises');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const JSZip = require('jszip');
module.exports = async function createCellContentFixture(directory) {
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  const doc = HwpDocument.createEmpty();
  try {
    doc.createBlankDocument();
    doc.pasteHtml(0, 0, 0, '<p>SOURCE</p><table><tr><td style="width:240px;height:60px"><p>FIRST</p></td><td style="width:240px;height:60px"><p>SHAPE</p></td></tr><tr><td style="width:240px;height:90px"><table><tr><td style="width:120px;height:60px"><p>NESTED</p></td><td style="width:120px;height:60px"><p>NESTED_SHAPE</p></td></tr></table></td><td style="width:240px;height:90px"><p>UNCHANGED 한글</p></td></tr></table>');
    doc.insertEquation(0, 0, 0, 'a over b', 1100, 0);
    const shape = JSON.parse(doc.createShapeControl(JSON.stringify({ sectionIdx: 0, paraIdx: 0, charOffset: 0, shapeType: 'rectangle', width: 3000, height: 1500, treatAsChar: true })));
    doc.setShapeProperties(0, shape.paraIdx, shape.controlIdx, JSON.stringify({ borderColor: 255, borderWidth: 75, fillType: 'solid', fillBgColor: 0x00ff00 }));
    const zip = await JSZip.loadAsync(doc.exportHwpx());
    let xml = await zip.file('Contents/section0.xml').async('string');
    const equation = xml.match(/<hp:equation\b[\s\S]*?<\/hp:equation>/)[0];
    const rectangle = xml.match(/<hp:rect\b[\s\S]*?<\/hp:rect>/)[0];
    let id = 100000;
    const eq = script => equation.replace(/\bid="\d+"/, `id="${++id}"`).replace(/<hp:script>[\s\S]*?<\/hp:script>/, `<hp:script>${script}</hp:script>`);
    const rect = () => rectangle.replace(/\bid="\d+"/, `id="${++id}"`);
    xml = xml.replace('<hp:t>FIRST</hp:t>', `<hp:t>BEFORE </hp:t>${eq('x^2')}<hp:t> MIDDLE </hp:t>${eq('y over z')}<hp:t> AFTER</hp:t>`)
      .replace('<hp:t>NESTED</hp:t>', eq('sqrt {q}'))
      .replace('<hp:t>SHAPE</hp:t>', rect())
      .replace('<hp:t>NESTED_SHAPE</hp:t>', rect())
      .replace('<hp:t>UNCHANGED 한글</hp:t>', '<hp:t>UNCHANGED 한글</hp:t><hp:ctrl><hp:header id="100500" applyPageType="BOTH"><hp:subList id="0" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0"><hp:p id="100501" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:t/></hp:run></hp:p></hp:subList></hp:header></hp:ctrl><hp:ctrl><hp:autoNum num="37" numType="PAGE"><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar="" supScript="0"/></hp:autoNum></hp:ctrl>');
    zip.file('Contents/section0.xml', xml);
    const source = new HwpDocument(await zip.generateAsync({ type: 'uint8array' }));
    try {
      await fs.mkdir(directory, { recursive: true }); const files = {};
      for (const format of ['hwp', 'hwpx']) {
        files[format] = path.join(directory, 'cell-content.' + format);
        await fs.writeFile(files[format], format === 'hwp' ? source.exportHwp() : source.exportHwpx());
      }
      return files;
    } finally { source.free(); }
  } finally { doc.free(); }
};

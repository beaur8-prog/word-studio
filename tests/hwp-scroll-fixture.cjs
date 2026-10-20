'use strict';
const fs = require('node:fs/promises');
const path = require('node:path');

async function createHwpScrollFixture(directory, count = 1000) {
  await fs.mkdir(directory, { recursive: true });
  const { default: init, HwpDocument } = await import('@rhwp/core');
  await init({ module_or_path: await fs.readFile(path.join(path.dirname(require.resolve('@rhwp/core')), 'rhwp_bg.wasm')) });
  const doc = HwpDocument.createEmpty();
  try {
    doc.createBlankDocument();
    const html = Array.from({ length: count }, (_, i) => `<p>SCROLL-PAGE-${String(i + 1).padStart(4, '0')} 한글 스크롤 검증</p>`).join('');
    if (!JSON.parse(doc.pasteHtml(0, 0, 0, html)).ok) throw Error('Could not create scroll fixture');
    for (let para = 1; para < doc.getParagraphCount(0); para++) doc.applyParaFormat(0, para, JSON.stringify({ pageBreakBefore: true }));
    if (doc.pageCount() !== count) throw Error('Incorrect scroll fixture page count');
    const file = path.join(directory, `scroll-${count}.hwp`), output = doc.exportHwpWithReport();
    try { await fs.writeFile(file, output.takeBytes()); } finally { output.free(); }
    return file;
  } finally { doc.free(); }
}
module.exports = { createHwpScrollFixture };

'use strict';
const fs = require('node:fs/promises');
const path = require('node:path');
const { randomFillSync, createHash } = require('node:crypto');
const sharp = require('sharp');

// Generated on demand: three valid, incompressible PNGs produce a ~21 MB file
// with >24 MiB of editable HTML. No large binaries or user documents are stored in Git.
async function createLargeHwpFixture(directory, copies = 1) {
  await fs.mkdir(directory, { recursive: true });
  const { default: init, HwpDocument } = await import('@rhwp/core');
  await init({ module_or_path: await fs.readFile(path.join(path.dirname(require.resolve('@rhwp/core')), 'rhwp_bg.wasm')) });
  const width = 1600, height = 1458;
  const raw = randomFillSync(Buffer.alloc(width * height * 3 * 3));
  const images = [];
  for (let i = 0; i < 3; i++) images.push(await sharp(raw.subarray(i * width * height * 3, (i + 1) * width * height * 3), { raw: { width, height, channels: 3 } }).png({ compressionLevel: 0 }).toBuffer());
  const largeImage = path.join(directory, 'single-large-image.png');
  await fs.writeFile(largeImage, await sharp(raw, { raw: { width, height: height * 3, channels: 3 } }).png({ compressionLevel: 0 }).toBuffer());
  const pictures = Array.from({ length: copies }, () => images).flat();
  const html = pictures.map((bytes, i) => `<p>LARGE-IMAGE-${i + 1}</p><p><img src="data:image/png;base64,${bytes.toString('base64')}" width="320" height="292"></p>`).join('');
  const doc = HwpDocument.createEmpty(), files = {};
  try {
    doc.createBlankDocument();
    if (!JSON.parse(doc.pasteHtml(0, 0, 0, html)).ok) throw Error('Could not create large fixture');
    for (let para = 0; para < doc.getParagraphCount(0); para++) {
      if (Number(doc.getTextRange(0, para, 0, 100).match(/LARGE-IMAGE-(\d+)/)?.[1]) > 1) doc.applyParaFormat(0, para, JSON.stringify({ pageBreakBefore: true }));
    }
    for (const format of ['hwp', 'hwpx']) {
      const output = format === 'hwp' ? doc.exportHwpWithReport() : doc.exportHwpxWithReport();
      try { const bytes = output.takeBytes(); const file = path.join(directory, 'large-images.' + format); await fs.writeFile(file, bytes); files[format] = { file, bytes: bytes.byteLength }; }
      finally { output.free(); }
    }
  } finally { doc.free(); }
  const result = { files, largeImage, imageSHA256: pictures.map(bytes => createHash('sha256').update(bytes).digest('hex')), htmlBytes: Buffer.byteLength(html), pages: pictures.length };
  await fs.writeFile(path.join(directory, 'fixture.json'), JSON.stringify(result, null, 2) + '\n');
  return result;
}
module.exports = { createLargeHwpFixture };

import * as CFB from 'cfb';

const LIMIT = 16 * 1024 * 1024, PIXELS = 4_000_000;
const view = (b: Uint8Array) => new DataView(b.buffer, b.byteOffset, b.byteLength);
type Bitmap = { width: number; height: number; rgb: Uint8Array };

/** Read only uncompressed Windows DIB pixels. Never activate the OLE server.
 * MS-OLEDS OLEPresentationStream; MS-WMF BitmapInfoHeader / DIBSTRETCHBLT.
 * Unsupported vector records are rejected rather than partially rendered. */
function dib(bytes: Uint8Array): Bitmap {
  const d = view(bytes);
  if (bytes.length < 40 || d.getUint32(0, true) !== 40 || d.getUint16(12, true) !== 1 || d.getUint32(16, true) !== 0) throw Error('Unsupported OLE bitmap');
  const width = d.getInt32(4, true), signedHeight = d.getInt32(8, true), height = Math.abs(signedHeight), depth = d.getUint16(14, true);
  if (width <= 0 || height <= 0 || width * height > PIXELS || ![1, 4, 8, 24, 32].includes(depth)) throw Error('Invalid OLE bitmap dimensions');
  const colors = depth <= 8 ? d.getUint32(32, true) || 2 ** depth : 0;
  if (colors > 2 ** Math.min(depth, 8)) throw Error('Invalid OLE bitmap palette');
  const start = 40 + colors * 4, stride = Math.ceil(width * depth / 32) * 4;
  if (start + stride * height > bytes.length) throw Error('Truncated OLE bitmap');
  const rgb = new Uint8Array(width * height * 3);
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
    const row = start + (signedHeight > 0 ? height - y - 1 : y) * stride;
    let from = row + x * depth / 8;
    if (depth <= 8) {
      const index = depth === 8 ? bytes[from] : depth === 4 ? (bytes[row + (x >> 1)] >> (x % 2 ? 0 : 4)) & 15 : (bytes[row + (x >> 3)] >> (7 - x % 8)) & 1;
      if (index >= colors) throw Error('Invalid OLE bitmap palette index');
      from = 40 + index * 4;
    }
    const to = (y * width + x) * 3;
    rgb[to] = bytes[from + 2]; rgb[to + 1] = bytes[from + 1]; rgb[to + 2] = bytes[from];
  }
  return { width, height, rgb };
}

function bitmapWmf(bytes: Uint8Array): Bitmap {
  const d = view(bytes);
  if (bytes.length < 18 || ![1, 2].includes(d.getUint16(0, true)) || d.getUint16(2, true) !== 9 || d.getUint16(4, true) !== 0x300 || d.getUint32(6, true) * 2 !== bytes.length) throw Error('Invalid OLE WMF');
  let output: Bitmap | undefined, width = 0, height = 0, ended = false, painted = 0;
  for (let at = 18, count = 0; at < bytes.length;) {
    if (++count > 10000 || at + 6 > bytes.length) throw Error('Invalid OLE WMF records');
    const size = d.getUint32(at, true) * 2, op = d.getUint16(at + 4, true);
    if (size < 6 || at + size > bytes.length) throw Error('Truncated OLE WMF record');
    const data = bytes.subarray(at + 6, at + size), r = view(data);
    if (op === 0) { if (size !== 6 || at + size !== bytes.length) throw Error('Invalid OLE WMF end'); ended = true; }
    else if (op === 0x103) { if (data.length !== 2 || r.getUint16(0, true) !== 8) throw Error('Unsupported OLE mapping'); }
    else if (op === 0x20b) { if (data.length !== 4 || r.getInt32(0, true) !== 0) throw Error('Unsupported OLE origin'); }
    else if (op === 0x20c) {
      if (data.length !== 4 || output) throw Error('Unsupported OLE extent');
      height = r.getInt16(0, true); width = r.getInt16(2, true);
      if (width <= 0 || height <= 0 || width * height > PIXELS) throw Error('Invalid OLE extent');
    } else if (op === 0xb41) {
      if (!width || data.length < 60 || ++painted > 64) throw Error('Invalid OLE bitmap record');
      const raster = r.getUint32(0, true);
      // Full-extent raster presentations only. Scaling/cropping and arbitrary
      // graphics state require a complete metafile renderer, not guesswork.
      if (![0x00cc0020, 0x00ee0086, 0x008800c6].includes(raster) ||
        [r.getInt16(4, true), r.getInt16(6, true), r.getInt16(8, true), r.getInt16(10, true), r.getInt16(12, true), r.getInt16(14, true), r.getInt16(16, true), r.getInt16(18, true)].some((n, i) => n !== [height, width, 0, 0, height, width, 0, 0][i])) throw Error('Unsupported OLE raster operation');
      const image = dib(data.subarray(20));
      if (image.width !== width || image.height !== height) throw Error('OLE bitmap extent mismatch');
      output ??= { width, height, rgb: new Uint8Array(width * height * 3).fill(255) };
      for (let i = 0; i < output.rgb.length; i++) output.rgb[i] = raster === 0x00cc0020 ? image.rgb[i] : raster === 0x00ee0086 ? output.rgb[i] | image.rgb[i] : output.rgb[i] & image.rgb[i];
    } else if (op === 0x626) {
      // Comment escape: an alternate EMF presentation or producer comment.
      // Ignore its inert bytes; raster fallback records remain authoritative.
      if (data.length < 4 || r.getUint16(0, true) !== 15 || r.getUint16(2, true) > data.length - 4) throw Error('Unsupported OLE escape');
    } else if (![0x1e, 0x127, 0x107].includes(op)) throw Error('Unsupported OLE vector record');
    at += size;
  }
  if (!ended || !output) throw Error('Missing OLE raster preview');
  return output;
}

function bmp(image: Bitmap): Uint8Array {
  const { width, height, rgb } = image, stride = Math.ceil(width * 3 / 4) * 4;
  const bytes = new Uint8Array(54 + stride * height), d = view(bytes);
  bytes.set([0x42, 0x4d]); d.setUint32(2, bytes.length, true); d.setUint32(10, 54, true); d.setUint32(14, 40, true);
  d.setInt32(18, width, true); d.setInt32(22, -height, true); d.setUint16(26, 1, true); d.setUint16(28, 24, true);
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
    const from = (y * width + x) * 3, to = 54 + y * stride + x * 3;
    bytes[to] = rgb[from + 2]; bytes[to + 1] = rgb[from + 1]; bytes[to + 2] = rgb[from];
  }
  return bytes;
}

export function olePreview(bytes: Uint8Array): { bytes: Uint8Array; description: string } | undefined {
  try {
    if (bytes.length > LIMIT) return;
    const magic = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
    const offset = [0, 4].find(offset => magic.every((b, i) => bytes[offset + i] === b));
    if (offset === undefined) return;
    const compound = CFB.read(bytes.subarray(offset), { type: 'array' });
    const entry = compound.FileIndex.find(e => /^\x02OlePres\d{3}$/.test(e.name));
    if (!entry?.content || entry.content.length > LIMIT) return;
    const data = new Uint8Array(entry.content), d = view(data);
    if (data.length < 40 || d.getUint32(0, true) !== 0xffffffff || d.getUint32(8, true) !== 4) return;
    const kind = d.getUint32(4, true), length = d.getUint32(36, true);
    if (length > LIMIT || 40 + length > data.length) return;
    const payload = data.subarray(40, 40 + length);
    const image = kind === 3 ? bitmapWmf(payload) : kind === 8 ? dib(payload) : undefined;
    if (!image) return;
    // Preserve chart labels/values as accessible text as well as pixels. Read
    // cached values only; external workbook links are never followed.
    const chart = compound.FileIndex.find(e => e.name === 'OOXMLChartContents')?.content;
    const text = chart && chart.length <= 1024 * 1024 ? new TextDecoder().decode(new Uint8Array(chart)) : '';
    const description = [...text.matchAll(/<(?:a:t|c:v)>([^<]{0,1000})<\/(?:a:t|c:v)>/g)].map(m => m[1]).join(' · ').slice(0,10000);
    return { bytes: bmp(image), description };
  } catch { return; }
}

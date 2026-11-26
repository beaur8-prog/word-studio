import { strFromU8, strToU8 } from 'fflate';
import documentLimits from '../../electron/document-limits.json';

type RasterFormat = 'bmp' | 'png';
export type StagedRaster = {
  format: RasterFormat; encoded: string; bytes: number; width: number; height: number; marker: Uint8Array;
};
type Archive = Record<string, Uint8Array>;
const metadata = 'Contents/content.hpf';
const sameBytes = (a: Uint8Array | undefined, b: Uint8Array) =>
  !!a && a.length === b.length && a.every((byte, i) => byte === b[i]);

function bitmapMarker(index: number) {
  const bytes = new Uint8Array(58), view = new DataView(bytes.buffer);
  view.setUint16(0, 0x4d42, true);
  view.setUint32(2, bytes.length, true);
  view.setUint32(10, 54, true);
  view.setUint32(14, 40, true);
  view.setInt32(18, 1, true);
  view.setInt32(22, 1, true);
  view.setUint16(26, 1, true);
  view.setUint16(28, 24, true);
  view.setUint32(34, 4, true);
  view.setUint32(54, index, true);
  return bytes;
}

function bmpDimensions(header: DataView, bytes: number) {
  const width = header.getInt32(18, true), height = header.getInt32(22, true);
  if (header.getUint16(0, true) !== 0x4d42 || header.getUint32(2, true) !== bytes || header.getUint32(10, true) !== 54 ||
      header.getUint32(14, true) !== 40 || width <= 0 || height <= 0 ||
      header.getUint16(26, true) !== 1 || header.getUint16(28, true) !== 24 || header.getUint32(30, true) !== 0 ||
      Math.ceil(width * 3 / 4) * 4 * height !== bytes - 54) return;
  return { width, height };
}

function pngDimensions(header: DataView) {
  if (header.getUint32(0) !== 0x89504e47 || header.getUint32(4) !== 0x0d0a1a0a ||
      header.getUint32(8) !== 13 || header.getUint32(12) !== 0x49484452) return;
  const width = header.getUint32(16), height = header.getUint32(20);
  if (!width || !height || header.getUint8(26) || header.getUint8(27) || header.getUint8(28) > 1) return;
  return { width, height };
}

/** Preserve the native picture control without including the full raster in
 * one paste block. Explicit display dimensions keep the staging image inert
 * with respect to the space occupied by the picture. */
export function stageLargeRaster(tag: string, limit: number, index: number): { tag: string; raster: StagedRaster } | undefined {
  if (!limit || tag.length <= limit || index < 1 || index > 0xffffff) return;
  const source = /\ssrc="data:image\/(bmp|png);base64,([^"]+)"/i.exec(tag);
  if (!source) return;
  const format = source[1].toLowerCase() as RasterFormat, encoded = source[2];
  if (encoded.length <= limit || encoded.length % 4) return;
  const bytes = (encoded.length / 4) * 3 - (encoded.endsWith('==') ? 2 : encoded.endsWith('=') ? 1 : 0);
  if (bytes > documentLimits.embeddedImageBytes) return;
  if (!['width', 'height'].every(name => {
    const size = Number(tag.match(new RegExp('\\s' + name + '="([0-9]+(?:\\.[0-9]+)?)"'))?.[1]);
    return Number.isFinite(size) && size > 0;
  })) return;
  const content = Uint8Array.from(atob(encoded.slice(0, 72)), char => char.charCodeAt(0));
  if (content.length !== 54) return;
  const header = new DataView(content.buffer);
  const dimensions = format === 'bmp' ? bmpDimensions(header, bytes) : pngDimensions(header);
  if (!dimensions || Math.max(dimensions.width, dimensions.height) * 75 > 0x7fffffff) return;
  const marker = bitmapMarker(index);
  const cheapSource = ` src="data:image/bmp;base64,${btoa(String.fromCharCode(...marker))}"`;
  return { tag: tag.slice(0, source.index) + cheapSource + tag.slice(source.index + source[0].length),
    raster: { format, encoded, bytes, ...dimensions, marker } };
}

/** Restore the bytes at the picture's own resolved asset, with the same
 * full-image extent the native bitmap parser computes for an image pasted directly. */
export function restoreLargeRaster(picture: string, files: Archive, raster: StagedRaster) {
  const failure = () => Error('한글 파일에 큰 그림의 원본 데이터를 저장하지 못했습니다.');
  const ref = /<hc:img\b[^>]*\bbinaryItemIDRef="(image\d+)"/.exec(picture)?.[1];
  const content = files[metadata];
  if (!ref || !content) throw failure();
  const manifest = strFromU8(content);
  const items = [...manifest.matchAll(/<opf:item\b[^>]*\/>/g)].filter(([item]) => item.includes(`id="${ref}"`));
  if (items.length !== 1) throw failure();
  const [original] = items[0], href = /\bhref="([^"]+)"/.exec(original)?.[1];
  const target = `BinData/${ref}.${raster.format}`;
  if (!href || ![`BinData/${ref}.bmp`, `BinData/${ref}.octet-stream`].includes(href) ||
      !sameBytes(files[href], raster.marker) || (href !== target && files[target]) ||
      !picture.includes('<hp:imgClip left="0" right="75" top="0" bottom="75"/>') ||
      !picture.includes('<hp:imgDim dimwidth="75" dimheight="75"/>')) throw failure();
  const payload = atob(raster.encoded);
  if (payload.length !== raster.bytes) throw failure();
  const bytes = new Uint8Array(raster.bytes);
  for (let i = 0; i < bytes.length; i++) bytes[i] = payload.charCodeAt(i);
  const patched = original.replace(`href="${href}"`, `href="${target}"`).replace(/\bmedia-type="[^"]*"/, `media-type="image/${raster.format}"`);
  if (href !== target) delete files[href];
  files[target] = bytes;
  files[metadata] = strToU8(manifest.replace(original, () => patched));
  const w = raster.width * 75, h = raster.height * 75;
  return picture.replace('<hp:imgClip left="0" right="75" top="0" bottom="75"/>',
    `<hp:imgClip left="0" right="${w}" top="0" bottom="${h}"/>`)
    .replace('<hp:imgDim dimwidth="75" dimheight="75"/>', `<hp:imgDim dimwidth="${w}" dimheight="${h}"/>`);
}

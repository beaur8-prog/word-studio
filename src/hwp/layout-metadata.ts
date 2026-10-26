import * as CFB from 'cfb';
import JSZip from 'jszip';
import { Inflate } from 'fflate';

export type HwpParagraphLayout = { pageBreak: boolean; columnBreak: boolean; columns?: { count: number; space: number } };
export type HwpLayoutMetadata = { paragraphs: Record<string, HwpParagraphLayout>; warnings: string[] };
const LIMIT = 32 * 1024 * 1024;
const columns = (count: number, gap: number) => count >= 1 && count <= 12 && gap >= 0 && gap <= 75000
  ? { count, space: Math.round(gap / 75 * 100) / 100 } : undefined;

/** Only top-level paragraph headers and column definitions are needed. Nested
 * table/header paragraphs must not shift the main document's paragraph index. */
export function parseHwpSection(bytes: Uint8Array, section: number): Record<string, HwpParagraphLayout> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength), result: Record<string, HwpParagraphLayout> = {};
  let paragraph = -1, current: HwpParagraphLayout | undefined;
  for (let offset = 0; offset < bytes.length;) {
    if (offset + 4 > bytes.length) throw new Error('Truncated HWP record header');
    const header = view.getUint32(offset, true), tag = header & 1023, level = (header >>> 10) & 1023;
    let length = header >>> 20; offset += 4;
    if (length === 4095) {
      if (offset + 4 > bytes.length) throw new Error('Truncated extended HWP record');
      length = view.getUint32(offset, true); offset += 4;
    }
    if (offset + length > bytes.length) throw new Error('Truncated HWP record');
    if (tag === 66 && level === 0) {
      if (length < 12) throw new Error('Invalid paragraph header');
      const flags = bytes[offset + 11];
      current = { pageBreak: !!(flags & 4), columnBreak: !!(flags & 8) };
      result[`${section}:${++paragraph}`] = current;
    } else if (tag === 71 && level === 1 && length >= 8 && current && view.getUint32(offset, true) === 0x636f6c64) {
      current.columns = columns((view.getUint16(offset + 4, true) >>> 2) & 255, view.getUint16(offset + 6, true));
    }
    offset += length;
  }
  return result;
}

export function parseHwpxSection(xml: string, section: number): Record<string, HwpParagraphLayout> {
  const result: Record<string, HwpParagraphLayout> = {};
  let depth = 0, paragraph = -1, current: HwpParagraphLayout | undefined;
  const attr = (tag: string, name: string) => tag.match(new RegExp(`\\b${name}=["']([^"']*)["']`))?.[1];
  for (const match of xml.matchAll(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\/?[\w.-]+:(?:p|colPr)\b[^>]*>/g)) {
    const tag = match[0];
    if (tag.startsWith('<!')) continue;
    if (/^<\/[\w.-]+:p\b/.test(tag)) { depth--; continue; }
    if (/^<[\w.-]+:p\b/.test(tag)) {
      if (depth === 0) {
        current = { pageBreak: attr(tag, 'pageBreak') === '1', columnBreak: attr(tag, 'columnBreak') === '1' };
        result[`${section}:${++paragraph}`] = current;
      }
      if (!tag.endsWith('/>')) depth++;
    } else if (depth === 1 && current) current.columns = columns(Number(attr(tag, 'colCount')), Number(attr(tag, 'sameGap') || 0));
  }
  if (depth !== 0) throw new Error('Invalid HWPX paragraph nesting');
  return result;
}

async function inflateBounded(bytes: Uint8Array): Promise<Uint8Array> {
  const chunks: Uint8Array[] = []; let size = 0;
  // HWP streams may contain padding after the final DEFLATE block. The web
  // DecompressionStream rejects that padding; fflate stops at the real end.
  const inflater = new Inflate(chunk => {
    if ((size += chunk.length) > LIMIT) throw new Error('HWP layout metadata exceeds budget');
    chunks.push(chunk);
  });
  for (let offset = 0; offset < bytes.length; offset += 4096)
    inflater.push(bytes.subarray(offset, offset + 4096), offset + 4096 >= bytes.length);
  const output = new Uint8Array(size); let offset = 0;
  for (const chunk of chunks) { output.set(chunk, offset); offset += chunk.length; }
  return output;
}

export async function readHwpLayoutMetadata(bytes: Uint8Array, hwpx = false): Promise<HwpLayoutMetadata> {
  const output: HwpLayoutMetadata = { paragraphs: {}, warnings: [] };
  try {
    if (hwpx) {
      const zip = await JSZip.loadAsync(bytes);
      for (const file of Object.values(zip.files).filter(f => /^Contents\/section\d+\.xml$/.test(f.name))) {
        const size = (file as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? LIMIT + 1;
        if (size > LIMIT) throw new Error('HWPX section metadata exceeds budget');
        const section = Number(file.name.match(/section(\d+)\.xml$/)![1]);
        Object.assign(output.paragraphs, parseHwpxSection(await file.async('string'), section));
      }
    } else {
      const cfb = CFB.read(bytes, { type: 'array' });
      const header = cfb.FileIndex[cfb.FullPaths.findIndex(p => p.endsWith('/FileHeader'))]?.content;
      if (!header || header.length < 40) throw new Error('Missing HWP file header');
      const compressed = !!(header[36] & 1);
      for (let i = 0; i < cfb.FullPaths.length; i++) {
        const match = cfb.FullPaths[i].match(/\/BodyText\/Section(\d+)$/); if (!match) continue;
        let content: Uint8Array = new Uint8Array(cfb.FileIndex[i].content);
        if (content.length > LIMIT) throw new Error('HWP section metadata exceeds budget');
        if (compressed) content = await inflateBounded(content);
        Object.assign(output.paragraphs, parseHwpSection(content, Number(match[1])));
      }
    }
  } catch {
    output.paragraphs = {};
    output.warnings.push('원본의 일부 쪽·단 나누기 정보를 읽지 못했습니다. 편집본의 쪽 배치를 원본과 비교하세요.');
  }
  return output;
}

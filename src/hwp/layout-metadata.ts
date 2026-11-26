import * as CFB from 'cfb';
import JSZip from 'jszip';
import { Inflate } from 'fflate';
import documentLimits from '../../electron/document-limits.json';
import { hwpCellRasterSource, isHwpCellRasterBrightness, type HwpCellRasterPicture } from '../hwp-cell-raster-fill';

export type HwpParagraphLayout = { pageBreak: boolean; columnBreak: boolean; hideFirstPageNumber?: boolean; columns?: { count: number; space: number; balance?: boolean }; lineSpacings?: number[] };
export type HwpLayoutMetadata = { paragraphs: Record<string, HwpParagraphLayout>; warnings: string[];
  fontSpace?: boolean[]; latinFaces?: (string | undefined)[];
  hwpxCellFontSpace?: boolean[]; hwpxCellRasterFills?: Record<number, HwpCellRasterPicture>; conversionBlock?: string };
const LIMIT = 32 * 1024 * 1024;
const columns = (count: number, gap: number) => count >= 1 && count <= 12 && gap >= 0 && gap <= 75000
  ? { count, space: Math.round(gap / 75 * 100) / 100 } : undefined;

/** Read top-level paragraph headers, columns and stored line leading. Nested
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
      const flags = view.getUint16(offset + 4, true);
      current.columns = columns((flags >>> 2) & 255, view.getUint16(offset + 6, true));
      if (current.columns && (flags & 3) === 1) current.columns.balance = true;
    } else if (tag === 71 && level === 1 && length >= 8 && current && view.getUint32(offset, true) === 0x73656364) {
      // The native HWPX serializer does not retain all section visibility bits.
      if (view.getUint32(offset + 4, true) & 32) current.hideFirstPageNumber = true;
    } else if (tag === 69 && level === 1 && current) {
      if (length % 36) throw new Error('Invalid HWP line segments');
      current.lineSpacings = Array.from({ length: length / 36 }, (_, i) => view.getInt32(offset + i * 36 + 20, true) / 75);
    }
    offset += length;
  }
  return result;
}

export function parseHwpxSection(xml: string, section: number): Record<string, HwpParagraphLayout> {
  const result: Record<string, HwpParagraphLayout> = {};
  let depth = 0, paragraph = -1, current: HwpParagraphLayout | undefined;
  const attr = (tag: string, name: string) => tag.match(new RegExp(`\\b${name}=["']([^"']*)["']`))?.[1];
  for (const match of xml.matchAll(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\/?[\w.-]+:(?:p|colPr|lineseg|visibility)\b[^>]*>/g)) {
    const tag = match[0];
    if (tag.startsWith('<!')) continue;
    if (/^<\/[\w.-]+:p\b/.test(tag)) { depth--; continue; }
    if (/^<[\w.-]+:p\b/.test(tag)) {
      if (depth === 0) {
        current = { pageBreak: attr(tag, 'pageBreak') === '1', columnBreak: attr(tag, 'columnBreak') === '1' };
        result[`${section}:${++paragraph}`] = current;
      }
      if (!tag.endsWith('/>')) depth++;
    } else if (depth === 1 && current && /:lineseg\b/.test(tag) && !tag.startsWith('</')) {
      const spacing = Number(attr(tag, 'spacing'));
      if (!Number.isFinite(spacing)) throw new Error('Invalid HWPX line spacing');
      (current.lineSpacings ??= []).push(spacing / 75);
    } else if (depth === 1 && current && /:colPr\b/.test(tag)) {
      current.columns = columns(Number(attr(tag, 'colCount')), Number(attr(tag, 'sameGap') || 0));
      if (current.columns && /balanced/i.test(attr(tag, 'type') || '')) current.columns.balance = true;
    } else if (depth === 1 && current && /:visibility\b/.test(tag) && ['1', 'true'].includes(attr(tag, 'hideFirstPageNum') ?? '')) {
      current.hideFirstPageNumber = true;
    }
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

async function hwpxLayoutParts(zip: JSZip) {
  // Match the native reader's OPF order and path resolution. ZIP entry order
  // and numeric basenames need not be the document's section order.
  const manifest = zip.file('Contents/content.hpf');
  if (!manifest || ((manifest as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? LIMIT + 1) > LIMIT)
    throw new Error('Missing or oversized HWPX manifest');
  const decode = (value: string) => value.replace(/&(?:amp|lt|gt|quot|apos|#\d+|#x[\da-f]+);/gi, entity => {
    const named: Record<string, string> = { '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&apos;': "'" };
    if (named[entity]) return named[entity];
    const point = Number.parseInt(entity.slice(entity[2] === 'x' ? 3 : 2, -1), entity[2] === 'x' ? 16 : 10);
    if (!Number.isFinite(point) || point < 1 || point > 0x10ffff) throw new Error('Invalid XML entity');
    return String.fromCodePoint(point);
  });
  const attr = (tag: string, name: string) => decode(tag.match(new RegExp(`(?:^|\\s)${name}\\s*=\\s*(["'])(.*?)\\1`))?.[2] ?? '');
  const items: { id: string; href: string; type: string }[] = [], spine: string[] = [];
  const xml = (await manifest.async('string')).replace(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>/g, '');
  for (const match of xml.matchAll(/<(?:[\w.-]+:)?(item|itemref)\b[^>]*>/g)) {
    if (match[1] === 'itemref') spine.push(attr(match[0], 'idref'));
    else items.push({ id: attr(match[0], 'id'), href: attr(match[0], 'href'), type: attr(match[0], 'media-type') });
  }
  const section = (item: typeof items[number] | undefined): item is typeof items[number] => !!item && item.type === 'application/xml' && item.href.includes('section');
  let ordered = spine.map(id => items.find(item => item.id === id)).filter(section);
  if (!ordered.length) ordered = items.filter(section).sort((a, b) => a.href < b.href ? -1 : a.href > b.href ? 1 : 0);
  if (!ordered.length) throw new Error('HWPX has no section manifest');
  const normalize = (path: string) => {
    const parts: string[] = [];
    for (const part of path.split('/')) {
      if (!part || part === '.') continue;
      if (part === '..') { if (!parts.pop()) throw new Error('Invalid package path'); }
      else parts.push(part);
    }
    return parts.join('/');
  };
  const part = (href: string) => {
    if (!href || href.startsWith('/') || /[:\\\x00]/.test(href)) throw new Error('Invalid package path');
    const relative = normalize('Contents/' + href);
    return zip.file(relative) ?? zip.file(normalize(href));
  };
  const sections = ordered.map(({ href }) => {
    const entry = part(href);
    if (!entry) throw new Error('Missing HWPX section');
    return entry;
  });
  const headers = items.filter(item => item.id === 'header' && item.type === 'application/xml');
  let header: JSZip.JSZipObject | null = null;
  if (headers.length === 1) {
    try { header = part(headers[0].href); }
    catch { /* Missing optional character metadata leaves section flags intact. */ }
  }
  return { sections, header, items, part, attr };
}

/** Preserve package-local TOTAL picture fills with safe native brightness and no other effect. */
async function hwpxRasterFills(xml: string, layout: Awaited<ReturnType<typeof hwpxLayoutParts>>) {
  const source = xml.replace(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>/g, '');
  const { items, part, attr } = layout;
  const result: Record<number, HwpCellRasterPicture> = {};
  const assets = new Map<string, string>();
  const usedIds = new Set<number>();
  let total = 0;
  for (const [, tag, body] of source.matchAll(/(<[\w.-]+:borderFill\b[^>]*>)([\s\S]*?)<\/[\w.-]+:borderFill\s*>/g)) {
    const idValue = attr(tag, 'id'), id = Number(idValue);
    if (!/^\d{1,8}$/.test(idValue)) continue;
    if (usedIds.has(id)) { delete result[id]; continue; }
    usedIds.add(id);
    if (/<[\w.-]+:(?:winBrush|gradation)\b/.test(body)) continue;
    const brushes = [...body.matchAll(/(<[\w.-]+:imgBrush\b[^>]*>)([\s\S]*?)<\/[\w.-]+:imgBrush\s*>/g)];
    if (brushes.length !== 1 || attr(brushes[0][1], 'mode') !== 'TOTAL') continue;
    const images = [...brushes[0][2].matchAll(/<[\w.-]+:img\b[^>]*>/g)];
    if (images.length !== 1) continue;
    const image = images[0][0], reference = attr(image, 'binaryItemIDRef');
    const authoredBrightness = attr(image, 'bright'), brightness = Number(authoredBrightness);
    if (attr(image, 'effect') !== 'REAL_PIC' || ['contrast', 'alpha'].some(name => attr(image, name) !== '0') ||
        !isHwpCellRasterBrightness(brightness) || String(brightness) !== authoredBrightness) continue;
    const cached = assets.get(reference);
    if (cached) { result[id] = { source: cached, brightness }; continue; }
    const matches = items.filter(item => item.id === reference);
    if (matches.length !== 1 || !/^image\/(?:jpe?g|png)$/i.test(matches[0].type)) continue;
    let entry;
    try { entry = part(matches[0].href); }
    catch { continue; }
    const size = (entry as unknown as { _data?: { uncompressedSize?: number } } | null)?._data?.uncompressedSize ?? LIMIT + 1;
    if (!entry || size < 16 || size > documentLimits.embeddedImageBytes || total + size > LIMIT) continue;
    const type = /png$/i.test(matches[0].type) ? 'png' : 'jpeg';
    const encoded = await entry.async('base64').catch(() => '');
    const url = `data:image/${type};base64,${encoded}`;
    if (!hwpCellRasterSource(url)) continue;
    assets.set(reference, url);
    result[id] = { source: url, brightness };
    total += size;
  }
  return Object.keys(result).length ? result : undefined;
}

/** The Reader indexes character shapes in definition order. Use the XML
 * flags only if the declared IDs and that order agree for the whole set. */
function readHwpxFontSpace(xml: string): boolean[] | undefined {
  const flags: boolean[] = [];
  for (const [tag] of xml.matchAll(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<(?:[\w.-]+:)?charPr\b[^>]*>/g)) {
    if (tag.startsWith('<!')) continue;
    const id = /\sid\s*=\s*(["'])(\d+)\1/.exec(tag)?.[2];
    const space = /\suseFontSpace\s*=\s*(["'])(0|1|false|true)\1/.exec(tag)?.[2];
    if (id !== String(flags.length) || space === undefined) return;
    flags.push(space === '1' || space === 'true');
  }
  return flags.length ? flags : undefined;
}

function readHwpxLatinFaces(xml: string, attr: (tag: string, name: string) => string) {
  const content = xml.replace(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>/g, '');
  const groups = [...content.matchAll(/(<(?:[\w.-]+:)?fontface\b[^>]*>)([\s\S]*?)<\/(?:[\w.-]+:)?fontface\s*>/g)]
    .filter(group => attr(group[1], 'lang') === 'LATIN');
  if (groups.length !== 1) return;
  const names: string[] = [];
  for (const [tag] of groups[0][2].matchAll(/<(?:[\w.-]+:)?font\b[^>]*>/g)) {
    if (attr(tag, 'id') !== String(names.length)) return;
    names.push(attr(tag, 'face'));
  }
  const result: (string | undefined)[] = [];
  for (const [, tag, body] of content.matchAll(/(<(?:[\w.-]+:)?charPr\b[^>]*>)([\s\S]*?)<\/(?:[\w.-]+:)?charPr\s*>/g)) {
    if (attr(tag, 'id') !== String(result.length)) return;
    const reference = /<(?:[\w.-]+:)?fontRef\b[^>]*>/.exec(body)?.[0];
    const latin = reference && attr(reference, 'latin');
    result.push(latin && /^\d+$/.test(latin) ? names[Number(latin)] : undefined);
  }
  return result.length && names.length ? result : undefined;
}

/** DocInfo records contain the authored face name before the renderer
 * resolves an alias. Retain the Latin slot in character-shape order. */
function readHwpFontMetadata(bytes: Uint8Array) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength), result: boolean[] = [];
  const fonts: (string | undefined)[] = [], latinIds: number[] = [];
  const decoder = new TextDecoder('utf-16le');
  let groups: number[] | undefined;
  for (let offset = 0; offset < bytes.length;) {
    const header = view.getUint32(offset, true);
    offset += 4;
    let length = header >>> 20;
    if (length === 4095) { length = view.getUint32(offset, true); offset += 4; }
    if (offset + length > bytes.length) throw new Error('Truncated HWP character metadata');
    const tag = header & 1023;
    if (tag === 17 && length >= 32)
      groups = Array.from({ length: 7 }, (_, i) => view.getUint32(offset + 4 + 4 * i, true));
    else if (tag === 19) {
      const count = length >= 3 ? view.getUint16(offset + 1, true) : 0;
      fonts.push(count > 0 && count <= 500 && 3 + count * 2 <= length
        ? decoder.decode(bytes.subarray(offset + 3, offset + 3 + count * 2)) : undefined);
    } else if (tag === 21) {
      if (length < 50) throw new Error('Invalid HWP character shape');
      result.push(!!(view.getUint32(offset + 46, true) & (1 << 25)));
      latinIds.push(view.getUint16(offset + 2, true));
    }
    offset += length;
  }
  const latinFaces = groups && groups.reduce((sum, count) => sum + count, 0) === fonts.length
    ? latinIds.map(id => id < groups[1] ? fonts[groups[0] + id] : undefined) : undefined;
  return { fontSpace: result, latinFaces };
}

export async function readHwpLayoutMetadata(bytes: Uint8Array, hwpx = false): Promise<HwpLayoutMetadata> {
  const output: HwpLayoutMetadata = { paragraphs: {}, warnings: [] };
  try {
    if (hwpx) {
      const zip = await JSZip.loadAsync(bytes);
      const layout = await hwpxLayoutParts(zip);
      const { sections, header } = layout;
      for (const [section, file] of sections.entries()) {
        const size = (file as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? LIMIT + 1;
        if (size > LIMIT) throw new Error('HWPX section metadata exceeds budget');
        Object.assign(output.paragraphs, parseHwpxSection(await file.async('string'), section));
      }
      if (header && ((header as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? LIMIT + 1) <= LIMIT) {
        const xml = await header.async('string').catch(() => '');
        try { output.hwpxCellFontSpace = readHwpxFontSpace(xml); }
        catch { /* Section flags remain available when optional fonts cannot be read. */ }
        try { output.latinFaces = readHwpxLatinFaces(xml, layout.attr); }
        catch { /* A malformed optional face list keeps the resolved run fonts. */ }
        try { output.hwpxCellRasterFills = await hwpxRasterFills(xml, layout); }
        catch { /* A malformed optional brush must not hide body content. */ }
      }
    } else {
      const cfb = CFB.read(bytes, { type: 'array' });
      const headerPath = (cfb.FullPaths[0] + 'FileHeader').toLowerCase();
      const header = cfb.FileIndex[cfb.FullPaths.findIndex(p => p.toLowerCase() === headerPath)]?.content;
      if (!header || header.length < 40) throw new Error('Missing HWP file header');
      // Tracking ViewText can contain deleted and current text together. The
      // native selection/export APIs omit those revision ranges, so accepting
      // either stream here would silently change numbers (for example 7 + 4).
      // Keep both original streams intact until revision semantics are supported.
      if (header[37] & 0x40) output.conversionBlock = '이 문서에는 변경 내용 추적 정보가 있습니다. 삭제 전·후의 숫자나 문장이 함께 표시될 수 있어 Word 편집본을 만들 수 없습니다. 한글에서 변경 내용을 모두 적용하고 추적을 끈 사본을 저장한 뒤 다시 여세요.';
      const compressed = !!(header[36] & 1);
      const info = cfb.FileIndex[cfb.FullPaths.findIndex(p => p.toLowerCase() === (cfb.FullPaths[0] + 'DocInfo').toLowerCase())]?.content;
      if (info?.length && info.length <= LIMIT) {
        try {
          const bytes = new Uint8Array(info);
          Object.assign(output, readHwpFontMetadata(compressed ? await inflateBounded(bytes) : bytes));
        } catch { /* Older or incomplete character metadata does not affect paragraph flags. */ }
      }
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

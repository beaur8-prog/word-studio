import type { HwpDocument } from '@rhwp/core';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import documentLimits from '../../electron/document-limits.json';
import { imageFloating, type ImageFloating } from '../image-floating';
import { entities } from './cell-content';
import { restoreLargeRaster, stageLargeRaster, type StagedRaster } from './native-bitmap-paste';
import { nativeStandalonePictureAttribute } from './export-picture-line';

const escape = (value: string) => value.replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c]!));
const units = (emu = 0) => Math.round(emu / 127);
function attributes(tag: string, values: Record<string, string | number>) {
  for (const [key, value] of Object.entries(values)) {
    const pattern = new RegExp(`\\b${key}="[^"]*"`), replacement = `${key}="${escape(String(value))}"`;
    tag = pattern.test(tag) ? tag.replace(pattern, replacement) : tag.replace(/\s*\/?>$/, end => ' ' + replacement + end);
  }
  return tag;
}
function position(g: ImageFloating) {
  const h = g.horizontalPosition, v = g.verticalPosition;
  const horizontal = h.relative === 'page' ? 'PAPER' : h.relative === 'margin' ? 'PAGE' : h.relative === 'column' ? 'COLUMN' : 'PARA';
  const vertical = v.relative === 'page' ? 'PAPER' : v.relative === 'margin' ? 'PAGE' : 'PARA';
  return { treatAsChar: 0, affectLSpacing: 0, flowWithText: 0, allowOverlap: 1, holdAnchorAndSO: 0,
    vertRelTo: vertical, horzRelTo: horizontal, vertAlign: (v.align || 'top').toUpperCase(), horzAlign: (h.align || 'left').toUpperCase(),
    vertOffset: units((v.offset || 0) - (g.margins?.top || 0)), horzOffset: units((h.offset || 0) - (g.margins?.left || 0)) };
}

type PictureRecord = {
  geometry?: ImageFloating;
  description: string;
  raster?: StagedRaster;
  source?: string;
  standalone: boolean;
};
type PictureEvent = { at: number; markup: string; marker: boolean };
const pictureMarkup = '<hp:pic\\b[^>]*>[\\s\\S]*?<\\/hp:pic>';
const pictureAssociationFailed = () => Error('일부 떠 있는 그림을 한글로 저장하지 못했습니다.');

function orderedBindings(events: readonly PictureEvent[], records: ReadonlyMap<string, PictureRecord>) {
  const bindings = new Map<number, PictureRecord>();
  let pending: PictureRecord | undefined;
  for (const { at, markup, marker } of events) {
    if (marker) {
      if (pending) return;
      pending = records.get(markup);
      if (!pending) throw pictureAssociationFailed();
    } else if (pending) {
      bindings.set(at, pending);
      pending = undefined;
    }
  }
  return pending ? undefined : bindings;
}

function pictureEvents(xml: string, prefix: string, marker: string) {
  return [...xml.matchAll(new RegExp(marker + '|' + pictureMarkup, 'g'))]
    .map(match => ({ at: match.index!, markup: match[0], marker: match[0].startsWith(prefix) }));
}

/** The engine can emit all text runs before the pictures they precede.
 * Accept such a group only inside its own paragraph, with every control
 * still carrying the same raster bytes in the same order. */
function groupedBindings(xml: string, marker: string, records: ReadonlyMap<string, PictureRecord>, files: Record<string, Uint8Array>) {
  const manifest = files['Contents/content.hpf'];
  if (!manifest) throw pictureAssociationFailed();
  const assets = new Map<string, Uint8Array>();
  for (const [tag] of strFromU8(manifest).matchAll(/<opf:item\b[^>]*>/g)) {
    const id = /\bid="([^"]+)"/.exec(tag)?.[1];
    const href = /\bhref="([^"]+)"/.exec(tag)?.[1];
    if (id && href && files[href]) assets.set(id, files[href]);
  }
  const sameRaster = (picture: string, record: PictureRecord) => {
    const ref = /<hc:img\b[^>]*\bbinaryItemIDRef="([^"]+)"/.exec(picture)?.[1];
    const bytes = ref && assets.get(ref);
    if (!bytes) return false;
    if (record.raster) {
      const staged = record.raster.marker;
      return bytes.length === staged.length && bytes.every((value, i) => value === staged[i]);
    }
    if (!record.source) return false;
    const decoded = atob(record.source);
    return bytes.length === decoded.length && bytes.every((value, i) => value === decoded.charCodeAt(i));
  };
  const stack: PictureEvent[][] = [], result = new Map<number, PictureRecord>();
  const scanner = new RegExp('<(/?)hp:p\\b[^>]*>|(' + marker + ')|(' + pictureMarkup + ')', 'g');
  for (const match of xml.matchAll(scanner)) {
    if (match[1] === '') {
      if (!match[0].endsWith('/>')) stack.push([]);
    } else if (match[1] === '/') {
      const own = stack.pop();
      if (!own) throw pictureAssociationFailed();
      const ordered = orderedBindings(own, records);
      if (ordered) {
        for (const [at, record] of ordered) result.set(at, record);
        continue;
      }
      const markers = own.filter(event => event.marker), pictures = own.filter(event => !event.marker);
      if (markers.length !== pictures.length) throw pictureAssociationFailed();
      for (const [index, picture] of pictures.entries()) {
        const record = records.get(markers[index].markup);
        if (!record || !sameRaster(picture.markup, record)) throw pictureAssociationFailed();
        result.set(picture.at, record);
      }
    } else {
      const owner = stack.at(-1);
      if (match[2] && !owner) throw pictureAssociationFailed();
      owner?.push({ at: match.index!, markup: match[0], marker: Boolean(match[2]) });
    }
  }
  if (stack.length) throw pictureAssociationFailed();
  return result;
}

/** Preserve picture controls after the native HTML parser has imported their
 * raster data. Inline sentinels associate each picture even inside table/note
 * stories; neither alternative text nor order of archive assets is an ID. */
export function stageHwpPictures(html: string, pasteLimit = 0) {
  const prefix = `WORDSTUDIO-PICTURE-${crypto.randomUUID()}-`;
  const records = new Map<string, PictureRecord>();
  const standaloneIds: string[] = [];
  const staged = html.replace(/<img\b[^>]*>/gi, tag => {
    const source = tag.match(/\bdata-image-floating=("[^"]*"|'[^']*')/i)?.[1];
    const alt = tag.match(/\balt=("[^"]*"|'[^']*')/i)?.[1];
    const large = stageLargeRaster(tag, pasteLimit, records.size + 1);
    const standalone = !source && tag.includes(` ${nativeStandalonePictureAttribute}=""`);
    if (!source && !alt && !large && !standalone) return tag;
    const geometry = source ? imageFloating(entities(source.slice(1, -1))) : undefined;
    if (source && !geometry) throw Error('저장할 그림의 위치 정보가 올바르지 않습니다.');
    const token = `${prefix}${records.size}-END`;
    const encoded = /\ssrc="data:[^;,]+;base64,([^"]+)"/i.exec(tag)?.[1];
    records.set(token, { geometry, description: alt ? entities(alt.slice(1, -1)) : '',
      raster: large?.raster, source: encoded, standalone });
    return `<span>${token}</span>` + (large?.tag ?? tag);
  });
  return { html: staged, standaloneIds, restore(Core: typeof HwpDocument, doc: HwpDocument) {
    if (!records.size) return doc;
    let size = 0, restored = 0;
    const files = unzipSync(doc.exportHwpx(), { filter: file => {
      if ((size += file.originalSize) > documentLimits.documentInflatedBytes) throw Error('한글 그림 변환 데이터가 너무 큽니다.');
      return true;
    } });
    const marker = prefix + '\\d+-END', consumed = new Set<string>();
    for (const name of Object.keys(files).filter(n => /^Contents\/section\d+\.xml$/.test(n))) {
      const source = strFromU8(files[name]), events = pictureEvents(source, prefix, marker);
      const bindings = orderedBindings(events, records) ?? groupedBindings(source, marker, records, files);
      const xml = source.replace(new RegExp(marker + '|' + pictureMarkup, 'g'), (match: string, at: number) => {
        if (match.startsWith(prefix)) {
          if (!records.has(match) || consumed.has(match)) throw pictureAssociationFailed();
          consumed.add(match);
          return '';
        }
        const record = bindings.get(at);
        if (!record) return match;
        const { geometry: g, description, raster, standalone } = record;
        restored++;
        if (raster && (size += raster.bytes - raster.marker.length) > documentLimits.documentInflatedBytes)
          throw Error('한글 그림 변환 데이터가 너무 큽니다.');
        let pic = raster ? restoreLargeRaster(match, files, raster) : match;
        if (standalone) {
          const id = /<hc:img\b[^>]*\bbinaryItemIDRef="(image\d+)"/.exec(pic)?.[1];
          if (id) standaloneIds.push(id);
        }
        if (g) {
          const wrapping = g.wrap!.type === 0 ? g.behindDocument ? 'BEHIND_TEXT' : 'IN_FRONT_OF_TEXT' : g.wrap!.type === 3 ? 'TOP_AND_BOTTOM' : 'SQUARE';
          pic = pic.replace(/^<hp:pic\b[^>]*>/, tag => attributes(tag, { textWrap: wrapping, textFlow: ({ left: 'LEFT_ONLY', right: 'RIGHT_ONLY', largest: 'LARGEST_ONLY' } as Record<string, string>)[g.wrap!.side || ''] || 'BOTH_SIDES' }))
            .replace(/<hp:pos\b[^>]*\/>/, tag => attributes(tag, position(g)))
            .replace(/<hp:outMargin\b[^>]*\/>/, tag => attributes(tag, Object.fromEntries(['left', 'right', 'top', 'bottom'].map(key => [key, units((g.margins as Record<string, number> | undefined)?.[key])]))));
        }
        if (description) pic = pic.replace('</hp:pic>', () => `<hp:shapeComment>${escape(description)}</hp:shapeComment></hp:pic>`);
        return pic;
      }).replace(/<hp:linesegarray\b[^>]*>[\s\S]*?<\/hp:linesegarray>/g, '');
      files[name] = strToU8(xml);
    }
    if (restored !== records.size || consumed.size !== records.size) throw pictureAssociationFailed();
    return new Core(zipSync(files, { level: 1 }));
  } };
}

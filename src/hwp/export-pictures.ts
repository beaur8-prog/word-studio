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

/** Preserve picture controls after the native HTML parser has imported their
 * raster data. Inline sentinels associate each picture even inside table/note
 * stories; neither alternative text nor order of archive assets is an ID. */
export function stageHwpPictures(html: string, pasteLimit = 0) {
  const prefix = `WORDSTUDIO-PICTURE-${crypto.randomUUID()}-`;
  type PictureRecord = { geometry?: ImageFloating; description: string; raster?: StagedRaster; standalone: boolean };
  const records = new Map<string, PictureRecord>();
  const standaloneIds: string[] = [];
  const staged = html.replace(/<img\b[^>]*>/gi, tag => {
    const source = tag.match(/\bdata-image-floating=("[^"]*"|'[^']*')/i)?.[1];
    const alt = tag.match(/\balt=("[^"]*"|'[^']*')/i)?.[1];
    const large = stageLargeRaster(tag, pasteLimit, records.size + 1);
    if (!source && !alt && !large) return tag;
    const geometry = source ? imageFloating(entities(source.slice(1, -1))) : undefined;
    if (source && !geometry) throw Error('저장할 그림의 위치 정보가 올바르지 않습니다.');
    const token = `${prefix}${records.size}-END`;
    const standalone = !geometry && large?.raster.format === 'bmp' && tag.includes(` ${nativeStandalonePictureAttribute}=""`);
    records.set(token, { geometry, description: alt ? entities(alt.slice(1, -1)) : '', raster: large?.raster, standalone });
    return `<span>${token}</span>` + (large?.tag ?? tag);
  });
  return { html: staged, standaloneIds, restore(Core: typeof HwpDocument, doc: HwpDocument) {
    if (!records.size) return doc;
    let size = 0, restored = 0;
    const files = unzipSync(doc.exportHwpx(), { filter: file => {
      if ((size += file.originalSize) > documentLimits.documentInflatedBytes) throw Error('한글 그림 변환 데이터가 너무 큽니다.');
      return true;
    } });
    for (const name of Object.keys(files).filter(n => /^Contents\/section\d+\.xml$/.test(n))) {
      let pending: PictureRecord | undefined;
      const xml = strFromU8(files[name]).replace(new RegExp(`${prefix}\\d+-END|<hp:pic\\b[^>]*>[\\s\\S]*?<\\/hp:pic>`, 'g'), match => {
        if (match.startsWith(prefix)) {
          if (pending) throw Error('일부 떠 있는 그림을 한글로 저장하지 못했습니다.');
          pending = records.get(match); if (!pending) throw Error('한글 그림 변환 범위가 올바르지 않습니다.'); return '';
        }
        if (!pending) return match;
        const { geometry: g, description, raster, standalone } = pending; pending = undefined; restored++;
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
      if (pending) throw Error('일부 떠 있는 그림을 한글로 저장하지 못했습니다.');
      files[name] = strToU8(xml);
    }
    if (restored !== records.size) throw Error('일부 떠 있는 그림을 한글로 저장하지 못했습니다.');
    return new Core(zipSync(files, { level: 1 }));
  } };
}

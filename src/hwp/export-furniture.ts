import type { HwpDocument } from '@rhwp/core';
import { strFromU8, strToU8, zipSync, unzipSync } from 'fflate';
import type { PageSettings } from '../types';
import documentLimits from '../../electron/document-limits.json';
import { elementRanges } from './markup-ranges';
import { entities } from './cell-content';

const escape = (text: string) => text.replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]!));
const textOf = (xml: string) => entities(xml.replace(/<[^>]*>/g, ''));
const withoutLines = (xml: string) => xml.replace(/<hp:linesegarray\b[^>]*>[\s\S]*?<\/hp:linesegarray>/g, '');

/** Put stories through the same paste as body content so native font, table,
 * image and live-field references stay valid when moved to header controls. */
export function stageHwpFurniture(html: string, settings: PageSettings[]) {
  const prefix = `WORDSTUDIO-FURNITURE-${crypto.randomUUID()}-`;
  const stories: { section: number; kind: 'header' | 'footer'; apply: string; html: string; alignment?: 'center' | 'right'; xml?: string }[] = [];
  const replaced = new Set<number>();
  settings.forEach((page, section) => {
    if (!page.header && !page.footer && !page.headerHtml && !page.footerHtml && !page.pageNumbers && !page.differentOddEven) return;
    replaced.add(section);
    const types = page.differentOddEven ? ['ODD', 'EVEN'] : ['BOTH'];
    for (const apply of types) {
      const even = apply === 'EVEN';
      const header = even ? page.evenHeaderHtml || '' : page.headerHtml || '';
      const footer = even ? page.evenFooterHtml || '' : page.footerHtml || '';
      const numbers = even ? page.evenPageNumbers : page.pageNumbers;
      for (const kind of ['header', 'footer'] as const) {
        const rich = kind === 'header' ? header : footer;
        const text = even ? page[kind === 'header' ? 'evenHeader' : 'evenFooter'] : page[kind];
        let content = rich || `<p style="font-size:9pt;line-height:140%;margin:0">${escape(text || '')}</p>`;
        if (kind === 'footer' && numbers && !/data-page-number=/.test(header + footer)) {
          if (!rich && !text) content = '';
          content += '<p style="text-align:center;font-size:9pt;line-height:140%;margin:0"><span data-page-number="PAGE">1</span></p>';
        }
        // Explicit empty stories prevent inheritance from the preceding section.
        const tables = elementRanges(content, 'table');
        let alignment: 'center' | 'right' | undefined;
        if (tables.length === 1 && !textOf(content.slice(0, tables[0].start) + content.slice(tables[0].end)).trim() && /margin-left:\s*auto/.test(tables[0].open))
          alignment = /margin-right:\s*auto/.test(tables[0].open) ? 'center' : 'right';
        stories.push({ section, kind, apply, html: content, alignment });
      }
    }
  });
  const staged = html + stories.map((story, i) => `<p>${prefix}${i}-START</p>${story.html}<p>${prefix}${i}-END</p>`).join('');
  return { html: staged, restore(Core: typeof HwpDocument, doc: HwpDocument) {
    if (!stories.length) return doc;
    let size = 0;
    const files = unzipSync(doc.exportHwpx(), { filter: file => {
      if ((size += file.originalSize) > documentLimits.documentInflatedBytes) throw Error('한글 머리말 변환 데이터가 너무 큽니다.');
      return true;
    } });
    for (const name of Object.keys(files).filter(name => /^Contents\/section\d+\.xml$/.test(name))) {
      let xml = strFromU8(files[name]), active = -1, start = 0, from = 0;
      const removals: { start: number; end: number }[] = [];
      for (const p of elementRanges(xml, 'hp:p')) {
        const text = textOf(p.inner);
        const opening = stories.findIndex((_, i) => text === `${prefix}${i}-START`);
        if (opening >= 0) { if (active >= 0) throw Error('한글 머리말 변환 범위가 올바르지 않습니다.'); active = opening; start = p.start; from = p.end; }
        else if (active >= 0 && text === `${prefix}${active}-END`) {
          stories[active].xml = withoutLines(xml.slice(from, p.start)); removals.push({ start, end: p.end }); active = -1;
        }
      }
      if (active >= 0) throw Error('한글 머리말 변환 범위가 올바르지 않습니다.');
      for (const r of removals.reverse()) xml = xml.slice(0, r.start) + xml.slice(r.end);
      files[name] = strToU8(withoutLines(xml));
    }
    if (stories.some(story => !story.xml)) throw Error('일부 머리말·꼬리말을 한글로 저장하지 못했습니다.');
    for (const section of replaced) {
      const name = `Contents/section${section}.xml`;
      let xml = strFromU8(files[name]).replace(/<hp:(header|footer)\b[^>]*>[\s\S]*?<\/hp:\1>/g, '');
      const page = settings[section], margins = page.margins ?? { top: page.margin, bottom: page.margin };
      const controls = stories.filter(story => story.section === section).map((story, i) => {
        const margin = margins[story.kind === 'header' ? 'top' : 'bottom'];
        const distance = page[story.kind === 'header' ? 'headerDistance' : 'footerDistance'] ?? Math.max(0, margin / 2 - 3);
        const height = Math.round(Math.max(0, margin - distance) * 7200 / 25.4);
        return `<hp:ctrl><hp:${story.kind} id="${i}" applyPageType="${story.apply}"><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="${story.kind === 'footer' ? 'BOTTOM' : 'TOP'}" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="${height}" hasTextRef="0" hasNumRef="0">${story.xml}</hp:subList></hp:${story.kind}></hp:ctrl>`;
      }).join('');
      xml = xml.replace(/(<hp:run\b[^>]*>)/, '$1' + controls);
      files[name] = strToU8(xml);
    }
    const restored = new Core(zipSync(files, { level: 1 }));
    try {
      for (const story of stories) if (story.alignment) {
        const para = elementRanges(story.xml!, 'hp:p').findIndex(p => p.inner.includes('<hp:tbl'));
        if (para >= 0) restored.applyParaFormatInHf(story.section, story.kind === 'header', story.apply === 'EVEN' ? 1 : story.apply === 'ODD' ? 2 : 0, para, JSON.stringify({ alignment: story.alignment }));
      }
      return restored;
    } catch (error) { restored.free(); throw error; }
  } };
}

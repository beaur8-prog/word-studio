import type { HwpDocument } from '@rhwp/core';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { tableAnchor, type TableAnchor } from '../table-anchor-data';
import { elementRanges } from './markup-ranges';
import { entities } from './cell-content';
import documentLimits from '../../electron/document-limits.json';

const set = (tag: string, values: Record<string, string | number>) => {
  for (const [key, value] of Object.entries(values)) {
    const pattern = new RegExp(`\\b${key}="[^"]*"`), text = `${key}="${value}"`;
    tag = pattern.test(tag) ? tag.replace(pattern, text) : tag.replace(/\s*\/?>$/, end => ' ' + text + end);
  }
  return tag;
};
/** Sentinels in an existing cell identify its real control after HTML import;
 * no marker paragraph or original text is removed. */
export function stageHwpTables(html: string) {
  const prefix = `WORDSTUDIO-TABLE-${crypto.randomUUID()}-`, records = new Map<string, TableAnchor>();
  let staged = html;
  for (const table of elementRanges(html, 'table').reverse()) {
    const value = table.open.match(/\bdata-table-anchor=(["'])([\s\S]*?)\1/)?.[2];
    const anchor = value ? tableAnchor(entities(value)) : undefined; if (!anchor?.overlay) continue;
    let depth = 0, at = -1;
    for (const m of table.inner.matchAll(/<\/?(?:table|p)\b[^>]*>/g)) {
      if (/^<table\b/.test(m[0])) depth++;
      else if (/^<\/table\b/.test(m[0])) depth--;
      else if (!depth && /^<p\b/.test(m[0])) { at = m.index! + m[0].length; break; }
    }
    if (at < 0) continue;
    const token = `${prefix}${records.size}-END`; records.set(token, anchor);
    const inner = table.inner.slice(0, at) + `<span>${token}</span>` + table.inner.slice(at);
    staged = staged.slice(0, table.start) + table.open + inner + '</table>' + staged.slice(table.end);
  }
  return { html: staged, restore(Core: typeof HwpDocument, doc: HwpDocument) {
    if (!records.size) return doc;
    let size = 0, restored = 0;
    const files = unzipSync(doc.exportHwpx(), { filter: file => {
      if ((size += file.originalSize) > documentLimits.documentInflatedBytes) throw Error('한글 표 배치 데이터가 너무 큽니다.');
      return true;
    } });
    for (const name of Object.keys(files).filter(n => /^Contents\/section\d+\.xml$/.test(n))) {
      let xml = strFromU8(files[name]);
      for (const table of elementRanges(xml, 'hp:tbl').reverse()) {
        const tokens = table.inner.match(new RegExp(`${prefix}\\d+-END`, 'g')) || [];
        if (!tokens.length) continue;
        if (tokens.length !== 1) throw Error('한글 표의 배치 위치가 모호합니다.');
        const anchor = records.get(tokens[0]); if (!anchor) throw Error('한글 표의 배치 정보를 찾지 못했습니다.');
        const open = set(table.open, { textWrap: 'IN_FRONT_OF_TEXT' });
        let found = false;
        const inner = table.inner.replace(/<hp:pos\b[^>]*\/>/, tag => {
          found = true;
          return set(tag, { treatAsChar: 0, affectLSpacing: 0, flowWithText: 0, allowOverlap: 1, holdAnchorAndSO: 0,
            vertRelTo: 'PAPER', horzRelTo: 'PAPER', vertAlign: 'BOTTOM', horzAlign: 'LEFT',
            vertOffset: Math.round(anchor.y * 75), horzOffset: Math.round(anchor.x * 75) });
        // The recorded pixel anchor already includes the original outer gap.
        // The HTML bridge otherwise adds its default 283-HWPUNIT margins on
        // every save, moving a bottom-anchored table upward by another 3.8px.
        }).replace(/<hp:outMargin\b[^>]*\/>/, tag => set(tag, { left: 0, right: 0, top: 0, bottom: 0 })).replace(tokens[0], '');
        if (!found) throw Error('한글 표의 위치 속성을 찾지 못했습니다.');
        xml = xml.slice(0, table.start) + open + inner + '</hp:tbl>' + xml.slice(table.end); restored++;
      }
      files[name] = strToU8(xml.replace(/<hp:linesegarray\b[^>]*>[\s\S]*?<\/hp:linesegarray>/g, ''));
    }
    if (restored !== records.size) throw Error('일부 떠 있는 표를 한글로 저장하지 못했습니다.');
    return new Core(zipSync(files, { level: 1 }));
  } };
}

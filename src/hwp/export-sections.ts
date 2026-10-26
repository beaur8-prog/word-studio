import type { HwpDocument } from '@rhwp/core';
import { strFromU8, strToU8, zipSync, unzipSync } from 'fflate';
import type { PageSettings } from '../types';
import documentLimits from '../../electron/document-limits.json';
import { elementRanges } from './markup-ranges';
import { entities } from './cell-content';

/** Keep one shared style/image dictionary while partitioning the native paste
 * into real sections. Separate HTML pastes would duplicate or remap those IDs. */
export function stageHwpSections(html: string, page: PageSettings) {
  const prefix = `WORDSTUDIO-SECTION-${crypto.randomUUID()}-`, settings = [page];
  const staged = html.replace(/<div\b[^>]*\bdata-section-page=(["'])([\s\S]*?)\1[^>]*>\s*<\/div>/gi, (_, _quote, value: string) => {
    const next = JSON.parse(entities(value)) as PageSettings;
    if (!next || !['A4', 'Letter', 'Custom'].includes(next.size) || !['portrait', 'landscape'].includes(next.orientation)) throw Error('한글 구역의 용지 설정이 올바르지 않습니다.');
    return `<p>${prefix}${settings.push(next) - 1}-END</p>`;
  });
  return { html: staged, restore(Core: typeof HwpDocument, doc: HwpDocument, applyPage: (doc: HwpDocument, page: PageSettings, section: number) => void) {
    if (settings.length === 1) return doc;
    let size = 0;
    const files = unzipSync(doc.exportHwpx(), { filter: file => {
      if ((size += file.originalSize) > documentLimits.documentInflatedBytes) throw Error('한글 구역 변환 데이터가 너무 큽니다.');
      return true;
    } });
    const xml = strFromU8(files['Contents/section0.xml']);
    const paragraphs = elementRanges(xml, 'hp:p');
    const sectionDef = xml.match(/<hp:secPr\b[^>]*>[\s\S]*?<\/hp:secPr>/)?.[0];
    if (!sectionDef || !paragraphs.length) throw Error('한글 구역 정의를 찾지 못했습니다.');
    const start = xml.slice(0, paragraphs[0].start), end = xml.slice(paragraphs.at(-1)!.end);
    const groups: string[][] = [[]]; let found = 0;
    for (const p of paragraphs) {
      const text = [...p.inner.matchAll(/<hp:t\b[^>]*>([\s\S]*?)<\/hp:t>/g)].map(m => entities(m[1])).join('');
      if (text.startsWith(prefix)) {
        if (text !== `${prefix}${++found}-END`) throw Error('한글 구역 변환 위치가 올바르지 않습니다.');
        groups.push([]); continue;
      }
      groups.at(-1)!.push(xml.slice(p.start, p.end));
    }
    if (found !== settings.length - 1) throw Error('표·각주 안의 구역 나누기는 한글로 저장할 수 없습니다. DOCX 또는 백업으로 저장하세요.');
    // Empty sections are intentional blank pages, not discarded markers.
    const empty = '<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:t/></hp:run></hp:p>';
    groups.forEach((group, index) => {
      if (!group.length) group.push(empty);
      let body = group.join('').replace(/<hp:secPr\b[^>]*>[\s\S]*?<\/hp:secPr>/g, '')
        .replace(/<hp:linesegarray\b[^>]*>[\s\S]*?<\/hp:linesegarray>/g, '')
        .replace(/<hp:(header|footer)\b[^>]*>[\s\S]*?<\/hp:\1>/g, '');
      body = body.replace(/(<hp:run\b[^>]*>)/, '$1' + sectionDef);
      files[`Contents/section${index}.xml`] = strToU8(start + body + end);
    });
    let content = strFromU8(files['Contents/content.hpf']);
    content = content.replace(/<opf:item\b[^>]*\bid="section\d+"[^>]*\/>/g, '')
      .replace(/<opf:itemref\b[^>]*\bidref="section\d+"[^>]*\/>/g, '')
      .replace('</opf:manifest>', groups.map((_, i) => `<opf:item id="section${i}" href="Contents/section${i}.xml" media-type="application/xml"/>`).join('') + '</opf:manifest>')
      .replace('</opf:spine>', groups.map((_, i) => `<opf:itemref idref="section${i}" linear="yes"/>`).join('') + '</opf:spine>');
    files['Contents/content.hpf'] = strToU8(content);
    files['Contents/header.xml'] = strToU8(strFromU8(files['Contents/header.xml']).replace(/\bsecCnt="\d+"/, `secCnt="${groups.length}"`));
    const restored = new Core(zipSync(files, { level: 1 }));
    try {
      if (restored.getSectionCount() !== settings.length) throw Error('한글 구역을 다시 읽지 못했습니다.');
      settings.forEach((page, section) => {
        if (section) restored.applyParaFormat(section, 0, JSON.stringify({ pageBreakBefore: false }));
        applyPage(restored, page, section);
      });
      return restored;
    } catch (error) { restored.free(); throw error; }
  } };
}

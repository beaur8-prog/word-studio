import type { HwpDocument } from '@rhwp/core';
import { strFromU8, strToU8, zipSync, unzipSync } from 'fflate';
import type { PageSettings } from '../types';
import documentLimits from '../../electron/document-limits.json';
import { elementRanges } from './markup-ranges';
import { entities } from './cell-content';

type Columns = { count: number; space: number; balance?: boolean };
type Change = { columns?: Columns; columnBreak?: boolean };
const attr = (tag: string, name: string) => tag.match(new RegExp(`\\b${name}=["']([^"']*)["']`))?.[1];
function paragraphText(inner: string) {
  let depth = 0, text = '';
  for (const m of inner.matchAll(/<\/?hp:p\b[^>]*>|<hp:t\b[^>]*>([\s\S]*?)<\/hp:t>/g)) {
    if (m[0].startsWith('</hp:p')) depth--;
    else if (m[0].startsWith('<hp:p')) { if (!m[0].endsWith('/>')) depth++; }
    else if (!depth) text += entities(m[1].replace(/<[^>]*>/g, ''));
  }
  return text;
}

/** Native clipboard HTML discards column groups and explicit column breaks.
 * Stage structural markers, then replace them with normalized HWPX controls.
 * All styles/images still come from the same native HTML paste. */
export function stageHwpColumns(html: string, page: PageSettings) {
  const prefix = `WORDSTUDIO-COLUMNS-${crypto.randomUUID()}-`, changes: Change[] = [];
  let columns: Columns = page.columns ?? { count: 1, space: 0 }, output = '', cursor = 0;
  const stack: (Columns | undefined)[] = [];
  const marker = (change: Change) => `<p>${prefix}${changes.push(change) - 1}-END</p>`;
  for (const token of html.matchAll(/<\/?div\b[^>]*>/g)) {
    const tag = token[0]; output += html.slice(cursor, token.index); cursor = token.index! + tag.length;
    if (tag.startsWith('</')) {
      const previous = stack.pop();
      if (previous) { columns = previous; output += marker({ columns }); }
      else output += tag;
      continue;
    }
    const count = Number(attr(tag, 'data-column-count'));
    if (Number.isInteger(count) && count >= 1 && count <= 12) {
      stack.push(columns);
      columns = { count, space: Math.max(0, Math.min(1000, Number(attr(tag, 'data-column-space')) || 0)), balance: attr(tag, 'data-column-balance') !== 'false' };
      output += marker({ columns });
    } else if (attr(tag, 'data-column-break') !== undefined) {
      stack.push(undefined); output += marker({ columnBreak: true });
      // Keep a matching empty wrapper for the ordinary closing-tag path.
      output += '<div>';
    } else { stack.push(undefined); output += tag; }
  }
  output += html.slice(cursor);
  return { html: output, restore(Core: typeof HwpDocument, doc: HwpDocument) {
    if (!changes.length) return doc;
    let size = 0;
    const files = unzipSync(doc.exportHwpx(), { filter: file => {
      if ((size += file.originalSize) > documentLimits.documentInflatedBytes) throw Error('한글 다단 변환 데이터가 너무 큽니다.');
      return true;
    } });
    const name = 'Contents/section0.xml', xml = strFromU8(files[name]);
    let result = '', at = 0, pending: Change = {}, pageBreak = false, found = 0, carriedRuns = '';
    let paragraphIndex = 0;
    const pageBreaks: number[] = [];
    let keptParagraphs = 0;
    for (const p of elementRanges(xml, 'hp:p')) {
      const sourceIndex = paragraphIndex++;
      result += xml.slice(at, p.start); at = p.end;
      const text = paragraphText(p.inner);
      const match = text.startsWith(prefix) ? text.slice(prefix.length).match(/^(\d+)-END$/) : null;
      if (match) {
        const properties = JSON.parse(doc.getParaPropertiesAt(0, sourceIndex));
        const change = changes[Number(match[1])]; if (!change) throw Error('한글 다단 변환 위치가 올바르지 않습니다.');
        pending = { ...pending, ...change };
        // applyParaFormat stores pageBreakBefore in the paragraph shape, not
        // the XML paragraph's pageBreak attribute. A removed marker must carry
        // both forms of the break to the following real paragraph.
        pageBreak ||= attr(p.open, 'pageBreak') === '1' || properties.pageBreakBefore;
        found++;
        // The initial marker may own the section definition and default
        // furniture added by setHwpPage. Move those runs with the marker.
        carriedRuns += p.inner.replace(new RegExp(`<hp:t\\b[^>]*>${prefix}${match[1]}-END<\\/hp:t>`), '')
          .replace(/<hp:linesegarray\b[^>]*>[\s\S]*?<\/hp:linesegarray>/g, '')
          .replace(/<hp:run\b[^>]*>\s*<\/hp:run>/g, '');
        continue;
      }
      let paragraph = xml.slice(p.start, p.end);
      if (carriedRuns) paragraph = paragraph.replace(p.open, p.open + carriedRuns);
      if (pageBreak) {
        paragraph = paragraph.replace(/\bpageBreak="[^"]*"/, 'pageBreak="1"');
        pageBreaks.push(keptParagraphs);
      }
      if (pending.columnBreak) paragraph = paragraph.replace(/\bcolumnBreak="[^"]*"/, 'columnBreak="1"');
      if (pending.columns) {
        const c = pending.columns;
        paragraph = paragraph.replace(/<hp:ctrl>\s*<hp:colPr\b[^>]*\/>\s*<\/hp:ctrl>/g, '');
        const control = `<hp:ctrl><hp:colPr id="" type="${c.balance ? 'BalancedNewspaper' : 'NEWSPAPER'}" layout="LEFT" colCount="${c.count}" sameSz="1" sameGap="${Math.round(c.space * 75)}"/></hp:ctrl>`;
        paragraph = paragraph.replace(/(<hp:run\b[^>]*>)/, '$1' + control);
      }
      result += paragraph; keptParagraphs++; pending = {}; pageBreak = false; carriedRuns = '';
    }
    result += xml.slice(at);
    if (found !== changes.length) throw Error('표·각주 안의 다단 설정은 한글로 저장할 수 없습니다. DOCX 또는 백업으로 저장하세요.');
    files[name] = strToU8(result.replace(/<hp:linesegarray\b[^>]*>[\s\S]*?<\/hp:linesegarray>/g, ''));
    const restored = new Core(zipSync(files, { level: 1 }));
    try {
      for (const index of pageBreaks) restored.applyParaFormat(0, index, JSON.stringify({ pageBreakBefore: true }));
      return restored;
    } catch (error) { restored.free(); throw error; }
  } };
}

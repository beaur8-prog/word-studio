import type JSZip from 'jszip';
import { pageMargins, paperDimensions, type Comment, type PageSettings } from './types';
import { W, all, val, child, ancestor, docxFormat } from './docx-format';
import { overlapInstruction } from './word-overlap';
import { symbolFont, symbolElement } from './document-symbols';
import { normalizeVmlImages } from './docx-images';
import { drawingFloating } from './image-floating';
const parse = (text: string) => new DOMParser().parseFromString(text, 'application/xml');
// A standard paragraph style identifies the otherwise invisible separator
// required by Word between adjacent tables and after a table inside a cell.
export const TABLE_GAP_STYLE = 'WordStudioTableGap';

/** Temporary bookmarks let Mammoth preserve exact positions, including repeated text. */
export async function prepareDocxMetadata(zip: JSZip, body: Document, pages: PageSettings[]) {
  const format = await docxFormat(zip);
  // An empty legacy floating-frame anchor has no inline line box. Treating
  // it as an ordinary paragraph after its table creates a spurious page.
  for (const p of all(body, 'p')) if (child(child(p, 'pPr'), 'framePr') && !ancestor(p, 'tc') &&
    Array.from(p.children).every(c => c.localName === 'pPr')) p.remove();
  const comments: Comment[] = [];
  const idMap = new Map<string, string>();
  const commentFile = zip.file('word/comments.xml');
  const resolved = new Set<string>();
  const extended = zip.file('word/commentsExtended.xml');
  if (extended) for (const el of Array.from(parse(await extended.async('string')).getElementsByTagName('*'))) {
    const attrs = Array.from(el.attributes);
    if (attrs.some(a => a.localName === 'done' && a.value === '1')) {
      const para = attrs.find(a => a.localName === 'paraId')?.value;
      if (para) resolved.add(para);
    }
  }
  if (commentFile) for (const el of all(parse(await commentFile.async('string')), 'comment')) {
    const id = crypto.randomUUID();
    idMap.set(val(el, 'id'), id);
    const date = val(el, 'date');
    const lastParagraph = all(el, 'p').at(-1);
    const paraId = lastParagraph && Array.from(lastParagraph.attributes).find(a => a.localName === 'paraId')?.value;
    comments.push({ id, quote: '', text: all(el, 'p').map(p => all(p, 't').map(t => t.textContent).join('')).join('\n'),
      createdAt: Number.isFinite(Date.parse(date)) ? new Date(date).toISOString() : new Date().toISOString(),
      resolved: Boolean(paraId && resolved.has(paraId)), author: val(el, 'author'), initials: val(el, 'initials') });
  }
  let sequence = 900000;
  const numbering = new WeakMap<Element, ReturnType<typeof format.numbering>>();
  for (const p of all(body, 'p')) numbering.set(p, format.numbering(p));
  const markers = new Map<string, { kind: string; id?: string; parts?: string[]; paragraph?: Element; cell?: Element; style?: string; eastAsianStyle?: string; nextPage?: PageSettings; width?: number; height?: number; alt?: string; floating?: string }>();
  const bookmark = (name: string) => {
    const fragment = body.createDocumentFragment();
    const id = String(++sequence);
    const start = body.createElementNS(W, 'w:bookmarkStart');
    start.setAttributeNS(W, 'w:id', id); start.setAttributeNS(W, 'w:name', name);
    const end = body.createElementNS(W, 'w:bookmarkEnd'); end.setAttributeNS(W, 'w:id', id);
    fragment.append(start, end);
    return fragment;
  };
  for (const p of all(body, 'p')) {
    let field: Element[] = [], depth = 0;
    const replace = (elements: Element[], instruction: string) => {
      const pageField = instruction.trim().match(/^(PAGE|NUMPAGES)\b/i);
      if (pageField) {
        const name = `ws_marker_${sequence}`; markers.set(name, { kind: 'pageField', id: pageField[1].toUpperCase() });
        elements[0].before(bookmark(name)); elements.forEach(el => el.remove()); return;
      }
      const parts = overlapInstruction(instruction); if (!parts.length) return;
      const name = `ws_marker_${sequence}`;
      const run = elements.flatMap(el => el.localName === 'r' ? [el] : all(el, 'r'))
        .find(r => all(r, 'instrText').some(t => t.textContent?.trim()) || all(r, 't').length);
      markers.set(name, { kind: 'overlap', parts, style: run ? format.run(run, 'eastAsia') : '' });
      elements[0].before(bookmark(name)); elements.forEach(el => el.remove());
    };
    for (const el of Array.from(p.children)) {
      if (el.localName === 'fldSimple') { replace([el], val(el, 'instr')); continue; }
      const kinds = all(el, 'fldChar').map(f => val(f, 'fldCharType'));
      if (kinds.includes('begin')) { if (!depth) field = []; depth++; }
      if (depth) field.push(el);
      if (kinds.includes('end') && depth && --depth === 0)
        replace(field, field.flatMap(e => all(e, 'instrText')).map(e => e.textContent).join(''));
    }
  }
  for (const kind of ['commentRangeStart', 'commentRangeEnd']) for (const el of all(body, kind)) {
    const id = idMap.get(val(el, 'id'));
    if (!id) continue;
    const name = `ws_marker_${sequence}`;
    markers.set(name, { kind: kind === 'commentRangeStart' ? 'start' : 'end', id });
    el.replaceWith(bookmark(name));
  }
  normalizeVmlImages(body);
  for (const drawing of all(body, 'drawing')) {
    const extent = drawing.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', 'extent')[0];
    const width = Number(extent?.getAttribute('cx')) / 9525, height = Number(extent?.getAttribute('cy')) / 9525;
    const name = `ws_marker_${sequence}`;
    const properties = drawing.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', 'docPr')[0];
    const alt = properties?.getAttribute('descr')?.trim() || properties?.getAttribute('title') || '';
    if (properties) properties.setAttribute('descr', name);
    const floating = drawingFloating(drawing);
    markers.set(name, { kind: 'image', alt, ...(floating ? { floating: JSON.stringify(floating) } : {}), ...(width > 0 && height > 0 && width <= 10000 && height <= 10000 ? { width, height } : {}) });
    drawing.before(bookmark(name));
  }
  for (const tab of all(body, 'tab').filter(el => el.parentElement?.localName === 'r')) {
    const name = `ws_marker_${sequence}`; markers.set(name, { kind: 'tab' }); tab.replaceWith(bookmark(name));
  }
  for (const sym of all(body, 'sym')) {
    const name = `ws_marker_${sequence}`;
    markers.set(name, { kind: 'symbol', id: val(sym, 'font'), parts: [String(parseInt(val(sym, 'char'), 16))] });
    sym.replaceWith(bookmark(name));
  }
  for (const run of all(body, 'r')) {
    const style = format.run(run);
    if (!style) continue;
    const start = `ws_marker_${sequence}`;
    markers.set(start, { kind: 'runStart', style, eastAsianStyle: format.run(run, 'eastAsia') }); run.before(bookmark(start));
    const end = `ws_marker_${sequence}`;
    markers.set(end, { kind: 'runEnd' }); run.after(bookmark(end));
  }
  for (const br of all(body, 'br').filter(el => ['page', 'column'].includes(val(el, 'type')))) {
    const name = `ws_marker_${sequence}`;
    markers.set(name, { kind: val(br, 'type') === 'column' ? 'columnBreak' : 'pageBreak' }); br.replaceWith(bookmark(name));
  }
  let section = 0;
  for (const paragraph of all(body, 'p')) {
    const name = `ws_marker_${sequence}`;
    const isEnd = all(paragraph, 'sectPr').length > 0;
    const parent = paragraph.parentElement;
    const cell = parent?.localName === 'tc' && Array.from(parent.children).find(c => c.localName === 'p') === paragraph ? parent : undefined;
    markers.set(name, { kind: 'paragraph', paragraph, cell,
      ...(isEnd ? { nextPage: pages[++section] } : {}) });
    const properties = Array.from(paragraph.children).find(e => e.localName === 'pPr');
    paragraph.insertBefore(bookmark(name), properties ? properties.nextSibling : paragraph.firstChild);
  }
  zip.file('word/document.xml', new XMLSerializer().serializeToString(body));
  return {
    buffer: await zip.generateAsync({ type: 'arraybuffer' }), comments,
    decorate(html: string) {
      // Anchor markers nested inside hyperlinks make the HTML parser close the
      // outer link. Convert markers before parsing, not after damage is done.
      const template = document.createElement('template');
      template.innerHTML = html.replace(/<a\s+id="(ws_marker_\d+)"\s*><\/a>/g, '<span data-ws-marker="$1"></span>')
        .replace(/<a\s+id="([^"]+)"\s*><\/a>/g, '<span data-docx-bookmark="$1"></span>');
      const sizedTables = new Map<HTMLTableElement, Element>();
      const tableGaps = new Map<HTMLElement, number>();
      for (const el of template.content.querySelectorAll('img')) {
        const name = el.getAttribute('alt') || '', marker = markers.get(name);
        if (marker?.kind !== 'image') continue;
        const anchor = [...template.content.querySelectorAll('[data-ws-marker]')].find(a => a.getAttribute('data-ws-marker') === name);
        el.setAttribute('data-inline-image', 'true'); el.setAttribute('alt', marker.alt || '');
        if (marker.floating) el.setAttribute('data-image-floating', marker.floating);
        if (marker.width && marker.height) { el.setAttribute('width', String(marker.width)); el.setAttribute('height', String(marker.height)); }
        if (anchor) {
          const original = el.parentElement; anchor.after(el);
          if (original && original !== el.parentElement && original.matches('p') && !original.textContent && !original.querySelector('img,[data-ws-marker],[data-docx-bookmark]')) original.remove();
        }
      }
      // Attach paragraph properties before deleting marker anchors.
      for (const anchor of template.content.querySelectorAll('[data-ws-marker]')) {
        const marker = markers.get(anchor.getAttribute('data-ws-marker')!);
        if (marker?.kind !== 'paragraph') continue;
        let p = anchor.closest<HTMLElement>('p,h1,h2,h3,h4,h5,h6,li');
        if (!p || !marker.paragraph) continue;
        if (p.tagName === 'LI') {
          const paragraph = document.createElement('p');
          while (p.firstChild && !(p.firstChild instanceof Element && p.firstChild.matches('ul,ol'))) paragraph.append(p.firstChild);
          p.prepend(paragraph); p = paragraph;
        }
        const list = numbering.get(marker.paragraph), item = p.closest('li');
        if (list && item) {
          for (const key of ['id', 'level', 'format', 'pattern', 'label'] as const) item.setAttribute(`data-list-${key}`, String(list[key]));
          const parentList = item.parentElement;
          if (parentList?.tagName === 'OL' && parentList.firstElementChild === item) parentList.setAttribute('start', String(list.value));
        }
        const cell = p.closest<HTMLElement>('td,th');
        if (cell && marker.cell) {
          format.cell(cell, marker.cell);
          const sourceTable = ancestor(marker.cell, 'tbl'), table = cell.closest('table');
          if (sourceTable && table && !sizedTables.has(table)) {
            sizedTables.set(table, sourceTable);
            table.style.marginLeft = `${Number(val(child(child(sourceTable, 'tblPr'), 'tblInd'), 'w')) / 15}px`;
            const alignment = val(child(child(sourceTable, 'tblPr'), 'jc'), 'val');
            if (['center', 'right', 'end'].includes(alignment)) {
              table.style.marginLeft = 'auto'; table.style.marginRight = alignment === 'center' ? 'auto' : '0px';
            }
            table.style.marginBottom = '0px';
            const floating = child(child(sourceTable, 'tblPr'), 'tblpPr');
            // Page-relative tables spanning the text area can wrap only above
            // and below. Narrow/sided floats require a separate wrap layout.
            const grid = child(sourceTable, 'tblGrid');
            const gridWidth = (grid ? all(grid, 'gridCol') : []).reduce((sum, col) => sum + Number(val(col, 'w')), 0) / 15;
            const page = pages[0], margins = pageMargins(page), bodyWidth = (paperDimensions(page)[0] - margins.left - margins.right) * 96 / 25.4;
            if (floating && val(floating, 'horzAnchor') === 'page' && val(floating, 'vertAnchor') === 'page' &&
              !val(floating, 'tblpXSpec') && !val(floating, 'tblpYSpec') && gridWidth >= bodyWidth * .8)
              table.dataset.tableAnchor = JSON.stringify({ x: Number(val(floating, 'tblpX')) / 15, y: Number(val(floating, 'tblpY')) / 15, bottom: Number(val(floating, 'bottomFromText')) / 15 });
          }
        }
        p.style.cssText = format.paragraph(marker.paragraph);
        const properties = child(marker.paragraph, 'pPr');
        if (val(child(properties, 'pStyle')) === TABLE_GAP_STYLE &&
            Array.from(marker.paragraph.children).every(c => c.localName === 'pPr' || c.localName.startsWith('bookmark'))) {
          const spacing = child(properties, 'spacing');
          if (val(spacing, 'lineRule') === 'exact') tableGaps.set(p, Math.max(0, Number(val(spacing, 'line')) / 15));
        }
        p.setAttribute('data-word-line-rule', format.lineRule(marker.paragraph));
        const tabs = format.tabs(marker.paragraph);
        if (tabs !== '[]') p.setAttribute('data-tab-stops', tabs);
        if (cell && marker.cell && !all(marker.cell, 't').some(t => t.textContent) && !all(marker.cell, 'tbl').length) {
          for (const property of ['font-family', 'font-size', 'color']) cell.style.setProperty(property, p.style.getPropertyValue(property));
        }
        const breakBefore = child(child(marker.paragraph, 'pPr'), 'pageBreakBefore');
        if (breakBefore && !['0', 'false', 'off'].includes(val(breakBefore))) {
          const pageBreak = document.createElement('div'); pageBreak.setAttribute('data-page-break', 'true'); p.before(pageBreak);
        }
        if (marker.nextPage) {
          const sectionBreak = document.createElement('div');
          sectionBreak.setAttribute('data-section-page', JSON.stringify(marker.nextPage));
          p.after(sectionBreak);
        }
      }
      // Word's minimum row height reserves the cell's top/bottom margins in
      // addition to the requested text height; HTML tr height includes them.
      // Wait until every cell has its formatting before measuring that reserve.
      for (const [table, sourceTable] of sizedTables) {
        const rows = all(sourceTable, 'tr').filter(row => ancestor(row, 'tbl') === sourceTable);
        if (rows.length !== table.rows.length) continue;
        rows.forEach((source, index) => {
          const split = child(child(source, 'trPr'), 'cantSplit');
          if (split) table.rows[index].dataset.rowSplit = ['0', 'false', 'off'].includes(val(split)) ? 'allow' : 'avoid';
          const property = child(child(source, 'trPr'), 'trHeight'), height = Number(val(property));
          if (height <= 0 || val(property, 'hRule') === 'auto') return;
          const row = table.rows[index], exact = val(property, 'hRule') === 'exact';
          const padding = exact ? 0 : Math.max(0, ...[...row.cells].map(cell =>
            (parseFloat(cell.style.paddingTop) || 0) + (parseFloat(cell.style.paddingBottom) || 0)));
          const border = exact ? 0 : Math.max(0, ...Array.from(source.children).filter(c => c.localName === 'tc').map(c => {
            const probe = document.createElement('td'); format.cell(probe, c);
            const merge = child(child(c, 'tcPr'), 'vMerge');
            const sides = merge && val(merge) !== 'restart' ? ['Top', 'Bottom'] : merge ? [] : ['Bottom'];
            return Math.max(0, ...sides.map(side => {
              const width = probe.style.getPropertyValue(`border-${side.toLowerCase()}-width`);
              return (parseFloat(width) || 0) * (width.endsWith('pt') ? 4 / 3 : 1);
            }));
          }));
          row.style.height = `${height / 15 + padding + border}px`;
          row.dataset.wordRowRule = exact ? 'exact' : 'atLeast';
        });
      }
      // An explicit cantSplit=false across the whole merge preserves its
      // permission to continue. Missing/default policy is not new evidence.
      for (const table of sizedTables.keys()) for (const [index, row] of [...table.rows].entries()) for (const cell of [...row.cells]) {
        if (cell.rowSpan > 1 && [...table.rows].slice(index, index + cell.rowSpan).every(r => r.dataset.rowSplit === 'allow' && r.dataset.wordRowRule !== 'exact')) cell.dataset.cellSplit = 'allow';
      }
      // Resolve preferred widths in their containing text area. Process outer
      // tables first so nested percentages use the actual padded parent cell.
      for (const table of template.content.querySelectorAll<HTMLTableElement>('table')) {
        const source = sizedTables.get(table); if (!source) continue;
        const section = all(body, 'sectPr').findIndex(s => !!(source.compareDocumentPosition(s) & Node.DOCUMENT_POSITION_FOLLOWING));
        const page = pages[section < 0 ? pages.length - 1 : section];
        const margins = page && pageMargins(page);
        const parent = table.parentElement?.closest<HTMLTableCellElement>('td,th');
        const parentWidth = parent?.getAttribute('colwidth')?.split(',').reduce((n, s) => n + Number(s), 0);
        const available = Math.max(1, parentWidth ? parentWidth - (parseFloat(parent!.style.paddingLeft) || 0) - (parseFloat(parent!.style.paddingRight) || 0) - 2
          : page && margins ? (paperDimensions(page)[0] - margins.left - margins.right) * 96 / 25.4 - Math.max(0, parseFloat(table.style.marginLeft) || 0) - 1 : 624);
        const grid = Array.from(child(source, 'tblGrid')?.children ?? []).map(c => Number(val(c, 'w')) / 15);
        const total = grid.reduce((n, w) => n + w, 0); if (!total || !grid.every(n => n > 0)) continue;
        const preferred = child(child(source, 'tblPr'), 'tblW');
        const borderPixels = (value: string) => (parseFloat(value) || 0) * (value.endsWith('pt') ? 4 / 3 : 1);
        const cells = [...table.querySelectorAll<HTMLTableCellElement>('td,th')].filter(cell => cell.closest('table') === table);
        // Collapsed outer strokes extend beyond the sum of colgroup widths.
        const borderReserve = (Math.max(0, ...cells.map(cell => borderPixels(cell.style.borderLeftWidth))) + Math.max(0, ...cells.map(cell => borderPixels(cell.style.borderRightWidth)))) / 2;
        const count = page?.columns?.count || 1;
        const columnWidth = parent ? available : (available - (count - 1) * (page?.columns?.space || 0)) / count;
        const requested = val(preferred, 'type') === 'pct' && Number(val(preferred, 'w')) > 0 ? columnWidth * Number(val(preferred, 'w')) / 5000 : total;
        const target = Math.min(columnWidth - borderReserve, requested);
        if (Math.abs(target - total) < 1) continue;
        for (const cell of table.querySelectorAll<HTMLTableCellElement>('td[colwidth],th[colwidth]')) {
          if (cell.closest('table') !== table) continue;
          const widths = cell.getAttribute('colwidth')!.split(',').map(Number);
          cell.setAttribute('colwidth', widths.map(w => Math.max(1, Math.floor(w * target / total))).join(','));
        }
      }
      const active = new Set<string>();
      let activeStyle = '';
      let eastAsianStyle = '';
      const walker = document.createTreeWalker(template.content, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT);
      const items: Node[] = []; while (walker.nextNode()) items.push(walker.currentNode);
      for (const item of items) {
        if (item instanceof HTMLElement && markers.has(item.getAttribute('data-ws-marker')!)) {
          const marker = markers.get(item.getAttribute('data-ws-marker')!)!;
          if (marker.kind === 'pageBreak' || marker.kind === 'columnBreak') continue;
          if (marker.kind === 'overlap') {
            const span = document.createElement('span'); span.dataset.wordOverlap = JSON.stringify(marker.parts);
            span.style.cssText = marker.style || ''; span.textContent = marker.parts!.join('');
            item.replaceWith(span); continue;
          }
          if (marker.kind === 'pageField') {
            const span = document.createElement('span'); span.dataset.pageNumber = marker.id; span.textContent = '1'; item.replaceWith(span); continue;
          }
          if (marker.kind === 'symbol') { item.replaceWith(symbolElement(marker.id!, Number(marker.parts![0]))); continue; }
          if (marker.kind === 'tab') {
            const tab = document.createElement('span'); tab.setAttribute('data-docx-tab', 'true'); tab.textContent = '\t';
            if (activeStyle) tab.style.cssText = activeStyle;
            item.replaceWith(tab); continue;
          }
          if (marker.kind === 'start') active.add(marker.id!);
          if (marker.kind === 'end') active.delete(marker.id!);
          if (marker.kind === 'runStart') { activeStyle = marker.style || ''; eastAsianStyle = marker.eastAsianStyle || activeStyle; }
          if (marker.kind === 'runEnd') { activeStyle = ''; eastAsianStyle = ''; }
          item.remove();
        } else if (item.nodeType === Node.TEXT_NODE && (active.size || activeStyle) && item.textContent) {
          const fragment = document.createDocumentFragment();
          const segments = activeStyle !== eastAsianStyle
            ? item.textContent.match(/[\u1100-\u11ff\u2e80-\ua4cf\uac00-\ud7af\uf900-\ufaff\uff00-\uffef]+|[^\u1100-\u11ff\u2e80-\ua4cf\uac00-\ud7af\uf900-\ufaff\uff00-\uffef]+/gu) ?? [] : [item.textContent];
          for (const text of segments) {
            const span = document.createElement('span');
            if (active.size) span.setAttribute('data-comment-ids', [...active].join(' '));
            span.style.cssText = /^[\u1100-\u11ff\u2e80-\ua4cf\uac00-\ud7af\uf900-\ufaff\uff00-\uffef]/u.test(text) ? eastAsianStyle : activeStyle;
            const font = symbolFont(span.style.cssText);
            if (font) for (const character of text) {
              const code = character.codePointAt(0)!;
              if (code <= 255 || (code >= 0xf000 && code <= 0xf0ff)) span.append(symbolElement(font, code));
              else span.append(document.createTextNode(character));
            }
            else span.textContent = text;
            fragment.append(span);
          }
          item.parentNode!.replaceChild(fragment, item);
          for (const id of active) { const comment = comments.find(c => c.id === id); if (comment) comment.quote += item.textContent; }
        }
      }
      for (const anchor of template.content.querySelectorAll('[data-ws-marker]')) {
        const kind = markers.get(anchor.getAttribute('data-ws-marker')!)?.kind;
        if (kind !== 'pageBreak' && kind !== 'columnBreak') continue;
        const p = anchor.closest('p,h1,h2,h3,h4,h5,h6');
        if (!p) { anchor.remove(); continue; }
        const range = document.createRange(); range.selectNodeContents(p); range.setStartAfter(anchor);
        const after = p.cloneNode(false); after.appendChild(range.extractContents()); anchor.remove();
        const pageBreak = document.createElement('div'); pageBreak.setAttribute(kind === 'columnBreak' ? 'data-column-break' : 'data-page-break', 'true');
        p.after(pageBreak);
        if (after.textContent || (after as Element).querySelector('img,br')) pageBreak.after(after);
        if (!p.textContent && !p.querySelector('img,br')) p.remove();
      }
      // Restore our structural table separators to margins. Leaving them as
      // body paragraphs can create an otherwise empty page after a full table.
      // An ordinary blank paragraph, or a separator edited to contain content,
      // remains an ordinary paragraph.
      for (const [p, gap] of tableGaps) {
        const table = p.previousElementSibling as HTMLElement | null;
        if (table?.tagName !== 'TABLE' || p.textContent || p.children.length) continue;
        if (p.nextElementSibling?.tagName !== 'TABLE' && !(p.parentElement?.matches('td,th') && !p.nextElementSibling)) continue;
        table.style.marginBottom = `${gap <= 1 / 15 ? 0 : gap}px`;
        p.remove();
      }
      // An inline table's following-paragraph spacing represents the gap
      // after the table. Keep it with that table when the next paragraph
      // flows onto a new page. Explicit breaks above have already inserted
      // a separate node, so their intentional top spacing stays untouched.
      for (const table of template.content.querySelectorAll<HTMLTableElement>('table')) {
        const next = table.nextElementSibling as HTMLElement | null;
        if (!next?.matches('p,h1,h2,h3,h4,h5,h6')) continue;
        const gap = parseFloat(next.style.marginTop);
        if (!(gap > 0)) continue;
        table.style.marginBottom = `${Math.max(parseFloat(table.style.marginBottom) || 0, gap)}px`;
        next.style.marginTop = '0px';
      }
      return template.innerHTML;
    },
  };
}

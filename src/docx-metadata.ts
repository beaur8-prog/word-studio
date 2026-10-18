import type JSZip from 'jszip';
import type { Comment, PageSettings } from './types';
import { W, all, val, child, ancestor, docxFormat, decorateCell } from './docx-format';
const parse = (text: string) => new DOMParser().parseFromString(text, 'application/xml');

/** Temporary bookmarks let Mammoth preserve exact positions, including repeated text. */
export async function prepareDocxMetadata(zip: JSZip, body: Document, pages: PageSettings[]) {
  const format = await docxFormat(zip);
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
      resolved: Boolean(paraId && resolved.has(paraId)) });
  }
  let sequence = 900000;
  const markers = new Map<string, { kind: string; id?: string; paragraph?: Element; cell?: Element; style?: string; eastAsianStyle?: string; nextPage?: PageSettings; width?: number; height?: number }>();
  const bookmark = (name: string) => {
    const fragment = body.createDocumentFragment();
    const id = String(++sequence);
    const start = body.createElementNS(W, 'w:bookmarkStart');
    start.setAttributeNS(W, 'w:id', id); start.setAttributeNS(W, 'w:name', name);
    const end = body.createElementNS(W, 'w:bookmarkEnd'); end.setAttributeNS(W, 'w:id', id);
    fragment.append(start, end);
    return fragment;
  };
  for (const kind of ['commentRangeStart', 'commentRangeEnd']) for (const el of all(body, kind)) {
    const id = idMap.get(val(el, 'id'));
    if (!id) continue;
    const name = `ws_marker_${sequence}`;
    markers.set(name, { kind: kind === 'commentRangeStart' ? 'start' : 'end', id });
    el.replaceWith(bookmark(name));
  }
  for (const drawing of all(body, 'drawing')) {
    const extent = drawing.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', 'extent')[0];
    const width = Number(extent?.getAttribute('cx')) / 9525, height = Number(extent?.getAttribute('cy')) / 9525;
    const name = `ws_marker_${sequence}`;
    markers.set(name, { kind: 'image', ...(width > 0 && height > 0 && width <= 10000 && height <= 10000 ? { width, height } : {}) });
    drawing.before(bookmark(name));
  }
  for (const tab of all(body, 'tab').filter(el => el.parentElement?.localName === 'r')) {
    const name = `ws_marker_${sequence}`; markers.set(name, { kind: 'tab' }); tab.replaceWith(bookmark(name));
  }
  for (const run of all(body, 'r')) {
    const style = format.run(run);
    if (!style) continue;
    const start = `ws_marker_${sequence}`;
    markers.set(start, { kind: 'runStart', style, eastAsianStyle: format.run(run, 'eastAsia') }); run.before(bookmark(start));
    const end = `ws_marker_${sequence}`;
    markers.set(end, { kind: 'runEnd' }); run.after(bookmark(end));
  }
  for (const br of all(body, 'br').filter(el => val(el, 'type') === 'page')) {
    const name = `ws_marker_${sequence}`;
    markers.set(name, { kind: 'pageBreak' }); br.replaceWith(bookmark(name));
  }
  let section = 0;
  for (const paragraph of all(body, 'p')) {
    const name = `ws_marker_${sequence}`;
    const isEnd = all(paragraph, 'sectPr').length > 0;
    const parent = paragraph.parentElement;
    const cell = parent?.localName === 'tc' && all(parent, 'p')[0] === paragraph ? parent : undefined;
    markers.set(name, { kind: 'paragraph', paragraph, cell,
      ...(isEnd ? { nextPage: pages[++section] } : {}) });
    const properties = Array.from(paragraph.children).find(e => e.localName === 'pPr');
    paragraph.insertBefore(bookmark(name), properties ? properties.nextSibling : paragraph.firstChild);
  }
  zip.file('word/document.xml', new XMLSerializer().serializeToString(body));
  return {
    buffer: await zip.generateAsync({ type: 'arraybuffer' }), comments,
    decorate(html: string) {
      const template = document.createElement('template'); template.innerHTML = html;
      const sizedTables = new Map<HTMLTableElement, Element>();
      let pendingImage: (typeof markers extends Map<string, infer T> ? T : never) | undefined;
      for (const el of template.content.querySelectorAll('a[id],img')) {
        if (el instanceof HTMLAnchorElement) {
          const marker = markers.get(el.id);
          if (marker?.kind === 'image') pendingImage = marker;
        } else if (pendingImage) {
          el.setAttribute('data-inline-image', 'true');
          if (pendingImage.width && pendingImage.height) {
            el.setAttribute('width', String(pendingImage.width)); el.setAttribute('height', String(pendingImage.height));
          }
          pendingImage = undefined;
        }
      }
      // Attach paragraph properties before deleting marker anchors.
      for (const anchor of template.content.querySelectorAll('a[id]')) {
        const marker = markers.get(anchor.id);
        if (marker?.kind !== 'paragraph') continue;
        const p = anchor.closest<HTMLElement>('p,h1,h2,h3,h4,h5,h6,li');
        if (!p || !marker.paragraph) continue;
        const cell = p.closest<HTMLElement>('td,th');
        if (cell && marker.cell) {
          decorateCell(cell, marker.cell);
          const sourceTable = ancestor(marker.cell, 'tbl'), table = cell.closest('table');
          if (sourceTable && table && !sizedTables.has(table)) {
            sizedTables.set(table, sourceTable);
            table.style.marginLeft = `${Number(val(child(child(sourceTable, 'tblPr'), 'tblInd'), 'w')) / 15}px`;
            table.style.marginBottom = '0px';
          }
        }
        p.style.cssText = format.paragraph(marker.paragraph);
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
          const property = child(child(source, 'trPr'), 'trHeight'), height = Number(val(property));
          if (height <= 0 || val(property, 'hRule') === 'auto') return;
          const row = table.rows[index], exact = val(property, 'hRule') === 'exact';
          const padding = exact ? 0 : Math.max(0, ...[...row.cells].map(cell =>
            (parseFloat(cell.style.paddingTop) || 0) + (parseFloat(cell.style.paddingBottom) || 0)));
          row.style.height = `${height / 15 + padding}px`;
          row.dataset.wordRowRule = exact ? 'exact' : 'atLeast';
        });
      }
      const active = new Set<string>();
      let activeStyle = '';
      let eastAsianStyle = '';
      const walker = document.createTreeWalker(template.content, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT);
      const items: Node[] = []; while (walker.nextNode()) items.push(walker.currentNode);
      for (const item of items) {
        if (item instanceof HTMLAnchorElement && markers.has(item.id)) {
          const marker = markers.get(item.id)!;
          if (marker.kind === 'pageBreak') continue;
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
            span.textContent = text; fragment.append(span);
          }
          item.parentNode!.replaceChild(fragment, item);
          for (const id of active) { const comment = comments.find(c => c.id === id); if (comment) comment.quote += item.textContent; }
        }
      }
      for (const anchor of template.content.querySelectorAll('a[id]')) {
        if (markers.get(anchor.id)?.kind !== 'pageBreak') continue;
        const p = anchor.closest('p,h1,h2,h3,h4,h5,h6');
        if (!p) { anchor.remove(); continue; }
        const range = document.createRange(); range.selectNodeContents(p); range.setStartAfter(anchor);
        const after = p.cloneNode(false); after.appendChild(range.extractContents()); anchor.remove();
        const pageBreak = document.createElement('div'); pageBreak.setAttribute('data-page-break', 'true');
        p.after(pageBreak);
        if (after.textContent || (after as Element).querySelector('img,br')) pageBreak.after(after);
        if (!p.textContent && !p.querySelector('img,br')) p.remove();
      }
      return template.innerHTML;
    },
  };
}

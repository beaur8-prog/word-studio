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
  const markers = new Map<string, { kind: string; id?: string; paragraph?: Element; cell?: Element; style?: string; nextPage?: PageSettings }>();
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
  for (const run of all(body, 'r')) {
    const style = format.run(run);
    if (!style) continue;
    const start = `ws_marker_${sequence}`;
    markers.set(start, { kind: 'runStart', style }); run.before(bookmark(start));
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
      const sizedTables = new Set<HTMLTableElement>();
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
            sizedTables.add(table);
            table.style.marginLeft = `${Number(val(child(child(sourceTable, 'tblPr'), 'tblInd'), 'w')) / 15}px`;
            table.style.marginBottom = '0px';
            const rows = all(sourceTable, 'tr').filter(row => ancestor(row, 'tbl') === sourceTable);
            // Vertical merges can leave an HTML row without any paragraph
            // anchors. Table row order still matches the source, including these rows.
            if (rows.length === table.rows.length) rows.forEach((source, index) => {
              const property = child(child(source, 'trPr'), 'trHeight'), height = Number(val(property));
              if (height > 0 && val(property, 'hRule') !== 'auto') table.rows[index].style.height = `${height / 15}px`;
            });
          }
        }
        p.style.cssText = format.paragraph(marker.paragraph);
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
      const active = new Set<string>();
      let activeStyle = '';
      const walker = document.createTreeWalker(template.content, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT);
      const items: Node[] = []; while (walker.nextNode()) items.push(walker.currentNode);
      for (const item of items) {
        if (item instanceof HTMLAnchorElement && markers.has(item.id)) {
          const marker = markers.get(item.id)!;
          if (marker.kind === 'pageBreak') continue;
          if (marker.kind === 'start') active.add(marker.id!);
          if (marker.kind === 'end') active.delete(marker.id!);
          if (marker.kind === 'runStart') activeStyle = marker.style || '';
          if (marker.kind === 'runEnd') activeStyle = '';
          item.remove();
        } else if (item.nodeType === Node.TEXT_NODE && (active.size || activeStyle) && item.textContent) {
          const span = document.createElement('span');
          if (active.size) span.setAttribute('data-comment-ids', [...active].join(' '));
          if (activeStyle) span.style.cssText = activeStyle;
          item.parentNode!.replaceChild(span, item); span.append(item);
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

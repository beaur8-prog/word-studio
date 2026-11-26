import { activeHwpLatinFont, hwpLatinFontAttribute, hwpLatinFontNamespace } from './hwp-latin-font';
import { hwpNativeGridAttribute, parseHwpNativeGrid } from './hwp-native-grid';
import { hwpRelatedTableLayoutAttribute, parseHwpRelatedTableLayout } from './hwp-related-table-layout';
import { hwpNestedTableLayoutAttribute, parseHwpNestedTableLayout } from './hwp-nested-table-layout';
import { hwpCenteredCellLabelsAttribute, hwpCenteredCellLabelsNamespace, parseHwpCenteredCellLabels } from './hwp-centered-cell-labels';
import { parseHwpTableContinuation } from './hwp-table-continuation';
import { hwpBoxedDigitAttribute, hwpBoxedDigitBorder, hwpBoxedDigitNamespace, parseHwpBoxedDigit } from './hwp-boxed-digit';
import { paragraphBorderBoxAttribute } from './paragraph-border-box';
import { flattenDocxCheckboxes, restoreDocxCheckboxInputs } from './docx-checkboxes';
import { encodeDocxXmlParts } from './docx-package-text';
import { cellTablePosition } from './cell-table-position';
import { tablePixelGrid } from './table-grid-widths';
import { rowSplitPolicy } from './row-split-policy';
import { hwpGridRowHeightNamespace } from './docx-table-cell';
import { hwpRowReserve } from './hwp-row-reserve';
import { hwpIntactRowAttribute, parseHwpIntactRow } from './hwp-intact-row';
import { rowHeight } from './table-row-resize';
import {parseDrawingGroup} from './drawing-group';
import type JSZip from 'jszip';
import { pageMargins, paperDimensions, type Comment, type PageSettings } from './types';
import { W, all, val, child, ancestor, docxFormat, tableCellGridOffset } from './docx-format';
import { overlapInstruction } from './word-overlap';
import { readWordReflection, wordReflectionAttribute, wordParagraphReflectionAttribute } from './word-reflection';
import { symbolFont, symbolElement } from './document-symbols';
import { normalizeVmlImages } from './docx-images';
import { drawingFloating } from './image-floating';
import { encodeHwpCellRasterFill, hwpCellRasterFillAttribute, hwpCellRasterSource } from './hwp-cell-raster-fill';
import { tableAnchor, tableWrapFits, paragraphTableNamespace } from './table-anchor';
import { hwpInlineTableTop, hwpInlineTableTopAttribute, hwpKeptTablePage, hwpTableTopGap } from './table-anchor-data';
import { cssPixels, encodeHwpCellLine, hwpCellLineNamespace, parseHwpCellLine } from './hwp-cell-line';
import { hwpSourceSpacesAttribute, hwpSourceSpacesNamespace, hwpSourceWrapAttribute, parseHwpSourceWrap } from './hwp-text-advances';
import { encodeHwpBodyLine, hwpBodyLineNamespace, parseHwpBodyLine } from './hwp-body-line';
const parse = (text: string) => new DOMParser().parseFromString(text, 'application/xml');
const wordChild = (parent: Element | undefined, name: string) =>
  Array.from(parent?.children ?? []).find(element => element.namespaceURI === W && element.localName === name);
function nestedFloatingPosition(table: Element): string | undefined {
  const cell = table.parentElement;
  if (cell?.localName !== 'tc' || ['txbxContent', 'footnote', 'endnote', 'comment', 'hdr', 'ftr'].some(name => ancestor(table, name))) return;
  const children = Array.from(cell.children);
  const pair = children.filter(child => child.localName === 'tbl');
  if (pair.length !== 2 || children.some(child => child.localName !== 'tbl' && child.localName !== 'tcPr' &&
    (child.localName !== 'p' || child.textContent?.trim() || ['tab', 'br', 'drawing', 'object', 'pict', 'fldChar'].some(name => all(child, name).length)))) return;
  const positions = pair.map(table => {
    const floating = child(child(table, 'tblPr'), 'tblpPr');
    if (!floating || val(floating, 'horzAnchor') !== 'margin' || val(floating, 'vertAnchor') !== 'margin' ||
      ['tblpXSpec', 'tblpYSpec', 'leftFromText', 'rightFromText', 'topFromText', 'bottomFromText']
        .some(name => val(floating, name) && val(floating, name) !== '0')) return;
    const value = `${val(floating, 'tblpX')} ${val(floating, 'tblpY')}`;
    return cellTablePosition(value) ? value : undefined;
  });
  return positions.every(Boolean) ? positions[pair.indexOf(table)] : undefined;
}

function pairedTableMargins(body: Document, positions: WeakMap<Element, string>) {
  const affected = new Set<Element>();
  for (const table of all(body, 'tbl')) {
    const position = nestedFloatingPosition(table);
    if (!position) continue;
    positions.set(table, position);
    affected.add(table);
    const host = ancestor(table.parentElement!, 'tbl');
    if (host) affected.add(host);
  }
  // For the cell-relative pair and its host, treat the explicit logical
  // table margins like their legacy aliases before applying inherited styles.
  for (const table of affected) {
    const margins = wordChild(wordChild(table, 'tblPr'), 'tblCellMar');
    for (const [logical, physical] of [['start', 'left'], ['end', 'right']]) {
      const edge = wordChild(margins, logical);
      if (!edge || wordChild(margins, physical)) continue;
      const alias = body.createElementNS(W, `w:${physical}`);
      for (const attribute of edge.attributes) alias.setAttributeNS(attribute.namespaceURI, attribute.name, attribute.value);
      edge.replaceWith(alias);
    }
  }
}
const reflectionBlock = 'p,h1,h2,h3,h4,h5,h6,li';
// A standard paragraph style identifies the otherwise invisible separator
// required by Word between adjacent tables and after a table inside a cell.
export const TABLE_GAP_STYLE = 'WordStudioTableGap';
// A zero-height ordinary paragraph gives a paragraph-relative floating table
// a stable standard OOXML anchor, independent of the following body text.
export const TABLE_ANCHOR_STYLE = 'WordStudioTableAnchor';
export const TABLE_OFFSET_STYLE = 'WordStudioTableOffset';
export const SECTION_BOUNDARY_STYLE = 'WordStudioSectionBoundary';
export const TABLE_END_STYLE = 'WordStudioTableEnd';

/** Temporary bookmarks let Mammoth preserve exact positions, including repeated text. */
export async function prepareDocxMetadata(zip: JSZip, body: Document, pages: PageSettings[],
  { preserveReflection = true, connectedBodyBorders = false, textBoxStory = false } = {}) {
  flattenDocxCheckboxes(body);
  const format = await docxFormat(zip);
  // Word requires a final paragraph after a body table. Our small structural
  // carrier is not an authored blank line. Keep it if an external editor has
  // added content, a bookmark, section settings or changed its geometry.
  for (const p of all(body, 'p')) {
    const props = child(p, 'pPr'), run = child(props, 'rPr');
    if (val(child(props, 'pStyle')) !== TABLE_END_STYLE || p.parentElement?.localName !== 'body' ||
      p.previousElementSibling?.localName !== 'tbl' || p.nextElementSibling && p.nextElementSibling.localName !== 'sectPr' ||
      !Array.from(p.children).every(c => c.localName === 'pPr' || c.localName === 'r' && Array.from(c.children).every(r => r.localName === 'rPr' && !r.children.length)) ||
      !Array.from(props?.children || []).every(c => ['pStyle','spacing','rPr','widowControl'].includes(c.localName)) ||
      !Array.from(run?.children || []).every(c => ['sz','szCs'].includes(c.localName))) continue;
    const style = document.createElement('p').style; style.cssText = format.paragraph(p);
    if (format.lineRule(p) === 'exact' && Math.abs((cssPixels(style.fontSize) ?? 0) - 4 / 3) < .001 &&
      Math.abs((cssPixels(style.lineHeight) ?? 0) - 4 / 3) < .001 &&
      Math.abs(parseFloat(style.marginTop) || 0) < .001 && Math.abs(parseFloat(style.marginBottom) || 0) < .001) p.remove();
  }

  // An empty legacy floating-frame anchor has no inline line box. Treating
  // it as an ordinary paragraph after its table creates a spurious page.
  for (const p of all(body, 'p')) if (child(child(p, 'pPr'), 'framePr') && !ancestor(p, 'tc') &&
    Array.from(p.children).every(c => c.localName === 'pPr')) p.remove();
  const cellTablePositions = new WeakMap<Element, string>();
  pairedTableMargins(body, cellTablePositions);
  const connectedBorders = connectedBodyBorders ? format.connectedBodyBorders(body) : undefined;
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
    comments.push({ id, quote: '', text: all(el, 'p').map(p => Array.from(p.getElementsByTagNameNS(W, '*')).map(n =>
      n.localName === 't' ? n.textContent : n.localName === 'tab' ? '\t' : n.localName === 'br' ? '\n' : '').join('')).join('\n'),
      createdAt: Number.isFinite(Date.parse(date)) ? new Date(date).toISOString() : new Date().toISOString(),
      resolved: Boolean(paraId && resolved.has(paraId)), author: val(el, 'author'), initials: val(el, 'initials') });
  }
  let sequence = 900000;
  const numbering = new WeakMap<Element, ReturnType<typeof format.numbering>>();
  for (const p of all(body, 'p')) numbering.set(p, format.numbering(p));
  const markers = new Map<string, { kind: string; id?: string; parts?: string[]; paragraph?: Element; cell?: Element; style?: string; eastAsianStyle?: string; page?: PageSettings; nextPage?: PageSettings; boundaryCarrier?: boolean; width?: number; height?: number; alt?: string; floating?: string; reflection?: string; reflectionEnd?: string; latinFont?: string }>();
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
  for (const sym of all(body, 'sym')) {
    const name = `ws_marker_${sequence}`;
    markers.set(name, { kind: 'symbol', id: val(sym, 'font'), parts: [String(parseInt(val(sym, 'char'), 16))] });
    sym.replaceWith(bookmark(name));
  }
  for (const run of all(body, 'r')) {
    const encoded = run.getAttributeNS(hwpBoxedDigitNamespace, 'box');
    const box = parseHwpBoxedDigit(encoded);
    if (box && run.namespaceURI === W) {
      const props = wordChild(run, 'rPr'), border = wordChild(props, 'bdr');
      const expected = hwpBoxedDigitBorder(box.size);
      const text = wordChild(run, 't');
      if (run.children.length === 2 && text?.textContent === box.digit &&
          Number(val(wordChild(props, 'sz'))) === Math.round(box.size * .7) &&
          val(border) === expected.style && Number(val(border, 'sz')) === expected.size &&
          Number(val(border, 'space')) === expected.space) {
        const name = 'ws_marker_' + sequence;
        markers.set(name, { kind: 'boxedDigit', id: encoded!, parts: [box.digit], style: format.run(run) });
        run.replaceWith(bookmark(name));
        continue;
      }
    }
    const style = format.run(run);
    const literal = preserveReflection && run.namespaceURI === W && wordChild(run, 't') &&
      Array.from(run.children).every(element => element.namespaceURI === W && ['rPr', 't', 'tab'].includes(element.localName));
    const reflection = literal ? readWordReflection(wordChild(run, 'rPr')) : undefined;
    if (!style && !reflection) continue;
    const latinFont = run.namespaceURI === W ? run.getAttributeNS(hwpLatinFontNamespace, 'face') : null;
    const start = `ws_marker_${sequence}`;
    markers.set(start, { kind: 'runStart', style, eastAsianStyle: format.run(run, 'eastAsia'), latinFont: latinFont && latinFont.length <= 2000 ? latinFont : undefined }); run.before(bookmark(start));
    const end = `ws_marker_${sequence}`;
    markers.set(end, { kind: 'runEnd' }); run.after(bookmark(end));
    if (reflection) Object.assign(markers.get(start)!, { reflection, reflectionEnd: end });
  }
  for (const tab of all(body, 'tab').filter(el => el.parentElement?.localName === 'r')) {
    const name = `ws_marker_${sequence}`; markers.set(name, { kind: 'tab' }); tab.replaceWith(bookmark(name));
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
    const boundaryProperties = child(paragraph, 'pPr'), boundarySpacing = child(boundaryProperties, 'spacing');
    const boundaryCarrier = isEnd && val(child(boundaryProperties, 'pStyle')) === SECTION_BOUNDARY_STYLE &&
      Array.from(paragraph.children).every(c => c.localName === 'pPr') &&
      Array.from(boundaryProperties?.children || []).every(c => ['pStyle', 'spacing', 'sectPr'].includes(c.localName)) &&
      val(boundarySpacing, 'lineRule') === 'exact' && Number(val(boundarySpacing, 'line')) === 1 &&
      Number(val(boundarySpacing, 'before')) === 0 && Number(val(boundarySpacing, 'after')) === 0 &&
      Array.from(boundarySpacing?.attributes || []).every(a => ['before', 'after', 'line', 'lineRule'].includes(a.localName));
    markers.set(name, { kind: 'paragraph', paragraph, cell, page: pages[section], boundaryCarrier,
      ...(isEnd ? { nextPage: pages[++section] } : {}) });
    const properties = Array.from(paragraph.children).find(e => e.localName === 'pPr');
    paragraph.insertBefore(bookmark(name), properties ? properties.nextSibling : paragraph.firstChild);
  }
  zip.file('word/document.xml', new XMLSerializer().serializeToString(body));
  await encodeDocxXmlParts(zip);
  return {
    buffer: await zip.generateAsync({ type: 'arraybuffer' }), comments, wordDefaultLanguage: format.wordDefaultLanguage,
    decorate(html: string) {
      // Anchor markers nested inside hyperlinks make the HTML parser close the
      // outer link. Convert markers before parsing, not after damage is done.
      const template = document.createElement('template');
      template.innerHTML = html.replace(/<a\s+id="(ws_marker_\d+)"\s*><\/a>/g, '<span data-ws-marker="$1"></span>')
        .replace(/<a\s+id="([^"]+)"\s*><\/a>/g, '<span data-docx-bookmark="$1"></span>');
      restoreDocxCheckboxInputs(template.content);
      const sizedTables = new Map<HTMLTableElement, Element>();
      const sizedCells = new WeakMap<HTMLElement, Element>();
      const tableGaps = new Map<HTMLElement, number>();
      const tableAnchorParagraphs = new Set<HTMLElement>();
      const tableOffsetParagraphs = new Map<HTMLElement, Element>();
      const verifiedLeadingTableOffsets = new WeakMap<Element, HTMLTableElement>();
      const inlineGroups=new Set<string>();
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
        const floating = marker.floating ? JSON.parse(marker.floating) : undefined;
        if (floating?.hwpCellRasterFill && marker.width && marker.height && marker.width < 10000 && marker.height < 10000) {
          const p = el.closest('p,h1,h2,h3,h4,h5,h6'), cell = p?.parentElement;
          const source = el.getAttribute('src');
          if (cell?.matches('td') && !cell.hasAttribute(hwpCellRasterFillAttribute) && hwpCellRasterSource(source)) {
            cell.setAttribute(hwpCellRasterFillAttribute, encodeHwpCellRasterFill(marker.width, marker.height, source, floating.hwpCellRasterBrightness));
            el.remove();
            continue;
          }
        }
        const group=floating ? parseDrawingGroup(JSON.stringify(floating.drawingGroup)) : undefined;
        if(group?.position==='inline' && !inlineGroups.has(group.id)){const extent=document.createElement('span');extent.dataset.drawingGroupAnchor=JSON.stringify(group);el.before(extent);inlineGroups.add(group.id);}
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
          if (list.suffix !== 'tab') item.setAttribute('data-list-suffix', list.suffix);
          if (list.font) item.setAttribute('data-list-font', list.font);
          item.setAttribute('data-list-start', String(list.start));
          const parentList = item.parentElement;
          if (parentList?.tagName === 'OL' && parentList.firstElementChild === item) parentList.setAttribute('start', String(list.value));
        }
        const cell = p.closest<HTMLElement>('td,th');
        if (cell && marker.cell) {
          format.cell(cell, marker.cell);
          const leadingLabels = marker.cell.getAttributeNS(hwpCenteredCellLabelsNamespace, 'leading');
          if (cell instanceof HTMLTableCellElement && cell.tagName === 'TD' && cell.style.verticalAlign === 'middle' &&
              cell.rowSpan === 1 && cell.colSpan === 1 &&
              parseHwpCenteredCellLabels(leadingLabels)) cell.setAttribute(hwpCenteredCellLabelsAttribute, leadingLabels!);
          const cellWidth = child(child(marker.cell, 'tcPr'), 'tcW');
          if (!cellWidth || val(cellWidth, 'type') === 'auto') cell.dataset.wordCellWidthAuto = 'true';
          sizedCells.set(cell, marker.cell);
          const sourceTable = ancestor(marker.cell, 'tbl'), table = cell.closest('table');
          if (sourceTable && table && !sizedTables.has(table)) {
            sizedTables.set(table, sourceTable);
            table.dataset.wordTableLayout = format.tableLayout(sourceTable);
            const preferredWidth = child(child(sourceTable, 'tblPr'), 'tblW');
            if (!preferredWidth || val(preferredWidth, 'type') === 'auto') table.dataset.wordTableWidthAuto = 'true';
            table.style.marginLeft = `${Number(val(child(child(sourceTable, 'tblPr'), 'tblInd'), 'w')) / 15}px`;
            const alignment = val(child(child(sourceTable, 'tblPr'), 'jc'), 'val');
            if (['center', 'right', 'end'].includes(alignment)) {
              table.style.marginLeft = 'auto'; table.style.marginRight = alignment === 'center' ? 'auto' : '0px';
            }
            table.style.marginBottom = '0px';
            const floating = child(child(sourceTable, 'tblPr'), 'tblpPr');
            const cellPosition = floating && table.closest('td') && cellTablePositions.get(sourceTable);
            if (cellPosition) table.dataset.cellTablePosition = cellPosition;
            // Page-relative tables spanning the text area can wrap only above
            // and below. Narrow/sided floats require a separate wrap layout.
            const grid = child(sourceTable, 'tblGrid');
            const gridWidth = (grid ? all(grid, 'gridCol') : []).reduce((sum, col) => sum + Number(val(col, 'w')), 0) / 15;
            const page = marker.page ?? pages[0], margins = pageMargins(page), bodyWidth = (paperDimensions(page)[0] - margins.left - margins.right) * 96 / 25.4;
            if (floating && val(floating, 'horzAnchor') === 'page' && val(floating, 'vertAnchor') === 'page' &&
              !val(floating, 'tblpXSpec') && !val(floating, 'tblpYSpec') && gridWidth >= bodyWidth * .8)
              table.dataset.tableAnchor = JSON.stringify({ x: Number(val(floating, 'tblpX')) / 15, y: Number(val(floating, 'tblpY')) / 15, bottom: Number(val(floating, 'bottomFromText')) / 15 });
            if (floating && !table.closest('td,th') && val(floating, 'horzAnchor') === 'margin' &&
              val(floating, 'vertAnchor') === 'page' && val(floating, 'tblpXSpec') === 'center' &&
              !val(floating, 'tblpYSpec') && gridWidth >= bodyWidth * .8) {
              const anchor = tableAnchor({ x: 0, y: Number(val(floating, 'tblpY')) / 15,
                bottom: Number(val(floating, 'bottomFromText')) / 15, horizontal: 'margin', horizontalAlign: 'center' });
              if (anchor) table.dataset.tableAnchor = JSON.stringify(anchor);
            }
            if (floating && val(floating, 'vertAnchor') === 'text' && ['page', 'margin', 'text'].includes(val(floating, 'horzAnchor')) &&
              !val(floating, 'tblpXSpec') && !val(floating, 'tblpYSpec') && gridWidth >= bodyWidth * .8 && (page.columns?.count || 1) === 1) {
              const anchor = tableAnchor({ x: Number(val(floating, 'tblpX')) / 15, y: Number(val(floating, 'tblpY')) / 15,
                bottom: Number(val(floating, 'bottomFromText')) / 15, reference: 'paragraph', ...(val(floating, 'horzAnchor') !== 'page' ? { horizontal: 'margin' } : {}) });
              if (anchor) table.dataset.tableAnchor = JSON.stringify(anchor);
            }
            // Support a left-side, paragraph-relative label with text on its
            // right. Read the standard geometry afresh: another Word editor
            // may retain our optional attributes while moving the table.
            if (floating && !table.closest('td,th') && val(floating, 'vertAnchor') === 'text' && ['margin', 'text'].includes(val(floating, 'horzAnchor')) &&
              !val(floating, 'tblpXSpec') && !val(floating, 'tblpYSpec') && Number(val(floating, 'tblpY')) === 0 &&
              gridWidth > 0 && (page.columns?.count || 1) === 1) {
              const x = Number(val(floating, 'tblpX')) / 15, right = Number(val(floating, 'rightFromText')) / 15;
              const anchor = tableAnchor({ x, y: 0, bottom: Number(val(floating, 'bottomFromText')) / 15,
                reference: 'paragraph', horizontal: 'margin', wrap: 'right', right });
              if (anchor && Number(val(floating, 'leftFromText')) / 15 >= x - .5 && tableWrapFits(anchor, gridWidth, bodyWidth))
                table.dataset.tableAnchor = JSON.stringify(anchor);
            }
            const nestedLayout = sourceTable.getAttributeNS(paragraphTableNamespace, 'nestedLayout');
            if (!table.closest('td,th') && parseHwpNestedTableLayout(nestedLayout))
              table.setAttribute(hwpNestedTableLayoutAttribute, nestedLayout!);
            const relatedLayout = sourceTable.getAttributeNS(paragraphTableNamespace, 'relatedLayout');
            if (!floating && !table.closest('td,th') && parseHwpRelatedTableLayout(relatedLayout))
              table.setAttribute(hwpRelatedTableLayoutAttribute, relatedLayout!);
            const keptPage = hwpKeptTablePage(sourceTable.getAttributeNS(paragraphTableNamespace, 'keptTablePage'));
            if (!floating && keptPage) table.dataset.hwpKeptTablePage = String(keptPage);
            const stored = tableAnchor(sourceTable.getAttributeNS(paragraphTableNamespace, 'anchor'));
            const nativeGrid = sourceTable.getAttributeNS(paragraphTableNamespace, 'nativeGrid');
            if (!stored && !floating && parseHwpNativeGrid(nativeGrid) && !table.closest('td,th'))
              table.setAttribute(hwpNativeGridAttribute, nativeGrid!);
            const inlineTop = hwpInlineTableTop(sourceTable.getAttributeNS(paragraphTableNamespace, 'inlineOuterTop'));
            if (!stored && !floating && inlineTop !== undefined && !table.closest('td,th'))
              table.setAttribute(hwpInlineTableTopAttribute, String(inlineTop));
            const topGap = hwpTableTopGap(sourceTable.getAttributeNS(paragraphTableNamespace, 'topGap'));
            const repeat = child(child(child(sourceTable, 'tr'), 'trPr'), 'tblHeader');
            const borders = sourceTable.getAttributeNS(paragraphTableNamespace, 'centeredBorders');
            if (!floating && !stored && borders === 'inline') table.dataset.hwpCenteredBorders = 'inline';
            if (!floating && repeat && !['0', 'false', 'off'].includes(val(repeat))) {
              const continuation = sourceTable.getAttributeNS(paragraphTableNamespace, 'continuation');
              if (parseHwpTableContinuation(continuation)) table.dataset.hwpTableContinuation = continuation!;
              if (topGap && !stored) table.dataset.hwpTableTopGap = String(topGap);
              if ((!stored || stored.reference === 'paragraph' && !stored.wrap) &&
                  (borders === '1' || borders === 'inline' || borders === 'continuation'))
                table.dataset.hwpCenteredBorders = borders === '1' ? 'true' : borders;
            }
            const unchangedOffset = (element: Element | null, pixels: number) => !pixels || !!element && element.localName === 'p' &&
              val(child(child(element, 'pPr'), 'pStyle')) === TABLE_OFFSET_STYLE &&
              Array.from(element.children).every(c => c.localName === 'pPr' || c.localName.startsWith('bookmark')) &&
              val(child(child(element, 'pPr'), 'spacing'), 'lineRule') === 'exact' &&
              Math.abs(Number(val(child(child(element, 'pPr'), 'spacing'), 'line')) / 15 - pixels) <= 1 / 15 + .001 &&
              // A newly carried gap must remain an ordinary paragraph when
              // an external editor adds separate formatting to its carrier.
              (!(Number(val(child(child(element, 'pPr'), 'spacing'), 'before')) > 0) || format.plainTableOffset(element));
            const indent = child(child(sourceTable, 'tblPr'), 'tblInd');
            // An external editor can change the visible inline geometry while
            // retaining ignorable attributes. Respect those edits, including
            // text entered in an offset paragraph, instead of restoring stale
            // positioning or removing the edited paragraph.
            if (stored?.reference === 'paragraph' && !stored.wrap && gridWidth >= bodyWidth * .8 && (page.columns?.count || 1) === 1 &&
              val(indent, 'type') === 'dxa' && Math.abs(Number(val(indent, 'w')) / 15 - stored.x + (stored.horizontal === 'margin' ? 0 : margins.left * 96 / 25.4)) <= 1 / 15 + .001 &&
              unchangedOffset(sourceTable.previousElementSibling, stored.y) && unchangedOffset(sourceTable.nextElementSibling, stored.bottom)) {
              table.dataset.tableAnchor = JSON.stringify(stored);
              const offset = sourceTable.previousElementSibling, previous = offset?.previousElementSibling;
              if (stored.y && sourceTable.parentElement?.localName === 'body' && offset && previous?.localName === 'tbl' &&
                !child(child(previous, 'tblPr'), 'tblpPr') && !previous.getAttributeNS(paragraphTableNamespace, 'anchor'))
                verifiedLeadingTableOffsets.set(offset, table);
            }
            const overlayPosition = sourceTable.getAttributeNS(paragraphTableNamespace, 'overlayPosition');
            if (stored?.overlay && floating && overlayPosition && overlayPosition.length < 500 &&
              val(floating, 'horzAnchor') === 'page' && val(floating, 'vertAnchor') === 'page' &&
              !val(floating, 'tblpXSpec') && !val(floating, 'tblpYSpec')) {
              try {
                const position = JSON.parse(overlayPosition);
                if ([position.x, position.y, position.width].every(Number.isFinite) && position.width > 0 &&
                  Math.abs(Number(val(floating, 'tblpX')) / 15 - position.x) <= 1 / 15 + .001 &&
                  Math.abs(Number(val(floating, 'tblpY')) / 15 - position.y) <= 1 / 15 + .001 &&
                  Math.abs(gridWidth - position.width) <= 1 / 15 + .001)
                  table.dataset.tableAnchor = JSON.stringify(stored);
              } catch { /* Changed or missing optional geometry keeps the standard table. */ }
            }
          }
        }
        p.style.cssText = format.paragraph(marker.paragraph);
        if (preserveReflection && marker.paragraph.namespaceURI === W) {
          const reflection = readWordReflection(wordChild(wordChild(marker.paragraph, 'pPr'), 'rPr'));
          if (reflection) p.setAttribute(wordParagraphReflectionAttribute, reflection);
        }
        p.dataset.wordParagraphStyle = format.paragraphStyle(marker.paragraph);
        p.dataset.wordParagraphStyleName = format.paragraphStyleName(marker.paragraph);
        const outline = format.outlineLevel(marker.paragraph);
        if (outline !== null) p.dataset.wordOutlineLevel = String(outline);
        const keeps = format.paragraphKeeps(marker.paragraph);
        if (keeps.next !== null) p.dataset.wordKeepNext = String(keeps.next);
        if (keeps.lines !== null) p.dataset.wordKeepLines = String(keeps.lines);
        const textAlignment = format.textAlignment(marker.paragraph);
        if (textAlignment) p.dataset.wordTextAlignment = textAlignment;
        const contextual = format.contextualSpacing(marker.paragraph);
        if (contextual !== null) p.dataset.wordContextualSpacing = String(contextual);
        const { bottom, box } = format.paragraphBorders(marker.paragraph);
        if (bottom) p.dataset.wordParagraphBottomBorder = JSON.stringify(bottom);
        const connectedBox = box ?? connectedBorders?.get(marker.paragraph);
        if (connectedBox) p.setAttribute(paragraphBorderBoxAttribute, JSON.stringify(connectedBox));
        if (/^H[1-6]$/.test(p.tagName)) p.dataset.wordHeadingMarks = format.headingMarks(marker.paragraph);
        const autoSpacing = format.autoSpacing(marker.paragraph);
        p.dataset.wordAutoSpaceDe = String(autoSpacing.de);
        p.dataset.wordAutoSpaceDn = String(autoSpacing.dn);
        const storedCellLine = marker.paragraph.getAttributeNS(hwpCellLineNamespace, 'cellLine');
        if (storedCellLine && storedCellLine.length < 500 && p.closest('td,th')) {
          try {
            const stored = JSON.parse(storedCellLine), line = parseHwpCellLine(stored.policy);
            const spacing = child(child(marker.paragraph, 'pPr'), 'spacing');
            // An external editor may preserve unknown attributes while changing
            // the standard line rule. Such explicit changes take precedence.
            const font = cssPixels(p.style.fontSize);
            const strut = stored.strutFont;
            const defaultFont = Number(val(child(child(child(marker.paragraph, 'pPr'), 'rPr'), 'sz')));
            const unchangedStrut = strut === undefined || strut === null || typeof strut === 'number' && strut > 0 && strut <= 2000 &&
              Number.isFinite(stored.defaultFont) && stored.defaultFont > 0 && defaultFont === stored.defaultFont;
            if (line && font !== null && Number.isFinite(stored.font) && Math.abs(font - stored.font) < .08 &&
              unchangedStrut && val(spacing, 'lineRule') === 'exact' && Number(val(spacing, 'line')) === stored.line) {
              p.dataset.hwpCellLine = encodeHwpCellLine(line);
              p.style.lineHeight = `${line.base}px`;
              // Run sizes remain explicit. Preserve the original paragraph
              // strut too, so a large space does not enlarge unmarked text or
              // move the small runs after saving. External default edits win.
              if (typeof strut === 'number' && strut > 0) p.style.fontSize = `${strut}px`;
            }
          } catch { /* Optional foreign metadata is never required to open DOCX. */ }
        }
        const properties = child(marker.paragraph, 'pPr');
        if (val(child(properties, 'pStyle')) === TABLE_GAP_STYLE &&
            Array.from(marker.paragraph.children).every(c => c.localName === 'pPr' || c.localName.startsWith('bookmark'))) {
          const spacing = child(properties, 'spacing');
          if (val(spacing, 'lineRule') === 'exact') {
            // Older exports put the whole gap in the line height. New body
            // separators use before-spacing so it does not follow a table
            // onto a new page. Both represent the same editable CSS margin;
            // retain spacing changes made by another DOCX editor too.
            const gap = ['before', 'line', 'after'].reduce((sum, name) => sum + Math.max(0, Number(val(spacing, name))), 0) / 15;
            if (Number.isFinite(gap)) tableGaps.set(p, gap);
          }
        }
        if (val(child(properties, 'pStyle')) === TABLE_ANCHOR_STYLE &&
          Array.from(marker.paragraph.children).every(c => c.localName === 'pPr' || c.localName.startsWith('bookmark')))
          tableAnchorParagraphs.add(p);
        if (val(child(properties, 'pStyle')) === TABLE_OFFSET_STYLE &&
          Array.from(marker.paragraph.children).every(c => c.localName === 'pPr' || c.localName.startsWith('bookmark')))
          tableOffsetParagraphs.set(p, marker.paragraph);
        p.setAttribute('data-word-line-rule', format.lineRule(marker.paragraph));
        const sourceSpaces = marker.paragraph.getAttributeNS(hwpSourceSpacesNamespace, 'fixed') === '1';
        const sourceWrap = parseHwpSourceWrap(marker.paragraph.getAttributeNS(hwpSourceSpacesNamespace, 'wrap'));
        if (sourceSpaces) p.setAttribute(hwpSourceSpacesAttribute, 'true');
        const wordCell = marker.paragraph.parentElement;
        const nestedWordCell = wordCell?.localName === 'tc' && wordCell.parentElement?.parentElement?.parentElement?.localName === 'tc';
        if (sourceWrap && ((sourceWrap.inlineTextBox || sourceWrap.floatingTextBox) ? textBoxStory && wordCell?.localName === 'body' :
            sourceWrap.nestedHanging ? nestedWordCell :
            sourceWrap.columnWidth !== undefined || sourceWrap.authoredBodyLines || sourceWrap.noteTerminal ? wordCell?.localName === 'body' :
            sourceWrap.columnCellHeight !== undefined || sourceWrap.expandedRowHeight !== undefined || sourceWrap.authoredCellFace || sourceWrap.cellWidthFallback ? wordCell?.localName === 'tc' && !nestedWordCell :
            sourceSpaces || sourceWrap.ends.length === 1 && wordCell?.localName === 'tc'))
          p.setAttribute(hwpSourceWrapAttribute, JSON.stringify(sourceWrap));
        const storedBodyLine = marker.paragraph.getAttributeNS(hwpBodyLineNamespace, 'bodyLine');
        const sourceInlineTableSpaces = marker.paragraph.getAttributeNS(hwpBodyLineNamespace, 'inlineTableSpaces') === '1';
        if (sourceInlineTableSpaces && p.tagName === 'P' && /^ +$/.test(p.textContent ?? '') &&
            !p.querySelector('br,img,svg,table,[data-docx-tab],[data-note-reference],[data-page-number]')) {
          const spacing = child(properties, 'spacing');
          if (val(spacing, 'lineRule') === 'exact' && val(spacing, 'line') === '2')
            p.dataset.hwpInlineTableSpaces = 'true';
        }
        if (storedBodyLine && storedBodyLine.length < 500) {
          try {
            const stored = JSON.parse(storedBodyLine), line = parseHwpBodyLine(stored.policy);
            const spacing = child(child(marker.paragraph, 'pPr'), 'spacing');
            const exactSpacing = val(spacing, 'lineRule') === 'exact' && stored.line === Math.round((line?.base ?? 0) * 15);
            const validEmpty = !line?.empty || p.tagName === 'P' && !p.textContent && exactSpacing;
            const validPrecision = !line?.precisePitch || p.tagName === 'P' && exactSpacing;
            if (line && validEmpty && validPrecision && Boolean(line.cell) === Boolean(p.closest('td,th')) &&
              val(spacing, 'lineRule') === stored.rule && Number(val(spacing, 'line')) === stored.line) {
              p.dataset.hwpBodyLine = encodeHwpBodyLine(line);
              p.style.lineHeight = `${line.base}px`;
              p.dataset.wordLineRule = 'exact';
            }
          } catch { /* Standard Word spacing remains usable without our optional policy. */ }
        }
        const tabs = format.tabs(marker.paragraph);
        if (tabs.ordinary.length) p.dataset.tabStops = JSON.stringify(tabs.ordinary);
        if (tabs.numbering.length) p.dataset.wordNumberingTabs = JSON.stringify(tabs.numbering);
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
          // This paragraph exists only to carry sectPr after a table. Keep
          // genuine blank paragraphs and any externally added content.
          if (marker.boundaryCarrier) p.remove();
        }
      }
      // Word's minimum row height reserves the cell's top/bottom margins in
      // addition to the requested text height; HTML tr height includes them.
      // Wait until every cell has its formatting before measuring that reserve.
      for (const [table, sourceTable] of sizedTables) {
        const rows = all(sourceTable, 'tr').filter(row => ancestor(row, 'tbl') === sourceTable);
        if (rows.length !== table.rows.length) continue;
        const inheritedSplit = format.tableRowSplit(sourceTable);
        rows.forEach((source, index) => {
          const split = child(child(source, 'trPr'), 'cantSplit');
          if (split) table.rows[index].dataset.rowSplit = ['0', 'false', 'off'].includes(val(split)) ? 'allow' : 'avoid';
          else if (inheritedSplit) table.rows[index].dataset.wordRowSplit = inheritedSplit;
          const property = child(child(source, 'trPr'), 'trHeight'), height = Number(val(property));
          if (table.rows[index].dataset.rowSplit === 'allow' && val(property, 'hRule') !== 'exact') {
            if (source.getAttributeNS(hwpGridRowHeightNamespace, 'continued') === '1') table.rows[index].dataset.hwpRowContinued = 'true';
            if (source.getAttributeNS(hwpGridRowHeightNamespace, 'intact') === '1') table.rows[index].dataset.hwpRowIntact = 'true';
          }
          if (height <= 0 || val(property, 'hRule') === 'auto') return;
          const row = table.rows[index], exact = val(property, 'hRule') === 'exact';
          const padding = exact ? 0 : Math.max(0, ...[...row.cells].map(cell =>
            (parseFloat(cell.style.paddingTop) || 0) + (parseFloat(cell.style.paddingBottom) || 0)));
          const cells = Array.from(source.children).filter(c => c.localName === 'tc');
          const border = exact ? 0 : Math.max(0, ...cells.map(c => {
            const probe = document.createElement('td'); format.cell(probe, c);
            const merge = child(child(c, 'tcPr'), 'vMerge');
            const sides = merge && val(merge) !== 'restart' ? ['Top', 'Bottom'] : merge ? [] : ['Bottom'];
            return Math.max(0, ...sides.map(side => {
              const width = probe.style.getPropertyValue(`border-${side.toLowerCase()}-width`);
              return (parseFloat(width) || 0) * (width.endsWith('pt') ? 4 / 3 : 1);
            }));
          }));
          let minimum = height / 15 + padding + border;
          const gridHeight = rowHeight(Number(source.getAttributeNS(hwpGridRowHeightNamespace, 'height')));
          // An external change to the standard height or merge topology
          // takes precedence over the original local grid minimum.
          if (height === 1 && val(property, 'hRule') === 'atLeast' && gridHeight !== null && gridHeight < padding + border) {
            const merges = cells.map(c => child(child(c, 'tcPr'), 'vMerge'));
            const insideMerge = merges.every(Boolean) && merges.some(merge => merge && val(merge) !== 'restart');
            const besideMergeStart = merges.some(Boolean) && merges.some(merge => !merge) &&
              merges.every(merge => !merge || val(merge) === 'restart') && cells.length === row.cells.length &&
              [...row.cells].every((cell, i) => merges[i] || (cell.children.length === 1 &&
                cell.firstElementChild?.matches('p') && cell.textContent === '' &&
                !cell.firstElementChild.querySelector(':not(span)')));
            if (insideMerge || besideMergeStart) {
              minimum = gridHeight;
              row.dataset.hwpGridRow = 'true';
            }
          }
          const storedReserve = source.getAttributeNS(hwpGridRowHeightNamespace, 'reserve');
          const reserve = hwpRowReserve(storedReserve);
          if (!exact && reserve && Math.abs(reserve.authoredHeight - minimum) <= .1 &&
              (reserve.kind === 1 || reserve.kind === 4 ? row.dataset.hwpRowContinued === 'true' : row.dataset.rowSplit === (reserve.kind === 3 ? 'avoid' : 'allow')))
            row.dataset.hwpRowReserve = storedReserve!;
          const storedIntactTop = source.getAttributeNS(hwpGridRowHeightNamespace, 'intactTop');
          const intactTop = parseHwpIntactRow(storedIntactTop);
          if (!exact && row.dataset.rowSplit === 'allow' && intactTop && Math.abs(intactTop.height - minimum) <= .16)
            row.setAttribute(hwpIntactRowAttribute, storedIntactTop!);
          row.style.height = `${minimum}px`;
          row.dataset.wordRowRule = exact ? 'exact' : 'atLeast';
        });
      }
      // An explicit cantSplit=false across the whole merge preserves its
      // permission to continue. Missing/default policy is not new evidence.
      for (const table of sizedTables.keys()) for (const [index, row] of [...table.rows].entries()) for (const cell of [...row.cells]) {
        if (cell.rowSpan > 1 && [...table.rows].slice(index, index + cell.rowSpan).every(r => rowSplitPolicy(r.dataset) === 'allow' && r.dataset.wordRowRule !== 'exact')) cell.dataset.cellSplit = 'allow';
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
        // tblInd positions a table without reducing its declared grid. The
        // percentage width is also relative to the complete text column.
        // Subtracting that indent again silently narrows columns on reopen.
        const available = Math.max(1, parentWidth ? parentWidth - (parseFloat(parent!.style.paddingLeft) || 0) - (parseFloat(parent!.style.paddingRight) || 0) - 2
          : page && margins ? (paperDimensions(page)[0] - margins.left - margins.right) * 96 / 25.4 - 1 : 624);
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
        const limit = columnWidth - borderReserve;
        // Preserve a fitting declared grid within the existing subpixel
        // border reserve. Quantization must not turn that tolerance into a
        // one-pixel shrink on every DOCX reopen.
        const fittingGrid = Math.abs(Math.min(requested, limit) - total) < 1;
        const widths = tablePixelGrid(grid, requested, fittingGrid ? Math.max(limit, Math.round(total)) : limit);
        if (!widths.length) continue;
        for (const cell of table.querySelectorAll<HTMLTableCellElement>('td[colwidth],th[colwidth]')) {
          if (cell.closest('table') !== table) continue;
          const sourceCell = sizedCells.get(cell); if (!sourceCell) continue;
          const offset = tableCellGridOffset(sourceCell), span = cell.getAttribute('colwidth')!.split(',').length;
          const columns = widths.slice(offset, offset + span);
          if (columns.length === span) cell.setAttribute('colwidth', columns.join(','));
        }
      }
      const active = new Set<string>();
      let activeStyle = '';
      let eastAsianStyle = '';
      let latinFont: string | undefined;
      let activeReflection: { value: string; block: Element; link: Element | null } | undefined;
      const reflectionMarkerNames = new Set<string>();
      for (const [name, marker] of markers) {
        if (marker.reflectionEnd) { reflectionMarkerNames.add(name); reflectionMarkerNames.add(marker.reflectionEnd); }
      }
      const reflectionMarkers = new Map<string, Element | null>();
      if (reflectionMarkerNames.size) {
        for (const el of template.content.querySelectorAll('[data-ws-marker]')) {
          const name = el.getAttribute('data-ws-marker')!;
          if (reflectionMarkerNames.has(name)) reflectionMarkers.set(name, reflectionMarkers.has(name) ? null : el);
        }
      }
      const walker = document.createTreeWalker(template.content, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT);
      const items: Node[] = []; while (walker.nextNode()) items.push(walker.currentNode);
      for (const item of items) {
        if (item instanceof HTMLElement && markers.has(item.getAttribute('data-ws-marker')!)) {
          const marker = markers.get(item.getAttribute('data-ws-marker')!)!;
          if (marker.kind === 'pageBreak' || marker.kind === 'columnBreak') continue;
          if (marker.kind === 'boxedDigit') {
            const span = document.createElement('span');
            span.setAttribute(hwpBoxedDigitAttribute, marker.id!);
            span.textContent = marker.parts![0];
            span.style.cssText = marker.style || '';
            item.replaceWith(span);
            continue;
          }
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
            if (active.size) {
              tab.setAttribute('data-comment-ids', [...active].join(' '));
              for (const id of active) { const comment = comments.find(c => c.id === id); if (comment) comment.quote += '\t'; }
            }
            item.replaceWith(tab); continue;
          }
          if (marker.kind === 'start') active.add(marker.id!);
          if (marker.kind === 'end') active.delete(marker.id!);
          if (marker.kind === 'runStart') {
            activeStyle = marker.style || '';
            eastAsianStyle = marker.eastAsianStyle || activeStyle;
            latinFont = marker.latinFont;
            activeReflection = undefined;
            const start = item.getAttribute('data-ws-marker')!;
            if (marker.reflection && marker.reflectionEnd && reflectionMarkers.get(start) === item) {
              const end = reflectionMarkers.get(marker.reflectionEnd);
              const block = item.closest(reflectionBlock), link = item.closest('a[href]');
              // The converter can split wrappers or discard markers. Require
              // both ends in the same rendered paragraph and link.
              if (end && block && end.closest(reflectionBlock) === block && end.closest('a[href]') === link &&
                (item.compareDocumentPosition(end) & Node.DOCUMENT_POSITION_FOLLOWING))
                activeReflection = { value: marker.reflection, block, link };
            }
          }
          if (marker.kind === 'runEnd') { activeStyle = ''; eastAsianStyle = ''; activeReflection = undefined; latinFont = undefined; }
          item.remove();
        } else if (item.nodeType === Node.TEXT_NODE && (active.size || activeStyle || activeReflection) && item.textContent) {
          const fragment = document.createDocumentFragment();
          const segments = activeStyle !== eastAsianStyle
            ? item.textContent.match(/[\u1100-\u11ff\u2e80-\ua4cf\uac00-\ud7af\uf900-\ufaff\uff00-\uffef]+|[^\u1100-\u11ff\u2e80-\ua4cf\uac00-\ud7af\uf900-\ufaff\uff00-\uffef]+/gu) ?? [] : [item.textContent];
          for (const text of segments) {
            const span = document.createElement('span');
            if (active.size) span.setAttribute('data-comment-ids', [...active].join(' '));
            span.style.cssText = /^[\u1100-\u11ff\u2e80-\ua4cf\uac00-\ud7af\uf900-\ufaff\uff00-\uffef]/u.test(text) ? eastAsianStyle : activeStyle;
            const font = symbolFont(span.style.cssText);
            if (!font && latinFont && activeHwpLatinFont(latinFont, span.style.fontFamily))
              span.setAttribute(hwpLatinFontAttribute, latinFont);
            if (!font && activeReflection && item.parentElement?.closest(reflectionBlock) === activeReflection.block &&
              item.parentElement?.closest('a[href]') === activeReflection.link)
              span.setAttribute(wordReflectionAttribute, activeReflection.value);
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
        if (after.textContent || (after as Element).querySelector('img,br')) {
          p.removeAttribute(wordParagraphReflectionAttribute);
          pageBreak.after(after);
        }
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
      for (const p of tableAnchorParagraphs) {
        const table = p.previousElementSibling as HTMLElement | null;
        if (table?.tagName === 'TABLE' && tableAnchor(table.dataset.tableAnchor)?.reference === 'paragraph' && !p.textContent && !p.children.length)
          p.remove();
      }
      for (const [p, source] of tableOffsetParagraphs) {
        const previous = p.previousElementSibling as HTMLElement | null, next = p.nextElementSibling as HTMLElement | null;
        if (p.textContent || p.children.length || ![previous, next].some(table => table?.tagName === 'TABLE' && tableAnchor(table.dataset.tableAnchor)?.reference === 'paragraph')) continue;
        if (p.parentNode === template.content && previous?.tagName === 'TABLE' && !previous.dataset.tableAnchor && next === verifiedLeadingTableOffsets.get(source)) {
          const spacing = child(child(source, 'pPr'), 'spacing');
          const before = Number(val(spacing, 'before')), line = Number(val(spacing, 'line'));
          const gap = `${(before + line) / 15}px`;
          if (before > 0 && cssPixels(gap)) previous.style.marginBottom = gap;
        }
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

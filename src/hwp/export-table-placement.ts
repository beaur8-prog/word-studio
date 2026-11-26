import type { HwpDocument } from '@rhwp/core';
import { strFromU8, strToU8, zipSync } from 'fflate';
import { tableAnchor } from '../table-anchor-data';
import { boundedNativeParts, nativeBodyRuns, nativePaintedControls, type NativePaintedControl } from './export-binary-pages';
import { elementRanges } from './markup-ranges';

export const nativeContinuedTableAttribute = 'data-print-native-continued-table';

type Point = { x: number; y: number };
export type TablePlacement = Point & { page: number; width: number; rows: number; columns: number };
type Carrier = { placement: TablePlacement; table: NativePaintedControl };
const epsilon = .16;
const close = (a: number, b: number) => Math.abs(a - b) <= epsilon;
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);
const hasInk = (text: string) => typeof text === 'string' && /[^\s\x00-\x1f\ufffc\ufeff]/u.test(text);

export function measuredContinuationPlacement(body: Element, sheet: Element, page: number): TablePlacement | undefined {
  const table = body.firstElementChild;
  if (!(table instanceof HTMLTableElement) || !table.hasAttribute(nativeContinuedTableAttribute) ||
      table.dataset.hwpCenteredBorders !== 'continuation' || table.querySelector('img,svg,table,[data-note-reference],[data-document-textbox]') ||
      !table.rows[0]?.cells.length || ![...table.rows[0].cells].every(cell => cell.tagName === 'TH')) return;
  const anchor = tableAnchor(table.dataset.tableAnchor);
  if (anchor?.reference !== 'paragraph' || anchor.horizontal !== 'margin' ||
      anchor.wrap || anchor.x >= .05 || anchor.y !== 0 || anchor.bottom !== 0) return;
  const bounds = table.getBoundingClientRect(), frame = body.getBoundingClientRect(), paper = sheet.getBoundingClientRect();
  const columns = table.querySelectorAll(':scope > colgroup > col').length;
  const widths = [...table.rows[0].cells].flatMap(cell => cell.getAttribute('colwidth')?.split(',').map(Number) ?? []);
  const width = widths.reduce((sum, n) => sum + n, 0);
  if (!close(bounds.left, frame.left) || !close(bounds.top, frame.top) || columns < 2 || widths.length !== columns ||
      widths.some(n => !(n > 0) || !Number.isFinite(n)) || width > frame.width + epsilon || bounds.bottom > paper.bottom) return;
  return { page, x: bounds.left - paper.left, y: bounds.top - paper.top, width, rows: table.rows.length, columns };
}

function carrierAt(doc: HwpDocument, placement: TablePlacement): Carrier | undefined {
  const controls = nativePaintedControls(doc, placement.page);
  const table = controls[0];
  if (controls.length !== 1 || table.type !== 'table' || table.controlIdx !== 0 || table.cellPath?.length ||
      Math.abs(table.w - placement.width) > .75 || !table.cells?.length) return;
  const key = [table.secIdx, table.paraIdx, table.controlIdx] as const;
  const { rowCount, colCount } = JSON.parse(doc.getTableDimensions(...key));
  const properties = JSON.parse(doc.getTableProperties(...key));
  const paragraph = JSON.parse(doc.getParaPropertiesAt(table.secIdx, table.paraIdx));
  const metrics = JSON.parse(doc.getParagraphLineMetrics(table.secIdx, table.paraIdx, '[]'));
  const paper = JSON.parse(doc.getPageInfo(placement.page));
  if (rowCount !== placement.rows || colCount !== placement.columns ||
      !properties.treatAsChar || properties.horzRelTo !== 'Para' || properties.horzAlign !== 'Left' ||
      properties.horzOffset !== 0 || properties.vertOffset !== 0 ||
      properties.outerTop <= 0 || properties.outerTop !== properties.outerBottom ||
      properties.outerTop !== properties.outerLeft || properties.outerTop !== properties.outerRight ||
      !paragraph.pageBreakBefore || paragraph.indent !== 0 || paragraph.marginLeft <= 0 || paragraph.marginRight !== 0 ||
      metrics.controlCount !== 1 || hasInk(doc.getTextRange(table.secIdx, table.paraIdx, 0, 0xffffffff)) ||
      !close(paper.marginLeft,placement.x) || !close(paper.marginTop + paper.marginHeader,placement.y) ||
      !close(table.x - placement.x, paragraph.marginLeft + properties.outerLeft / 75) ||
      !close(table.y - placement.y, properties.outerTop / 75)) return;
  const dx = placement.x - table.x, dy = placement.y - table.y;
  const runs = visibleRuns(doc, placement.page).filter(run =>
    run.secIdx === table.secIdx && run.parentParaIdx === table.paraIdx && run.cellPath?.[0]?.controlIndex === table.controlIdx);
  if (!runs.length || runs.some(run => run.x + dx < paper.bodyLeft - epsilon ||
      run.x + run.w + dx > paper.bodyRight + epsilon || run.y + dy < 0 || run.y + run.h + dy > paper.height + epsilon)) return;
  return { table, placement };
}

function correctedSection(xml: string, carriers: readonly Carrier[]) {
  const paragraphs = elementRanges(xml, 'hp:p');
  const corrections = carriers.map(({table}) => {
    const p = paragraphs[table.paraIdx]; if (!p) return;
    const tables = elementRanges(p.inner, 'hp:tbl');
    if (tables.length !== 1) return;
    const [range] = tables;
    const margin = range.inner.match(/^<hp:sz\b[^>]*\/><hp:pos\b[^>]*\/><hp:outMargin\b[^>]*\/>/)?.[0];
    if (!margin) return;
    const zero = margin.replace(/(<hp:outMargin\b[^>]*>)/, tag => tag.replace(/\b(left|right|top|bottom)="\d+"/g, '$1="0"'));
    if (zero === margin) return;
    return {index: p.start + p.open.length + range.start + range.open.length, margin, zero};
  });
  const ordered = corrections.filter(item => item !== undefined).sort((a,b) => b.index - a.index);
  if (ordered.length !== carriers.length) return;
  for (const item of ordered) {
    xml = xml.slice(0, item.index) + item.zero + xml.slice(item.index + item.margin.length);
  }
  return xml;
}

const visibleRuns = (doc: HwpDocument, page: number) =>
  (JSON.parse(doc.getPageTextLayout(page)).runs as (Point & { text: string; w: number; h: number; secIdx: number;
    parentParaIdx?: number; cellPath?: { controlIndex: number }[] })[]).filter(run => hasInk(run.text));

function translated(a: unknown, b: unknown, delta: Point): boolean {
  if (!a || !b || typeof a !== 'object' || typeof b !== 'object') return a === b;
  const p = a as Record<string, unknown>, q = b as Record<string, unknown>;
  return same(Object.keys(p), Object.keys(q)) && Object.keys(p).every(key =>
    key === 'x' || key === 'y' ? typeof p[key] === 'number' && typeof q[key] === 'number' &&
      close(q[key] - p[key], delta[key]) : translated(p[key], q[key], delta));
}

function retainsInk(source: HwpDocument, result: HwpDocument, carrier: Carrier): boolean {
  const { table, placement } = carrier;
  const [next, ...others] = nativePaintedControls(result, placement.page);
  if (!next || others.length || !close(next.x, placement.x) || !close(next.y, placement.y)) return false;
  const delta = { x: next.x - table.x, y: next.y - table.y };
  if (!translated(table, next, delta)) return false;
  const before = visibleRuns(source, placement.page), after = visibleRuns(result, placement.page);
  const paper = JSON.parse(result.getPageInfo(placement.page));
  const inside = (run: typeof before[number]) =>
    run.x >= paper.bodyLeft - epsilon && run.x + run.w <= paper.bodyRight + epsilon &&
    run.y >= 0 && run.y + run.h <= paper.height + epsilon;
  return before.length === after.length && before.every((run,i) => {
    const current = after[i];
    if (run.secIdx !== table.secIdx || run.parentParaIdx !== table.paraIdx ||
        run.cellPath?.[0]?.controlIndex !== table.controlIdx) return same(run, current);
    return inside(current) && translated(run, current, delta);
  });
}

export function restoreNativeContinuationPlacement(Core: typeof HwpDocument, document: HwpDocument,
  placements: readonly TablePlacement[] | undefined, pages: number): HwpDocument {
  if (!placements?.length || document.pageCount() !== pages) return document;
  let source: HwpDocument | undefined, result: HwpDocument | undefined;
  try {
    const bytes = document.exportHwpx(), parts = boundedNativeParts(bytes), original = new Core(bytes);
    source = original;
    if (original.pageCount() !== pages) return document;
    const carriers = placements.map(placement => carrierAt(original, placement)).filter((value): value is Carrier => !!value);
    if (!carriers.length) return document;
    const bySection = new Map<number, Carrier[]>();
    for (const carrier of carriers) {
      const section = carrier.table.secIdx, list = bySection.get(section) ?? [];
      list.push(carrier); bySection.set(section, list);
    }
    for (const [section, group] of bySection) {
      const name = `Contents/section${section}.xml`;
      const original = parts[name];
      const corrected = original && correctedSection(strFromU8(original), group);
      if (!corrected) return document;
      parts[name] = strToU8(corrected);
    }
    const candidate = new Core(zipSync(parts, { level: 1 }));
    result = candidate;
    for (const {table} of carriers)
      candidate.applyParaFormat(table.secIdx, table.paraIdx, JSON.stringify({ marginLeft: 0 }));
    if (candidate.pageCount() !== pages || !carriers.every(carrier => retainsInk(original, candidate, carrier))) return document;
    const affected = new Set(carriers.map(carrier => carrier.placement.page));
    for (let page = 0; page < pages; page++) if (!affected.has(page) &&
        (!same(visibleRuns(original, page),visibleRuns(candidate, page)) ||
        !same(nativePaintedControls(original, page),nativePaintedControls(candidate, page)))) return document;
    result = undefined;
    return candidate;
  } catch {
    // A carrier can already have a different story or a different native
    // geometry. Keep the ordinary save whenever the observed paint changes.
    return document;
  } finally {
    source?.free();
    result?.free();
  }
}

export type PaperHeaderTable = Point & { width: number; height: number; bottom: number; literal: string };

const tightText = (text: string) => text.replace(/\s/gu, '');

/** A full-width title table can originate above a two-column body's ordinary
 * top. Keep its authored paper coordinates when the printed table still uses
 * them and the column text starts just after its authored trailing margin. */
export function measuredPaperHeaderTable(body: Element, sheet: Element): PaperHeaderTable | undefined {
  const tables = body.querySelectorAll<HTMLTableElement>(':scope > table[data-table-anchor]');
  if (tables.length !== 1) return;
  const table = tables[0], anchor = tableAnchor(table.dataset.tableAnchor);
  if (!anchor || anchor.reference || anchor.horizontal || anchor.overlay || anchor.wrap ||
      table.rows.length !== 1 || table.rows[0].cells.length !== 1 ||
      table.rows[0].cells[0].rowSpan !== 1 || table.rows[0].cells[0].colSpan !== 1 ||
      table.querySelector('table,[data-document-note],[data-document-textbox],[data-note-reference]')) return;
  const column = table.nextElementSibling;
  if (!column?.matches('[data-column-count="2"]') ||
      [...body.children].slice(0, [...body.children].indexOf(table)).some(node =>
        node.textContent?.trim() || node.matches('table') || node.querySelector('table'))) return;
  const bounds = table.getBoundingClientRect(), frame = body.getBoundingClientRect(), paper = sheet.getBoundingClientRect();
  const literal = tightText(table.textContent ?? '');
  if (!literal || literal.length > 512 || bounds.top < paper.top || bounds.top >= frame.top ||
      bounds.bottom <= frame.top || bounds.bottom > frame.bottom || bounds.right > paper.right ||
      !close(bounds.left,frame.left) || Math.abs(bounds.width - frame.width) > .5 ||
      !close(bounds.left - paper.left,anchor.x) || !close(bounds.top - paper.top,anchor.y) ||
      !close(column.getBoundingClientRect().top, bounds.bottom + anchor.bottom)) return;
  return { x: anchor.x, y: anchor.y, width: bounds.width, height: bounds.height, bottom: anchor.bottom, literal };
}

function headerCellLiteral(doc: HwpDocument, table: NativePaintedControl, cells = 1) {
  const {secIdx: section, paraIdx: paragraph, controlIdx: control} = table;
  let literal = '';
  for (let cell = 0; cell < cells; cell++) {
    const count = doc.getCellParagraphCount(section, paragraph, control, cell);
    if (count < 1 || count > 24) return;
    for (let i = 0; i < count; i++) {
      const locator = JSON.stringify([{controlIndex:control, cellIndex:cell, cellParaIndex:i}]);
      literal += doc.getTextInCellByPath(section, paragraph, locator, 0, 0xffffffff);
    }
  }
  return tightText(literal);
}

/** Use the table's exact native control and retain ordinary output if the
 * placement would put content on another physical sheet or lose a run. */
export function restoreNativePaperHeaderTable(Core: typeof HwpDocument, document: HwpDocument,
  placement?: PaperHeaderTable): HwpDocument {
  if (!placement) return document;
  let source: HwpDocument | undefined, result: HwpDocument | undefined;
  try {
    const bytes = document.exportHwpx(), parts = boundedNativeParts(bytes);
    const original = source = new Core(bytes), pages = original.pageCount();
    if (pages !== document.pageCount()) return document;
    const controls = nativePaintedControls(original, 0);
    const matches = controls.filter(t => {
      if (t.type !== 'table' || t.secIdx !== 0 || t.cellPath?.length ||
          Math.abs(t.w - placement.width) > .5 || Math.abs(t.h - placement.height) > 1) return false;
      const d = JSON.parse(original.getTableDimensions(t.secIdx, t.paraIdx, t.controlIdx));
      return d.rowCount === 1 && d.colCount === 1 && headerCellLiteral(original, t) === placement.literal;
    });
    if (matches.length !== 1) return document;
    const [table] = matches;
    const props = JSON.parse(original.getTableProperties(table.secIdx, table.paraIdx, table.controlIdx));
    if (!props.treatAsChar || props.textWrap !== 'TopAndBottom' || props.horzRelTo !== 'Para' ||
        props.horzOffset || props.vertOffset || props.hasCaption || props.outerLeft <= 0 ||
        props.outerRight !== props.outerLeft || props.outerTop !== props.outerLeft || props.outerBottom !== props.outerLeft) return document;
    const name = 'Contents/section0.xml', xml = strFromU8(parts[name]);
    const p = elementRanges(xml, 'hp:p')[table.paraIdx];
    if (!p) return document;
    const ranges = elementRanges(p.inner, 'hp:tbl');
    if (ranges.length !== 1) return document;
    const [range] = ranges;
    const prefix = /^<hp:sz\b[^>]*\/><hp:pos\b[^>]*\/><hp:outMargin\b[^>]*\/>/.exec(range.inner);
    if (!prefix || !prefix[0].includes('treatAsChar="1"')) return document;
    const units = (n: number) => Math.round(n * 75);
    const position = `<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PAPER" horzRelTo="PAPER" vertAlign="TOP" horzAlign="LEFT" vertOffset="${units(placement.y)}" horzOffset="${units(placement.x)}"/>`;
    const margin = `<hp:outMargin left="0" right="0" top="0" bottom="${units(placement.bottom)}"/>`;
    const changed = prefix[0].replace(/<hp:pos\b[^>]*\/>/, position).replace(/<hp:outMargin\b[^>]*\/>/, margin);
    const at = p.start + p.open.length + range.start + range.open.length;
    parts[name] = strToU8(xml.slice(0, at) + changed + xml.slice(at + prefix[0].length));

    const candidate = result = new Core(zipSync(parts, {level:1}));
    if (candidate.pageCount() !== pages) return document;
    const updated = nativePaintedControls(candidate, 0);
    const moved = updated.find(t => t.type === 'table' && t.secIdx === table.secIdx &&
      t.paraIdx === table.paraIdx && t.controlIdx === table.controlIdx);
    if (!moved || !close(moved.x, placement.x) || !close(moved.y, placement.y) ||
        moved.w !== table.w || moved.h !== table.h) return document;
    const controlsInFlow = (list: NativePaintedControl[]) => list.map(item =>
      JSON.stringify([item.type, item.secIdx, item.paraIdx, item.controlIdx, item.w, item.h,
        item.rowStart, item.rowEnd, item.cellPath])).sort();
    if (!translated(table.cells,moved.cells,{x:moved.x - table.x,y:moved.y - table.y})) return document;
    if (!same(controlsInFlow(controls),controlsInFlow(updated))) return document;
    const ink = (doc: HwpDocument) => nativeBodyRuns(doc, 0).filter(run => hasInk(run.text));
    const inkIdentities = (doc: HwpDocument) => ink(doc).map(({x, y, ...run}) => JSON.stringify(run)).sort();
    if (!same(inkIdentities(original),inkIdentities(candidate))) return document;
    const paper = JSON.parse(candidate.getPageInfo(0));
    if (ink(candidate).some(run => run.x < -epsilon || run.y < -epsilon ||
        run.x + run.w > paper.width + epsilon || run.y + run.h > paper.height + epsilon)) return document;
    for (let n = 1; n < pages; n++) {
      if (!same(nativeBodyRuns(original,n),nativeBodyRuns(candidate,n)) ||
          !same(nativePaintedControls(original,n),nativePaintedControls(candidate,n))) return document;
    }
    result = undefined;
    return candidate;
  } catch {
    return document;
  } finally {
    source?.free();
    result?.free();
  }
}

export type RasterTablePair = {
  label: string; labelCells: number; labelSize: readonly [number, number]; diagramSize: readonly [number, number];
  picture: { x: number; y: number; width: number; height: number };
};

const singleTableCell = (table: HTMLTableElement) => {
  if (table.rows.length !== 1 || table.rows[0].cells.length !== 1 || table.querySelector('table')) return false;
  const cell = table.rows[0].cells[0];
  return cell.rowSpan === 1 && cell.colSpan === 1;
};

/** Two measured inline tables can share a sheet even when the native paste
 * adds enough extra outer margin to push the large diagram onto a new sheet. */
export function measuredRasterTablePair(body: Element, paper: DOMRect): RasterTablePair | undefined {
  if (body.children.length !== 2) return;
  const [heading, diagram] = body.children;
  if (!(heading instanceof HTMLTableElement) || !(diagram instanceof HTMLTableElement) ||
      heading.rows.length !== 1 || heading.querySelector('table') || !singleTableCell(diagram) ||
      heading.matches('[data-table-anchor]') || diagram.matches('[data-table-anchor]') ||
      heading.querySelector('img,svg,[data-document-textbox],[data-note-reference]') || diagram.textContent?.trim()) return;
  const cells = [...heading.rows[0].cells];
  if (!cells.length || cells.length > 8 || cells.some(cell => cell.rowSpan !== 1 || cell.colSpan !== 1)) return;
  const label = tightText(heading.textContent ?? '');
  const images = diagram.querySelectorAll<HTMLImageElement>('img');
  if (!label || label.length > 128 || images.length !== 1 || diagram.querySelector('svg,[data-document-textbox],[data-note-reference]')) return;
  const [image] = images, source = image.getAttribute('src') ?? '';
  if (!image.hasAttribute('data-image-floating') || !/^data:image\/(?:bmp|png);base64,/.test(source) ||
      !image.naturalWidth || !image.naturalHeight) return;
  const width = Number(image.getAttribute('width')), height = Number(image.getAttribute('height'));
  if (width <= 0 || height <= 0 || ![width,height].every(Number.isFinite)) return;
  const frame = body.getBoundingClientRect(), a = heading.getBoundingClientRect(), b = diagram.getBoundingClientRect();
  const figure = image.getBoundingClientRect();
  if (!close(a.top,frame.top) || !close(a.left,frame.left) || !close(a.left,b.left) ||
      b.top < a.bottom || b.bottom > frame.bottom + epsilon || b.height <= frame.height / 2 ||
      figure.width <= 0 || figure.height <= 0 || figure.left < paper.left || figure.right > paper.right ||
      figure.top < b.top || figure.bottom > b.bottom) return;
  return { label, labelCells:cells.length, labelSize:[a.width,a.height], diagramSize:[b.width,b.height],
    picture:{x:figure.left-paper.left,y:figure.top-paper.top,width,height} };
}

const sizeMatches = (control: NativePaintedControl, size: readonly [number, number]) =>
  Math.abs(control.w - size[0]) <= 1.1 && Math.abs(control.h - size[1]) <= 1.1;

function spilledRasterTable(doc: HwpDocument, proof: RasterTablePair) {
  const pages = doc.pageCount();
  let match: { page: number; label: NativePaintedControl; diagram: NativePaintedControl; image: NativePaintedControl } | undefined;
  for (let page = 1; page + 1 < pages; page++) {
    const previous = nativePaintedControls(doc, page - 1), current = nativePaintedControls(doc, page);
    const [label] = previous, diagram = current.find(control => control.type === 'table');
    const image = current.find(control => control.type === 'image');
    if (previous.length !== 1 || current.length !== 2 || label.type !== 'table' || !diagram || !image ||
        diagram.secIdx !== label.secIdx || diagram.paraIdx !== label.paraIdx + 1 ||
        image.secIdx !== diagram.secIdx || image.paraIdx !== diagram.paraIdx ||
        label.cellPath?.length || diagram.cellPath?.length || image.cellPath?.length !== 1 ||
        label.cells?.length !== proof.labelCells || diagram.cells?.length !== 1 ||
        !sizeMatches(label,proof.labelSize) || !sizeMatches(diagram,proof.diagramSize) ||
        !close(image.x,proof.picture.x) || !close(image.y,proof.picture.y) ||
        !close(image.w,proof.picture.width) || !close(image.h,proof.picture.height) ||
        headerCellLiteral(doc,label,proof.labelCells) !== proof.label ||
        nativeBodyRuns(doc,page).some(run => hasInk(run.text))) continue;
    const {secIdx: section, paraIdx: paragraph, controlIdx: control} = diagram;
    const format = JSON.parse(doc.getTableProperties(section,paragraph,control));
    if (!format.treatAsChar || format.hasCaption || format.outerTop <= 0 || format.outerBottom !== format.outerTop ||
        format.outerLeft !== format.outerTop || format.outerRight !== format.outerTop ||
        paragraph+1 >= doc.getParagraphCount(section) ||
        !JSON.parse(doc.getParaPropertiesAt(section,paragraph+1)).pageBreakBefore) continue;
    if (match) return;
    match = {page,label,diagram,image};
  }
  return match;
}

/** Rejoin the label and its raster table only when vertical table margin
 * removal alone makes them fit and preserves the other physical sheets. */
export function restoreNativeRasterTablePair(Core: typeof HwpDocument, document: HwpDocument,
  proofs: readonly RasterTablePair[] = [], printedPages = 0): HwpDocument {
  if (proofs.length !== 1 || document.pageCount() <= printedPages) return document;
  let source: HwpDocument | undefined, result: HwpDocument | undefined;
  try {
    const bytes = document.exportHwpx(), previous = source = new Core(bytes), pages = previous.pageCount();
    const match = spilledRasterTable(previous, proofs[0]);
    if (!match || pages !== document.pageCount()) return document;
    const {page,label,diagram,image} = match;
    const parts = boundedNativeParts(bytes), name = `Contents/section${diagram.secIdx}.xml`;
    const xml = strFromU8(parts[name]), p = elementRanges(xml,'hp:p')[diagram.paraIdx];
    if (!p) return document;
    const tables = elementRanges(p.inner,'hp:tbl');
    if (tables.length !== 1) return document;
    const [t] = tables;
    const prefix = /^<hp:sz\b[^>]*\/><hp:pos\b[^>]*\/><hp:outMargin\b[^>]*\/>/.exec(t.inner)?.[0];
    if (!prefix) return document;
    const replacement = prefix.replace(/<hp:outMargin\b[^>]*\/>/, tag => tag.replace(/\b(top|bottom)="\d+"/g,'$1="0"'));
    if (replacement === prefix) return document;
    const at = p.start + p.open.length + t.start + t.open.length;
    parts[name] = strToU8(xml.slice(0,at) + replacement + xml.slice(at+prefix.length));
    const candidate = result = new Core(zipSync(parts,{level:1}));
    if (candidate.pageCount() !== pages - 1) return document;
    const placed = nativePaintedControls(candidate,page-1);
    const [kept,moved,picture] = placed;
    const sheet = JSON.parse(candidate.getPageInfo(page-1));
    if (placed.length !== 3 || !same(kept,label) || !same(picture,image) ||
        moved.y < label.y + label.h || moved.y + moved.h > sheet.height - sheet.marginBottom - sheet.marginFooter + epsilon ||
        moved.x < 0 || moved.x + moved.w > sheet.width ||
        !translated(diagram,moved,{x:0,y:moved.y-diagram.y}) ||
        !same(nativeBodyRuns(previous,page-1).filter(run => hasInk(run.text)),
          nativeBodyRuns(candidate,page-1).filter(run => hasInk(run.text)))) return document;
    for (let i = 0; i < pages-1; i++) {
      if (i === page-1) continue;
      const old = i < page ? i : i+1;
      if (!same(nativeBodyRuns(previous,old),nativeBodyRuns(candidate,i)) ||
          !same(nativePaintedControls(previous,old),nativePaintedControls(candidate,i))) return document;
    }
    result = undefined;
    return candidate;
  } catch {
    return document;
  } finally {
    source?.free();
    result?.free();
  }
}

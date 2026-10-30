import type { HwpDocument } from '@rhwp/core';
import type { PageSettings } from '../types';
import documentLimits from '../../electron/document-limits.json';
import type { HwpProgressReporter } from './progress';
import type { HwpLayoutMetadata } from './layout-metadata';
import { cellContents, restoreCellContent, restoreHardLineBreaks, entities, type CellPath } from './cell-content';
import { serializedParts } from './serialized-parts';
import { olePreview } from './ole-preview';
import { headerDefinitions, headerPreviewPage, headerLayer, hiddenStories, editableHeaderStories } from './header-stories';
import { stageHwpFields } from './export-fields';
import { importHwpNotes, stageHwpNotes } from './note-stories';
import { stageHwpColumns } from './export-columns';
import { stageHwpSections } from './export-sections';
import { stageHwpFurniture } from './export-furniture';
import { stageHwpPictures } from './export-pictures';
import { hwpBookmark, hwpTabAttribute, stageHwpTextNavigation } from './text-navigation';
import { hwpTableGrid } from './table-grid';
import { nestedTableGeometry } from './nested-tables';
import { hwpCellBorders } from './cell-borders';
import { physicalLineHeight, paragraphCss, sourceLinePitches, cellParagraphCss } from './paragraph-layout';
import { restoreScriptFonts, type HwpFontRun } from './script-fonts';
import { emptyLineCompatibilityWarning } from '../page-empty-lines';
import { encodeHwpBodyLine } from '../hwp-body-line';
import { stageHwpTables } from './export-tables';

export interface EditableHwp { html: string; page: PageSettings; warnings: string[]; editorPages?: number }
export const HWP_CONVERSION_NOTICE = '본문, 기본 글자·문단 서식, 표와 그림을 변환합니다. 한글과 Word의 조판 방식이 달라 줄바꿈과 쪽 수가 달라질 수 있습니다. 원본 파일은 유지하고 변환한 문서는 새 파일로 저장하세요.';
// Expanded HTML includes base64 pictures, so a 21 MB image document already
// needs about 28 MB before markup. This is separate from the source-file cap.
export const HWP_HTML_LIMIT = documentLimits.hwpEditableHtmlBytes;
export function hwpHtmlBytes(value: string): number {
  // Do not allocate another document-sized buffer just to measure the Worker input.
  let bytes = value.length;
  for (let i = 0; i < value.length; i++) {
    const c = value.charCodeAt(i);
    if (c <= 0x7f) continue;
    if (c <= 0x7ff) bytes++;
    else {
      bytes += 2;
      if (c >= 0xd800 && c <= 0xdbff && i + 1 < value.length) {
        const next = value.charCodeAt(i + 1);
        if (next >= 0xdc00 && next <= 0xdfff) i++;
      }
    }
  }
  return bytes;
}
export function checkHwpHtmlSize(bytes: number) {
  if (bytes > HWP_HTML_LIMIT) throw new Error(`편집용 변환 데이터가 약 ${Math.ceil(bytes / 1024 / 1024 * 10) / 10} MiB로 한도 ${HWP_HTML_LIMIT / 1024 / 1024} MiB를 넘습니다. 원본 파일 크기가 아니라 압축을 푼 본문·그림을 편집 형식으로 변환한 크기입니다. 그림을 줄이거나 문서를 나눠 주세요.`);
}
const escape = (text: string) => text.replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!);
const fragment = (html: string) => html.match(/<!--StartFragment-->([\s\S]*?)<!--EndFragment-->/)?.[1].trim() ?? html;
type Control = { list: number; para: number; controlIndex: number; ctrlId: string; userDesc: string; props: { TreatAsChar?: number } };
type PageDef = { width: number; height: number; landscape: boolean; marginTop: number; marginRight: number; marginBottom: number; marginLeft: number; marginHeader: number; marginFooter: number; marginGutter: number };
const mm = (value: number) => Math.round(value * 25.4 / 7200 * 100) / 100;
const hu = (value: number) => Math.round(value * 7200 / 25.4);
const px = (value: number) => Math.round(value / 75 * 100) / 100;
const cssStyle = (tag: string, css: string) => tag.includes('style="') ? tag.replace(/style="([^"]*)"/, (_, old) => `style="${old};${css}"`) : tag.replace(/>$/, ` style="${css}">`);


function tableGeometry(doc: HwpDocument, section: number, para: number, control: number, html: string, pitches: Map<string, number>, tabs?: (cell: number, paragraph: number) => string): string {
  const dims = JSON.parse(doc.getTableDimensions(section, para, control));
  if (dims.cellCount > 50000 || dims.colCount > 1000 || dims.rowCount > 10000) throw new Error('편집할 표의 셀 수가 너무 많습니다.');
  const tableProperties = JSON.parse(doc.getTableProperties(section, para, control));
  const cells = Array.from({ length: dims.cellCount }, (_, index) => ({
    index, ...JSON.parse(doc.getCellInfo(section, para, control, index)), ...JSON.parse(doc.getCellProperties(section, para, control, index)),
  })).map(cell => {
    // Stored cell padding is inactive unless hasMargin is enabled. Inherited
    // table padding can be much smaller; using the inactive values wraps
    // short headings into extra lines and increases every following row.
    for (const side of ['Left', 'Right', 'Top', 'Bottom']) {
      const key = 'padding' + side;
      if (!cell.applyInnerMargin || Number(cell[key]) < 0) cell[key] = Math.max(0, Number(tableProperties[key]) || 0);
    }
    return cell;
  }).sort((a, b) => a.row - b.row || a.col - b.col);
  // Solve cumulative grid edges from both merged and unmerged cells. Dividing
  // every merged width equally expands the fine grids used by Hancom forms.
  const grid = (axis: 'col' | 'row', length: number) => {
    if (axis === 'row' && cells.some(c => c.height === 0)) {
      // Zero-height cells are automatic rows, not zero-length grid edges.
      // Pinning such an edge puts all remaining merged-cell height in the
      // final row, then CSS grows the earlier automatic rows again.
      const heights: (number | undefined)[] = Array.from({ length }, (_, row) => {
        const values = cells.filter(c => c.row === row && c.rowSpan === 1 && c.height > 0).map(c => c.height);
        return values.length ? Math.max(...values) : undefined;
      });
      const constraints = cells.filter(c => c.rowSpan > 1 && c.height > 0)
        .sort((a, b) => a.rowSpan - b.rowSpan).map(c => ({ start: c.row, count: c.rowSpan, height: c.height }));
      constraints.push({ start: 0, count: length, height: Math.max(0, Number(tableProperties.tableHeight) || 0) });
      for (const constraint of constraints) {
        const indexes = Array.from({ length: constraint.count }, (_, i) => constraint.start + i).filter(i => i < length);
        const unknown = indexes.filter(i => heights[i] === undefined);
        if (!unknown.length || !constraint.height) continue;
        const remaining = Math.max(0, constraint.height - indexes.reduce((sum, i) => sum + (heights[i] ?? 0), 0));
        for (const i of unknown) heights[i] = remaining / unknown.length;
      }
      return heights.map(height => Math.max(1, px(height ?? 0)));
    }
    return hwpTableGrid(cells, axis, length);
  };
  const columns = grid('col', dims.colCount), rows = grid('row', dims.rowCount);
  // Some stored HWP row layouts continue their cell contents onto later
  // sheets even under RowBreak. Keep that observed behavior, rather than
  // treating every multi-cell row as an indivisible form or guessing from
  // the enum alone. Repeated header cells are not continuation evidence.
  const continuationRows = new Set<number>();
  const cellPages = new Map<number, Set<number>>();
  for (const box of JSON.parse(doc.getTableCellBboxes(section, para, control))) {
    const pages = cellPages.get(box.cellIdx) ?? new Set<number>(); pages.add(box.pageIndex); cellPages.set(box.cellIdx, pages);
  }
  for (const c of cells) if ((!c.isHeader || !tableProperties.repeatHeader) && c.rowSpan === 1 &&
    (cellPages.get(c.index)?.size ?? 0) > 1) continuationRows.add(c.row);
  const style = (tag: string, css: string) => tag.includes('style="') ? tag.replace(/style="([^"]*)"/, (_, old) => `style="${old}${css}"`) : tag.replace(/>$/, ` style="${css}">`);
  let depth = 0, cellIndex = 0, row = 0, cellParagraph = 0, header = false, paragraphStyle = '';
  return html.replace(/<\/?(?:table|tr|td|p)\b[^>]*>/g, (tag, offset: number) => {
    if (/^<table\b/.test(tag)) { depth++; return depth === 1 ? style(tag, `width:${columns.reduce((sum, n) => sum + n, 0)}px;margin-bottom:${px(Math.max(0, Number(tableProperties.outerBottom) || 0))}px;`) : tag; }
    if (tag === '</table>') { depth--; return tag; }
    if (depth !== 1) return tag;
    if (/^<p\b/.test(tag)) {
      try {
        const sourceIndex = tag.match(/\bdata-hwp-cell-paragraph="(\d+)"/)?.[1];
        if (sourceIndex !== undefined) cellParagraph = Number(sourceIndex);
        const c = cells[cellIndex - 1], index = c.index;
        const properties = JSON.parse(doc.getCellParaPropertiesAt(section, para, control, index, cellParagraph));
        const char = JSON.parse(doc.getCellCharPropertiesAt(section, para, control, index, cellParagraph, 0));
        const key = `${section}:${para}:${control}:${index}:${cellParagraph}`;
        const pitch = pitches.get(key);
        if (properties.lineSpacingType === 'Percent' && Number(properties.lineSpacing) === 0 &&
          cellParagraph + 1 < doc.getCellParagraphCount(section, para, control, index)) {
          const next = cellParagraph + 1, nextKey = `${section}:${para}:${control}:${index}:${next}`;
          const firstY = pitches.get(`top:${key}`), nextY = pitches.get(`top:${nextKey}`), height = pitches.get(`height:${key}`);
          if (height && firstY !== undefined && nextY !== undefined && Math.abs(firstY - nextY) < .5 &&
            pitches.get(`page:${key}`) === pitches.get(`page:${nextKey}`) &&
            JSON.parse(doc.getLineInfoInCell(section, para, control, index, cellParagraph, 0)).lineCount === 1 &&
            JSON.parse(doc.getLineInfoInCell(section, para, control, index, next, 0)).lineCount === 1) {
            const following = fragment(doc.exportSelectionInCellHtml(section, para, control, index, next, 0, next, 0xffffffff));
            const reserve = physicalLineHeight(JSON.parse(doc.getCellParaPropertiesAt(section, para, control, index, next)),
              JSON.parse(doc.getCellCharPropertiesAt(section, para, control, index, next, 0)));
            if (!following.replace(/<[^>]*>/g, '').trim() && !/<(?:img|table|br)\b|<!--\s*rhwp:/i.test(following) && reserve >= height)
              tag = tag.replace(/>$/, ` data-hwp-zero-line="${height},${Math.round((reserve - height) * 100) / 100}">`);
          }
        }
        cellParagraph++;
        // A one-line source cell can be shorter than CSS's em * percentage.
        // Preserve its measured content pitch, without enforcing a fixed row.
        let cap: number | undefined;
        if (c.rowSpan === 1 && doc.getCellParagraphCount(section, para, control, index) === 1 &&
          JSON.parse(doc.getLineInfoInCell(section, para, control, index, 0, 0)).lineCount === 1)
          cap = px(c.height - c.paddingTop - c.paddingBottom);
        const text = html.slice(offset + tag.length, html.indexOf('</p>', offset));
        const paragraphCssValue = cellParagraphCss(properties, char, { pitch, cap,
          inkHeight: pitches.get(`height:${key}`),
          final: cellParagraph === doc.getCellParagraphCount(section, para, control, index),
          lines: JSON.parse(doc.getLineInfoInCell(section, para, control, index, cellParagraph - 1, 0)).lineCount,
          simple: !/<(?:img|table)\b|<!--\s*rhwp:/i.test(text),
        });
        if (text.includes('\t') && !tag.includes('data-tab-stops')) tag = tag.replace(/>$/, `${tabs?.(index, cellParagraph - 1) ?? ''}>`);
        return cssStyle(tag, paragraphCssValue);
      } catch { return tag; }
    }
    if (tag === '</td>') return `${cellParagraph ? '' : `<p style="${paragraphStyle}"></p>`}</${header ? 'th' : 'td'}>`;
    if (/^<tr\b/.test(tag)) {
      if (continuationRows.has(row)) tag = tag.replace(/>$/, ' data-row-split="allow">');
      return style(tag, `height:${rows[row++] ?? 1}px;`);
    }
    if (!/^<td\b/.test(tag)) return tag;
    const c = cells[cellIndex++]; if (!c) return tag;
    cellParagraph = 0; header = !!c.isHeader && !!tableProperties.repeatHeader;
    // RowBreak explicitly permits a break at row boundaries, including
    // through a vertical merge. It must still work when preceding edits
    // move a formerly single-page merge to a new boundary. This does not
    // grant permission to cut an ordinary row through its cell text.
    if (!header && c.rowSpan > 1 && (Number(tableProperties.pageBreak) === 2 || (cellPages.get(c.index)?.size ?? 0) > 1))
      tag = tag.replace(/>$/, ' data-cell-split="allow">');
    try { paragraphStyle = paragraphCss(JSON.parse(doc.getCellParaPropertiesAt(section, para, control, c.index, 0)), JSON.parse(doc.getCellCharPropertiesAt(section, para, control, c.index, 0, 0)),
      c.rowSpan === 1 ? px(c.height - c.paddingTop - c.paddingBottom) : undefined, false); }
    catch { paragraphStyle = 'line-height:100%;font-size:10pt;'; }
    const width = columns.slice(c.col, c.col + c.colSpan).map(n => Math.max(1, Math.round(n)));
    const borders = hwpCellBorders(c);
    let noWrap = false;
    if (doc.getCellParagraphCount(section, para, control, c.index) === 1) {
      const text = doc.exportSelectionInCellHtml(section, para, control, c.index, 0, 0, 0, 0xffffffff).replace(/<[^>]*>|<!--.*?-->/g, '').trim();
      noWrap = /^\d+[.)]?$/.test(text);
    }
    // TH has app defaults, but resetting them must not overwrite a fill or
    // text style already supplied by the HWP cell. Put defaults first.
    if (header) tag = tag.includes('style="')
      ? tag.replace('style="', 'style="background-color:transparent;color:#000;font-weight:normal;')
      : style(tag, 'background-color:transparent;color:#000;font-weight:normal;');
    return style(tag.replace(/^<td/, header ? '<th' : '<td').replace(/>$/, ` colwidth="${width.join(',')}">`), `width:${px(c.width)}px;height:${px(c.height)}px;padding:${px(c.paddingTop)}px ${px(c.paddingRight)}px ${px(c.paddingBottom)}px ${px(c.paddingLeft)}px;vertical-align:${['top', 'middle', 'bottom'][c.verticalAlign] ?? 'top'};${noWrap ? 'white-space:nowrap;' : ''}${borders}`);
  });
}

function pageSettings(doc: HwpDocument, section: number, warnings: Set<string>): PageSettings {
  const p: PageDef = JSON.parse(doc.getPageDef(section));
  const letter = Math.abs(p.width - 61200) < 150 && Math.abs(p.height - 79200) < 150;
  const custom = !letter && (Math.abs(p.width - 59528) > 150 || Math.abs(p.height - 84188) > 150);
  const column = JSON.parse(doc.getColumnDef(section));
  const hideEmptyLines = JSON.parse(doc.getSectionDef(section)).hideEmptyLine === true;
  if (hideEmptyLines) warnings.add(emptyLineCompatibilityWarning);
  const page: PageSettings = { size: custom ? 'Custom' : letter ? 'Letter' : 'A4', ...(custom ? { width: mm(p.width), height: mm(p.height) } : {}), orientation: p.landscape ? 'landscape' : 'portrait',
    ...(hideEmptyLines ? { hideEmptyLines: true } : {}),
    ...(column.columnCount > 1 ? { columns: { count: column.columnCount, space: px(column.spacing), balance: column.columnType === 1 } } : {}),
    margin: mm(p.marginLeft), margins: { top: mm(p.marginTop + p.marginHeader), right: mm(p.marginRight), bottom: mm(p.marginBottom + p.marginFooter), left: mm(p.marginLeft) },
    header: '', footer: '', pageNumbers: false, headerDistance: mm(p.marginTop), footerDistance: mm(p.marginBottom), gutter: mm(p.marginGutter) };
  for (const isHeader of [true, false]) {
    const hf = JSON.parse(doc.getHeaderFooter(section, isHeader, 0));
    if (hf.exists) {
      page[isHeader ? 'header' : 'footer'] = hf.text ?? '';
    }
  }
  return page;
}

/** The engine's selection HTML intentionally omits controls. Interleave their
 * HTML at the raw paragraph offsets instead of losing tables and pictures. */
export function editableHwp(doc: HwpDocument, onProgress?: HwpProgressReporter, metadata?: HwpLayoutMetadata, noteStory = false): EditableHwp {
  onProgress?.({ label: '문서의 표·그림 구조를 확인하는 중…' });
  const warnings = new Set<string>(metadata?.warnings);
  const controls: Control[] = JSON.parse(doc.getControls());
  const pitches = sourceLinePitches(doc);
  let equations: ReturnType<typeof cellContents> | undefined;
  let parts: ReturnType<typeof serializedParts> | undefined;
  const serialized = () => {
    if (parts) return parts;
    try { return parts = serializedParts(doc); }
    catch { warnings.add('일부 개체의 보조 서식 정보를 읽지 못했습니다. 원본 보기와 비교하세요.'); return parts = {}; }
  };
  const equationMetadata = () => equations ??= cellContents(doc, serialized());
  const noteStories = noteStory ? { notes: [], warnings: [] } : importHwpNotes(doc, serialized(), noteDoc => editableHwp(noteDoc, undefined, undefined, true));
  for (const warning of noteStories.warnings) warnings.add(warning);
  const noteReference = (control: { noteId?: string }) => {
    const note = noteStories.notes.find(n => n.id === control.noteId);
    return note ? `<span data-note-reference="${note.id}" data-note-kind="${note.kind}" data-note-label="${escape(note.label)}">${escape(note.label)}</span>` : undefined;
  };
  const equationHtml = (equation: { script: string; fontSize: number; color: number; width: number; height: number }) => {
    const svg = doc.renderEquationPreview(equation.script, equation.fontSize || 1200, equation.color || 0);
    const width = equation.width > 0 ? px(equation.width) : Number(svg.match(/\bwidth="([\d.]+)"/)?.[1]);
    const height = equation.height > 0 ? px(equation.height) : Number(svg.match(/\bheight="([\d.]+)"/)?.[1]);
    if (!(width > 0 && height > 0)) throw new Error('Empty equation preview');
    warnings.add('수식은 원래 모양의 그림으로 가져옵니다. 수식 명령어는 그림 설명에 보존하지만 수식 자체는 편집할 수 없습니다.');
    return `<img data-inline-image="true" data-hwp-svg="${escape(svg)}" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==" width="${width}" height="${height}" alt="${escape('수식: ' + equation.script)}" />`;
  };
  let controlPositions: Map<string, any> | undefined;
  const positionKey = (section: number, para: number, control: number, path: CellPath = []) => `${section}:${para}:${control}:${JSON.stringify(path)}`;
  const controlPosition = (section: number, para: number, control: number, path: CellPath = []) => {
    if (!controlPositions) {
      controlPositions = new Map();
      for (let page = 0; page < doc.pageCount(); page++) for (const c of JSON.parse(doc.getPageControlLayout(page)).controls ?? []) {
        // Header/footer paragraph indexes are local to their own story. They
        // can equal body indexes and must never replace the body's anchor.
        if (c.headerFooter || (c.parentParaIdx !== undefined && !c.cellPath?.length)) continue;
        const key = positionKey(c.secIdx, c.paraIdx, c.controlIdx, c.cellPath ?? []);
        if (!controlPositions.has(key)) controlPositions.set(key, { ...c, page });
      }
    }
    return controlPositions.get(positionKey(section, para, control, path));
  };
  const floatingGeometry = (position: any, props: any) => ({
    horizontalPosition: { relative: 'page', offset: Math.round(position.x * 9525) },
    verticalPosition: { relative: 'page', offset: Math.round(position.y * 9525) },
    wrap: { type: ['BehindText', 'InFrontOfText'].includes(props.textWrap) ? 0 : props.textWrap === 'TopAndBottom' ? 3 : 1 },
    behindDocument: props.textWrap === 'BehindText', allowOverlap: props.allowOverlap, layoutInCell: true,
    margins: Object.fromEntries(['left', 'right', 'top', 'bottom'].map(side => [side, Math.round((props['outerMargin' + side[0].toUpperCase() + side.slice(1)] || 0) * 127)])),
  });
  const overlays = new Map<number, any[]>();
  const croppedPicture = (html: string, props: any) => {
    const crop = [props.cropLeft / props.originalWidth, props.cropTop / props.originalHeight,
      props.cropRight / props.originalWidth, props.cropBottom / props.originalHeight];
    if (crop.every(n => Number.isFinite(n) && n >= 0 && n < 1) && crop.some(n => n > 0) && crop[0] + crop[2] < 1 && crop[1] + crop[3] < 1)
      return html.replace(/<img\b/, `<img data-hwp-image-crop="${crop.join(',')}"`);
    return html;
  };
  const pictureHtml = (section: number, para: number, control: Control): string | undefined => {
    let props: any;
    try { props = JSON.parse(doc.getPictureProperties(section, para, control.controlIndex)); } catch { return; }
    if (props.treatAsChar) return croppedPicture(fragment(doc.exportControlHtml(section, para, '[]', control.controlIndex)), props);
    const position = controlPosition(section, para, control.controlIndex);
    if (!position) return;
    let images = overlays.get(position.page);
    if (!images) {
      const page = JSON.parse(doc.getPageOverlayImages(position.page));
      images = [...page.behind ?? [], ...page.front ?? []]; overlays.set(position.page, images);
    }
    const image = images.find(i => Math.abs(i.bbox.x - position.x) < 1 && Math.abs(i.bbox.y - position.y) < 1 && Math.abs(i.bbox.width - position.w) < 1);
    const geometry = floatingGeometry(position, props);
    if (image && /^image\/(png|jpeg|bmp|gif|webp)$/.test(image.mime))
      return `<img data-inline-image="true" data-image-floating="${escape(JSON.stringify(geometry))}" src="data:${image.mime};base64,${image.base64}" width="${position.w}" height="${position.h}" alt="${escape(props.description || '그림')}" />`;
    // Square/top-and-bottom pictures are part of text flow, not overlay
    // planes. Their asset is still available from the individual control.
    const raw = fragment(doc.exportControlHtml(section, para, '[]', control.controlIndex)).match(/<img\b[^>]*>/)?.[0];
    if (!raw) return;
    return croppedPicture(raw.replace(/\s(?:width|height|alt|data-inline-image)="[^"]*"/g, '').replace(/\s*\/?>$/, ` data-inline-image="true" data-image-floating="${escape(JSON.stringify(geometry))}" width="${position.w}" height="${position.h}" alt="${escape(props.description || '그림')}" />`), props);
  };
  let layerPage = -1, layerRoot: any;
  const pageLayer = (page: number) => {
    if (layerPage !== page) {
      const tree = JSON.parse(doc.getPageLayerTreeWithProfile(page, 'print', true, true));
      if (tree.schemaVersion !== 1 || tree.unit !== 'px' || tree.coordinateSystem !== 'page-top-left-y-down') return;
      layerPage = page; layerRoot = tree.root;
    }
    return layerRoot;
  };
  const cellShapePosition = (section: number, para: number, path: CellPath, index: number, props: any) => {
    const position = controlPosition(section, para, index, path);
    if (position || !props.treatAsChar) return position;
    // The control-layout API omits inline shapes in nested tables. A unique
    // own-size subtree enclosed by the addressed source cell is unambiguous;
    // never crop the cell or take a similarly sized object elsewhere.
    const boxes = JSON.parse(doc.getTableCellBboxesByPath(section, para, JSON.stringify(path)));
    for (const box of boxes.filter((b: any) => b.cellIdx === path.at(-1)!.cellIndex)) {
      const matches: any[] = [];
      const find = (node: any, depth = 0) => {
        if (!node || depth > 40) return;
        const b = node.bounds;
        if (b && Math.abs(b.width - px(props.width)) < 1 && Math.abs(b.height - px(props.height)) < 1 &&
          b.x >= box.x - 1 && b.y >= box.y - 1 && b.x + b.width <= box.x + box.w + 1 && b.y + b.height <= box.y + box.h + 1) { matches.push(node); return; }
        node.children?.forEach((n: any) => find(n, depth + 1)); if (node.child) find(node.child, depth + 1);
      };
      find(pageLayer(box.pageIndex));
      if (matches.length === 1) { const b = matches[0].bounds; return { type: 'shape', page: box.pageIndex, x: b.x, y: b.y, w: b.width, h: b.height }; }
    }
  };
  const objectHtml = (position: any, props: any, label: string, cell = false, subtree?: any): string | undefined => {
    try {
      if (!position || !['shape', 'line', 'ellipse', 'polygon', 'group'].includes(position.type)) return;
      if (!pageLayer(position.page)) return;
      // Select one object subtree, never crop a page: a crop would also copy
      // unrelated text and background pictures underneath the shape.
      const matches: any[] = subtree ? [subtree] : [];
      const find = (node: any, depth = 0) => {
        if (!node || depth > 40) return;
        const b = node.bounds;
        if ((cell || node.layer?.zOrder === position.zOrder) && b &&
          [b.x - position.x, b.y - position.y, b.width - position.w, b.height - position.h].every(n => Math.abs(n) < 1)) {
          matches.push(node); return;
        }
        node.children?.forEach((child: any) => find(child, depth + 1));
        if (node.child) find(node.child, depth + 1);
      };
      if (!subtree) find(layerRoot);
      if (matches.length !== 1) return;
      const node = matches[0], text: string[] = [];
      let count = 0, stroke = 0;
      const supported = (node: any, depth = 0): boolean => {
        if (++count > 10000 || depth > 40) return false;
        if (node.kind === 'group') return Array.isArray(node.children) && node.children.every((child: any) => supported(child, depth + 1));
        if (node.kind === 'clipRect') return supported(node.child, depth + 1);
        return node.kind === 'leaf' && Array.isArray(node.ops) && node.ops.every((op: any) => {
          if (++count > 10000) return false;
          stroke = Math.max(stroke, Number(op.style?.strokeWidth ?? op.style?.width) || 0);
          if (op.type === 'textRun') text.push(op.displayText ?? op.text ?? '');
          return ['rectangle', 'ellipse', 'line', 'polygon', 'path', 'textRun'].includes(op.type);
        });
      };
      if (!supported(node)) return;
      const padding = Math.max(2, Math.ceil(stroke / 2 + 1));
      if (!Number.isFinite(padding) || padding > 1000) return;
      const payload = JSON.stringify({ ...node, previewPadding: padding });
      if (payload.length > 4 * 1024 * 1024) return;
      const b = node.bounds;
      const geometry = cell ? { ...floatingGeometry(position, props),
        horizontalPosition: { relative: 'column', offset: Math.round((px(props.horzOffset || 0) - padding) * 9525) },
        verticalPosition: { relative: 'paragraph', offset: Math.round((px(props.vertOffset || 0) - padding) * 9525) } }
        : floatingGeometry({ x: b.x - padding, y: b.y - padding }, props);
      warnings.add('도형·글상자는 원래 위치와 모양의 그림으로 가져옵니다. 글상자 내용은 그림 설명에 보존하지만 개체 내부는 편집할 수 없습니다.');
      return `<img data-inline-image="true" data-hwp-object="${escape(payload)}"${cell && props.treatAsChar ? '' : ` data-image-floating="${escape(JSON.stringify(geometry))}"`} src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==" width="${b.width + padding * 2}" height="${b.height + padding * 2}" alt="${escape([label || '도형', text.join('')].filter(Boolean).join(': '))}" />`;
    } catch { return; }
  };
  const shapeHtml = (section: number, para: number, control: Control): string | undefined => {
    try {
      const props = JSON.parse(doc.getShapeProperties(section, para, control.controlIndex));
      return objectHtml(controlPosition(section, para, control.controlIndex), props, control.userDesc, !!props.treatAsChar);
    } catch { return; }
  };
  const storyChanges = new Map<string, Partial<PageSettings>>(), restoredHeaders = new Set<string>();
  const definitions = headerDefinitions(serialized());
  const editableStories = editableHeaderStories(doc, definitions, isolated => editableHwp(isolated, undefined, undefined, true).html);
  // The raw selection exporter omits rich header stories. Reuse the isolated
  // header layer (never a page crop), with PAGE controls kept as live fields.
  for (const definition of definitions) {
    const key = `${definition.section}:${definition.para}`, page = pageSettings(doc, definition.section, warnings);
    const names = definition.apply === 0 ? [definition.kind, definition.kind === 'header' ? 'evenHeader' : 'evenFooter']
      : [definition.apply === 1 ? definition.kind === 'header' ? 'evenHeader' : 'evenFooter' : definition.kind];
    const text = [...definition.xml.matchAll(/<hp:t>([^<]*)<\/hp:t>/g)].map(m => m[1]).join('');
    const rich = editableStories.has(definition) || /<hp:(?:rect|line|pic|ellipse|polygon|container|tbl|equation|ole|autoNum)\b/.test(definition.xml);
    let html = editableStories.get(definition) ?? (text ? `<p>${text}</p>` : ''), preserved = editableStories.has(definition) || !rich;
    if (rich && !preserved) {
      const index = headerPreviewPage(doc, definition);
      const original = index === undefined ? undefined : headerLayer(pageLayer(index), definition.kind);
      if (original) {
        const layer = JSON.parse(JSON.stringify(original)), fields: any[] = [];
        const fieldCount = [...definition.xml.matchAll(/<hp:autoNum\b[^>]*numType="PAGE"/g)].length;
        const visit = (node: any) => {
          for (const op of node.ops ?? []) if (op.type === 'textRun' && op.displayText && op.displayText !== op.text) {
            const original = Array.from(op.text as string), display = Array.from(op.displayText as string);
            let from = 0, suffix = 0;
            while (from < Math.min(original.length, display.length) && original[from] === display[from]) from++;
            while (suffix < Math.min(original.length, display.length) - from && original.at(-suffix - 1) === display.at(-suffix - 1)) suffix++;
            const end = display.length - suffix, positions = op.displayPositions ?? op.positions;
            if (!original.slice(from, original.length - suffix).join('').trim() && /^\d+$/.test(display.slice(from, end).join('')) && positions?.length === display.length + 1) {
              fields.push({ op, from, end, bbox: { ...op.bbox, x: op.bbox.x + positions[from], width: positions[end] - positions[from] } });
            }
          }
          node.children?.forEach(visit); if (node.child) visit(node.child);
        };
        visit(layer);
        // Require an exact count; symbol projections must not be mistaken for
        // page fields in a story with unrelated special characters.
        if (fields.length === fieldCount) {
          const dynamic = fields.map(field => {
            const op = field.op, s = op.paintStyle ?? op.style, b = field.bbox;
            const left = b.x - (page.margins?.left ?? page.margin) * 96 / 25.4;
            const top = b.y - (page.headerDistance ?? 0) * 96 / 25.4;
            return `<p style="margin-left:${left}px;margin-top:${top}px;margin-bottom:0;line-height:${b.height}px;font-size:${s.fontSize}px;"><span style="font-family:${escape(JSON.stringify(s.fontFamily))};font-size:${s.fontSize}px;color:${s.color};${s.bold ? 'font-weight:bold;' : ''}"><span data-page-number="PAGE">1</span></span></p>`;
          }).join('');
          for (const field of fields) {
            const glyphs = Array.from(field.op.displayText as string);
            glyphs.fill(' ', field.from, field.end); field.op.displayText = glyphs.join('');
          }
          const b = layer.bounds;
          const image = objectHtml({ type: 'group', page: index, x: b.x, y: b.y, w: b.width, h: b.height }, { textWrap: 'InFrontOfText' }, definition.kind === 'header' ? '머리말' : '꼬리말', false, layer);
          if (image) { html = dynamic + `<p style="line-height:1px;margin:0;font-size:1px">${image}</p>`; preserved = true; }
        }
      }
    }
    if (!preserved) { warnings.add('일부 머리말·꼬리말의 개체를 복원하지 못했습니다. 원본 보기와 비교하세요.'); continue; }
    const patch = storyChanges.get(key) ?? { differentFirstPage: false };
    if (definition.apply) patch.differentOddEven = true;
    for (const name of names) { (patch as any)[name] = entities(text); (patch as any)[name + 'Html'] = rich ? html : undefined; }
    storyChanges.set(key, patch); restoredHeaders.add(key);
  }
  for (const hidden of hiddenStories(serialized())) {
    const key = `${hidden.section}:${hidden.para}`;
    const patch: Partial<PageSettings> = { ...storyChanges.get(key), differentFirstPage: true };
    if (hidden.header) { patch.firstHeader = ''; patch.firstHeaderHtml = ''; }
    if (hidden.footer) { patch.firstFooter = ''; patch.firstFooterHtml = ''; patch.firstPageNumbers = false; }
    storyChanges.set(key, patch);
  }
  const byParagraph = new Map<number, Control[]>();
  for (const c of controls) {
    if (c.list === 0) { const list = byParagraph.get(c.para) ?? []; list.push(c); byParagraph.set(c.para, list); }
  }
  const pieces: string[] = [];
  let globalPara = 0, size = 0;
  const append = (value: string) => { size += hwpHtmlBytes(value) + (pieces.length ? 1 : 0); checkHwpHtmlSize(size); pieces.push(value); };
  const firstPage = pageSettings(doc, 0, warnings);
  Object.assign(firstPage, storyChanges.get('0:0'));
  const total = Array.from({ length: doc.getSectionCount() }, (_, section) => doc.getParagraphCount(section)).reduce((sum, count) => sum + count, 0);
  onProgress?.({ label: '본문·표·그림을 변환하는 중…', completed: 0, total, unit: '문단' });
  for (let section = 0; section < doc.getSectionCount(); section++) {
    let currentPage = section ? { ...pageSettings(doc, section, warnings), ...storyChanges.get(`${section}:0`) } : firstPage;
    if (section) append(`<div data-section-page="${escape(JSON.stringify(currentPage))}"></div>`);
    const column = JSON.parse(doc.getColumnDef(section));
    let columnCount = column.columnCount ?? column.count ?? 1, columnSpace = px(column.spacing ?? 0), columnBalance = column.columnType === 1;
    const hasColumns = columnCount > 1 || Object.entries(metadata?.paragraphs ?? {}).some(([key, p]) => key.startsWith(section + ':') && (p.columns?.count ?? 1) > 1);
    let columnOpen = false;
    const openColumns = () => { if (hasColumns && !columnOpen) { append(`<div data-column-count="${columnCount}" data-column-space="${columnSpace}" data-column-balance="${columnBalance}">`); columnOpen = true; } };
    const closeColumns = () => { if (columnOpen) { append('</div>'); columnOpen = false; } };
    openColumns();
    for (let para = 0; para < doc.getParagraphCount(section); para++, globalPara++) {
      const properties = JSON.parse(doc.getParaPropertiesAt(section, para));
      const structure = metadata?.paragraphs[`${section}:${para}`];
      if (structure?.columns) { closeColumns(); columnCount = structure.columns.count; columnSpace = structure.columns.space; columnBalance = !!structure.columns.balance; openColumns(); }
      if (properties.pageBreakBefore || structure?.pageBreak) { closeColumns(); append('<div data-page-break="true"></div>'); openColumns(); }
      else if (structure?.columnBreak) append('<div data-column-break="true"></div>');
      const stories = storyChanges.get(`${section}:${para}`);
      if (para && stories) {
        closeColumns(); currentPage = { ...currentPage, ...stories, breakType: 'continuous' };
        append(`<div data-section-page="${escape(JSON.stringify(currentPage))}"></div>`); openColumns();
      }
      // Some forms are editable text tables over a paper/background picture
      // in the same anchor paragraph. Stacking both blocks doubles their
      // height. Preserve an unambiguous one-cell overlay as a local backdrop.
      const tableBackgrounds = new Map<number, { html: string; original: string }>(), backgroundControls = new Set<number>();
      const paragraphControls = byParagraph.get(globalPara) ?? [];
      for (const table of paragraphControls.filter(c => c.ctrlId === 'tbl')) {
        try {
          const props = JSON.parse(doc.getTableProperties(section, para, table.controlIndex));
          const dims = JSON.parse(doc.getTableDimensions(section, para, table.controlIndex));
          if (props.treatAsChar || props.textWrap !== 'InFrontOfText' || dims.cellCount !== 1) continue;
          const cell = JSON.parse(doc.getCellProperties(section, para, table.controlIndex, 0));
          const position = controlPosition(section, para, table.controlIndex); if (!position) continue;
          const paragraph = JSON.parse(doc.getCellParaPropertiesAt(section, para, table.controlIndex, 0, 0));
          // A paragraph-relative picture follows the text, not the cell edge.
          // Middle/bottom alignment is safe only when the source text already
          // fills the cell: otherwise its alignment reserve moves the backdrop.
          if (cell.verticalAlign !== 0) {
            const key = `${section}:${para}:${table.controlIndex}:0:0`;
            const top = pitches.get(`top:${key}`), ink = pitches.get(`height:${key}`);
            const char = JSON.parse(doc.getCellCharPropertiesAt(section, para, table.controlIndex, 0, 0, 0));
            const line = pitches.get(key) ?? physicalLineHeight(paragraph, char);
            const reserve = top === undefined || ink === undefined ? NaN
              : top - position.y - px(cell.paddingTop) - (paragraph.spacingBefore || 0) / 2 - Math.max(0, line - ink) / 2;
            if (!Number.isFinite(reserve) || Math.abs(reserve) > 2 || pitches.get(`page:${key}`) !== position.page) continue;
          }
          const candidates = paragraphControls.filter(c => c.ctrlId === 'gso' && !backgroundControls.has(c.controlIndex)).flatMap(c => {
            try {
              const picture = JSON.parse(doc.getPictureProperties(section, para, c.controlIndex)), box = controlPosition(section, para, c.controlIndex);
              if (!picture.treatAsChar || !box || box.page !== position.page || box.type !== 'image') return [];
              const overlap = Math.max(0, Math.min(box.x + box.w, position.x + position.w) - Math.max(box.x, position.x)) *
                Math.max(0, Math.min(box.y + box.h, position.y + position.h) - Math.max(box.y, position.y));
              if (overlap < Math.max(box.w * box.h, position.w * position.h) * .9) return [];
              const image = fragment(doc.exportControlHtml(section, para, '[]', c.controlIndex)).match(/<img\b[^>]*>/)?.[0];
              return image ? [{ c, box, image }] : [];
            } catch { return []; }
          });
          if (candidates.length !== 1) continue;
          const { c, box, image } = candidates[0];
          const geometry = { horizontalPosition: { relative: 'column', offset: Math.round((box.x - position.x - px(cell.paddingLeft)) * 9525) },
            verticalPosition: { relative: 'paragraph', offset: Math.round((box.y - position.y - px(cell.paddingTop) - (paragraph.spacingBefore || 0) / 2) * 9525) },
            wrap: { type: 0 }, behindDocument: true, layoutInCell: true, allowOverlap: true };
          tableBackgrounds.set(table.controlIndex, { original: image, html: image.replace(/\s(?:width|height|data-inline-image|data-image-floating)="[^"]*"/g, '').replace(/\s*\/?>$/, ` data-inline-image="true" data-image-floating="${escape(JSON.stringify(geometry))}" width="${box.w}" height="${box.h}" />`) });
          backgroundControls.add(c.controlIndex);
        } catch { /* Ambiguous overlays keep both original controls. */ }
      }
      const positions: number[] = JSON.parse(doc.getControlTextPositions(section, para));
      const textLength = doc.getParagraphLength(section, para);
      const fontRuns = new Map<number, readonly string[]>();
      const range = (start: number, end: number) => {
        const html = fragment(doc.exportSelectionHtml(section, para, start, para, end));
        // Control positions can extend past a control-only paragraph's text.
        // getTextRange takes a count, whereas selection HTML takes an end.
        if (start >= textLength || end <= start) return html;
        const text = doc.getTextRange(section, para, start, Math.min(end, textLength) - start);
        const withBreaks = restoreHardLineBreaks(html, [...text.matchAll(/\n/g)].map(m => m.index!));
        const runs: HwpFontRun[] = JSON.parse(doc.getCharShapeRuns(section, para, start, Math.min(end, textLength))).map((run: any) => {
          let fonts = fontRuns.get(run.charShapeId);
          if (!fonts) {
            const char = JSON.parse(doc.getCharPropertiesAt(section, para, run.startOffset));
            fonts = Array.isArray(char.fontFamilies) ? char.fontFamilies.map((font: unknown) => typeof font === 'string' ? font : '') : [];
            fontRuns.set(run.charShapeId, fonts!);
          }
          return { start: run.startOffset - start, end: run.endOffset - start, fonts };
        });
        return restoreScriptFonts(withBreaks, text, runs);
      };
      const whole = range(0, 0xffffffff);
      let open = whole.match(/^<p\b[^>]*>/)?.[0] ?? '<p style="">';
      if (whole.includes('\t')) open = open.replace(/>$/, `${hwpTabAttribute(equationMetadata().get(`${section}:${para}`)?.tabStops)}>`);
      const paragraphChar = JSON.parse(doc.getCharPropertiesAt(section, para, 0));
      open = cssStyle(open, paragraphCss(properties, paragraphChar));
      const pitch = pitches.get(`${section}:${para}`);
      if (pitch) open = cssStyle(open, `line-height:${pitch}px;`);
      const sizes = [...whole.matchAll(/font-size:\s*([\d.]+)(pt|px)/gi)].map(m => Number(m[1]) * (m[2].toLowerCase() === 'pt' ? 4 / 3 : 1));
      const largest = sizes.reduce((maximum, size) => Math.max(maximum, size), 0);
      const smallest = sizes.reduce((minimum, size) => Math.min(minimum, size), Infinity);
      const base = pitch ?? physicalLineHeight(properties, paragraphChar);
      if (properties.lineSpacingType === 'Percent' && !properties.fontLineHeight && Number(properties.lineSpacing) > 0 &&
        Number.isFinite(base) && base > 0 && Number(paragraphChar.fontSize) > 0 && sizes.length && largest - smallest > .1)
        open = cssStyle(open, `--hwp-body-line:${encodeHwpBodyLine({ percent: Number(properties.lineSpacing), font: Math.max(largest, Number(paragraphChar.fontSize) / 75), base })};`);
      const inline = (html: string) => html.replace(/^<p\b[^>]*>\s*/, '').replace(/<\/p>\s*$/, '');
      const onlyFloating = (html: string) => html.includes('data-image-floating') && !html.replace(/<[^>]*>/g, '').trim() &&
        [...html.matchAll(/<img\b[^>]*>/g)].every(m => m[0].includes('data-image-floating'));
      const onlyWhitespace = (html: string) => !/<(?:br|img|table|svg)\b|data-(?:note-reference|page-number|symbol-font)/i.test(html) &&
        !entities(html.replace(/<[^>]*>/g, '')).trim();
      let cursor = 0, content = '', contentBytes = 0, lastInlineTable = false;
      let inlineTail: { piece: number; line: number; width: number; left: number; available: number } | undefined;
      const appendInline = (value: string) => { contentBytes += hwpHtmlBytes(value); checkHwpHtmlSize(size + contentBytes); content += value; };
      for (const c of byParagraph.get(globalPara) ?? []) {
        if (backgroundControls.has(c.controlIndex)) continue;
        if (['secd', 'cold', 'head', 'foot', 'pgnp', 'pgct', 'pghd', 'nwno'].includes(c.ctrlId)) continue;
        const pos = Math.max(cursor, positions[c.controlIndex] ?? cursor);
        appendInline(inline(range(cursor, pos))); cursor = pos;
        if (c.ctrlId === 'bokm') {
          const ordinal = paragraphControls.filter(c => c.ctrlId === 'bokm').indexOf(c);
          const name = equationMetadata().get(`${section}:${para}`)?.controls.filter(c => c.kind === 'bookmark')[ordinal]?.bookmarkName;
          const bookmark = hwpBookmark(name);
          if (bookmark) appendInline(bookmark);
          else warnings.add('일부 책갈피 이름을 가져오지 못했습니다. 본문 내용은 유지합니다.');
          continue;
        }
        onProgress?.({ label: c.ctrlId === 'tbl' ? '표를 변환하는 중…' : c.ctrlId === 'gso' ? '그림·도형을 변환하는 중…' : '문서 개체를 변환하는 중…' });
        let controlHtml = fragment(doc.exportControlHtml(section, para, '[]', c.controlIndex));
        if (c.ctrlId === 'fn' || c.ctrlId === 'en') {
          const kind = c.ctrlId === 'fn' ? 'footNote' : 'endNote';
          const ordinal = paragraphControls.filter(x => x.ctrlId === c.ctrlId).indexOf(c);
          const note = equationMetadata().get(`${section}:${para}`)?.controls.filter(x => x.kind === kind)[ordinal];
          if (note) controlHtml = noteReference(note) ?? controlHtml;
        }
        if (c.ctrlId === 'atno') {
          const ordinal = paragraphControls.filter(c => c.ctrlId === 'atno').indexOf(c);
          const type = equationMetadata().get(`${section}:${para}`)?.controls.filter(c => c.kind === 'autoNum')[ordinal]?.numberType;
          if (type === 'PAGE' || type === 'TOTAL_PAGE') controlHtml = `<span data-page-number="${type === 'PAGE' ? 'PAGE' : 'NUMPAGES'}">1</span>`;
        }
        if (c.ctrlId === 'gso') {
          controlHtml = pictureHtml(section, para, c) ?? shapeHtml(section, para, c) ?? controlHtml;
          if (c.userDesc === 'OLE' && c.props.TreatAsChar) {
            const ordinal = paragraphControls.filter(c => c.userDesc === 'OLE').indexOf(c);
            const info = equationMetadata().get(`${section}:${para}`)?.controls.filter(c => c.kind === 'ole')[ordinal];
            const bytes = info?.binaryItem && serialized()[`BinData/${info.binaryItem}.OLE`];
            const preview = bytes ? olePreview(bytes) : undefined;
            const width = px((c.props as any).Width), height = px((c.props as any).Height);
            if (preview && width > 0 && height > 0 && width <= 10000 && height <= 10000) {
              let binary = ''; for (let i = 0; i < preview.bytes.length; i += 8192) binary += String.fromCharCode(...preview.bytes.subarray(i, i + 8192));
              controlHtml = `<img data-inline-image="true" src="data:image/bmp;base64,${btoa(binary)}" width="${width}" height="${height}" alt="${escape('OLE 미리보기' + (preview.description ? ': ' + preview.description : ''))}" />`;
              warnings.add('OLE 개체는 파일에 저장된 미리보기 그림으로 가져옵니다. 원본 OLE 데이터는 원본 한글 파일에 남아 있으며 편집본에서는 개체 내부를 편집할 수 없습니다.');
            }
          }
        }
        if (c.ctrlId === 'tbl') {
          const ordinal = (byParagraph.get(globalPara) ?? []).filter(c => c.ctrlId === 'tbl').indexOf(c);
          controlHtml = restoreCellContent(doc, section, para, c.controlIndex, ordinal, controlHtml, equationHtml, (path, index) => {
            try {
              const props = JSON.parse(doc.getCellShapePropertiesByPath(section, para, JSON.stringify(path), index));
              return objectHtml(cellShapePosition(section, para, path, index, props), props, props.description || '도형', true);
            } catch { return; }
          }, warnings, equationMetadata, restoredHeaders.has(`${section}:${para}`), (path, index, html) => {
            try {
              const props = JSON.parse(doc.getCellPicturePropertiesByPath(section, para, JSON.stringify(path), index));
              let picture = html.replace(/\s(?:alt|data-image-floating)="[^"]*"/g, '')
                .replace(/\s*\/?>$/, ` alt="${escape(props.description || '')}" />`);
              if (!props.treatAsChar) {
                const position = controlPosition(section, para, index, path);
                if (position) picture = picture.replace(/<img\b/, `<img data-image-floating="${escape(JSON.stringify(floatingGeometry(position, props)))}"`);
                else warnings.add('셀 안의 떠 있는 그림 위치를 확인하지 못해 일반 그림으로 가져왔습니다. 원본 배치를 확인하세요.');
              }
              return croppedPicture(picture, props);
            }
            catch { return html; }
          }, noteReference, (path, html) => nestedTableGeometry(doc, section, para, path, html, pitches));
          controlHtml = tableGeometry(doc, section, para, c.controlIndex, controlHtml, pitches,
            (cell, p) => hwpTabAttribute(equationMetadata().get(`${section}:${para}/t${ordinal}/c${cell}/p${p}`)?.tabStops));
          const background = tableBackgrounds.get(c.controlIndex);
          if (background) {
            const target = /(<t[dh]\b[^>]*>\s*<p\b[^>]*>)/;
            if (target.test(controlHtml)) {
              controlHtml = controlHtml.replace(target, '$1' + background.html);
              warnings.add('겹친 표와 배경 그림은 편집 가능한 표와 셀 배경 그림으로 함께 가져옵니다.');
            } else {
              controlHtml = `<p>${background.original}</p>` + controlHtml;
              warnings.add('겹친 표의 배경 그림을 별도 그림으로 보존합니다. 원본 배치를 확인하세요.');
            }
          }
        }
        if (c.ctrlId === 'eqed') {
          const equation = JSON.parse(doc.getEquationProperties(section, para, c.controlIndex, -1, -1));
          const script = equation.script || '';
          try {
            controlHtml = equationHtml({ ...equation, script });
          } catch { controlHtml = ''; }
        }
        if (!controlHtml || /^<!--/.test(controlHtml)) {
          let detail = '';
          if (c.ctrlId === 'eqed') detail = JSON.parse(doc.getEquationProperties(section, para, c.controlIndex, -1, -1)).script ?? '';
          controlHtml = `<span>[${escape(c.userDesc || c.ctrlId)}${detail ? ': ' + escape(detail) : ' — 원본 보기에서 확인'}]</span>`;
          if (c.ctrlId === 'gso') {
            try {
              const count = doc.getCellParagraphCount(section, para, c.controlIndex, 0);
              if (count) controlHtml = fragment(doc.exportSelectionInCellHtml(section, para, c.controlIndex, 0, 0, 0, count - 1, 0xffffffff));
            } catch { /* Non-textbox shapes retain a visible placeholder. */ }
          }
          warnings.add(`${c.userDesc || c.ctrlId} 개체는 편집 가능한 원본 개체로 변환하지 못해 위치에 텍스트 표시를 남깁니다.`);
        }
        if (/^<(?:table|p)\b/.test(controlHtml)) {
          const inlineTable = c.ctrlId === 'tbl' && !!JSON.parse(doc.getTableProperties(section, para, c.controlIndex)).treatAsChar;
          if (inlineTable && !doc.getTextRange(section, para, 0, 0xffffffff).trim()) {
            const lineIndex = JSON.parse(doc.getLineInfo(section, para, pos)).lineIndex;
            const char = JSON.parse(doc.getCharPropertiesAt(section, para, pos));
            const leading = structure?.lineSpacings?.[lineIndex] ??
              (['Percent', 'SpaceOnly'].includes(properties.lineSpacingType) ? physicalLineHeight(properties, char) - Number(char.fontSize) / 75 : 0);
            // An inline table occupies the line's object height, followed by
            // the source line's extra leading. Removing the placeholder text
            // line must not also remove that independent gap.
            if (Number.isFinite(leading) && leading > 0) {
              const props = JSON.parse(doc.getTableProperties(section, para, c.controlIndex));
              controlHtml = controlHtml.replace(/^<table\b[^>]*>/, tag => cssStyle(tag, `margin-bottom:${Math.round((px(Math.max(0, Number(props.outerBottom) || 0)) + leading) * 100) / 100}px;`));
            }
          }
          // Spaces before an inline table share the table's line. Keep their
          // text, but do not turn them into an additional full-height line.
          // A positive subpixel pitch also survives DOCX exact-line spacing.
          if (content.trim()) append((inlineTable && onlyWhitespace(content) ? cssStyle(open, 'line-height:0.1px;margin-bottom:0px;')
            : onlyFloating(content) ? cssStyle(open, 'line-height:1px;margin-top:0;margin-bottom:0;') : open) + content + '</p>');
          const tableWidth = Number(controlHtml.match(/^<table[^>]*style="[^"]*width:([\d.]+)px/)?.[1]);
          const pd: PageDef = JSON.parse(doc.getPageDef(section));
          if (c.ctrlId === 'tbl' && !inlineTable) {
            const props = JSON.parse(doc.getTableProperties(section, para, c.controlIndex));
            if (props.textWrap === 'InFrontOfText' && ['Page', 'Paper'].includes(props.vertRelTo) && props.vertAlign === 'Bottom' && !props.hasCaption) {
              const position = controlPosition(section, para, c.controlIndex);
              const pages = new Set(JSON.parse(doc.getTableCellBboxes(section, para, c.controlIndex)).map((b: any) => b.pageIndex));
              const height = px(pd.landscape ? pd.width : pd.height);
              const edge = height - (props.vertRelTo === 'Page' ? px(pd.marginBottom + pd.marginFooter) : 0);
              const gap = position && edge - position.y - position.h;
              if (position && pages.size === 1 && position.x >= 0 && Number.isFinite(gap) && gap >= -.5 && gap <= 1000)
                controlHtml = controlHtml.replace(/^<table\b/, `<table data-table-anchor="${escape(JSON.stringify({ x: position.x, y: Math.max(0, gap), bottom: 0,
                  overlay: props.vertRelTo === 'Page' ? 'page-bottom' : 'paper-bottom' }))}"`);
            }
          }
          if (inlineTable && columnCount === 1) {
            const position = controlPosition(section, para, c.controlIndex);
            if (position) controlHtml = controlHtml.replace(/^<table\b[^>]*>/, tag => cssStyle(tag, `margin-left:${Math.round((position.x - px(pd.marginLeft)) * 100) / 100}px;`));
          }
          const fullWidth = columnCount > 1 && tableWidth > px((pd.landscape ? pd.height : pd.width) - pd.marginLeft - pd.marginRight) / columnCount * 1.35;
          if (para === 0 && /^<table\b/.test(controlHtml) && tableWidth >= px((pd.landscape ? pd.height : pd.width) - pd.marginLeft - pd.marginRight) * .8) {
            const props = JSON.parse(doc.getTableProperties(section, para, c.controlIndex)), position = controlPosition(section, para, c.controlIndex);
            const pages = new Set(JSON.parse(doc.getTableCellBboxes(section, para, c.controlIndex)).map((b: any) => b.pageIndex));
            if (!props.treatAsChar && props.textWrap === 'TopAndBottom' && props.vertRelTo === 'Paper' && props.horzRelTo === 'Paper' && position && pages.size === 1)
              controlHtml = controlHtml.replace(/^<table\b/, `<table data-table-anchor="${escape(JSON.stringify({ x: position.x, y: position.y, bottom: px(props.outerBottom || 0) }))}"`);
          }
          // A standalone empty host paragraph is the table's live flow
          // anchor, not an extra text line. Preserve its nonnegative offset
          // instead of pinning to cached page y. The native renderer applies
          // outerTop only for some relocated hosts; retain that distance only
          // when the source line and actual table independently confirm it.
          // Visible host text, negative offsets, captions and columns need
          // their own exclusion layout and retain the explicit fallback.
          if (!inlineTable && !hasColumns && /^<table\b/.test(controlHtml) && !/^<table\b[^>]*\bdata-table-anchor=/.test(controlHtml) &&
            paragraphControls.length === 1 && !doc.getTextRange(section, para, 0, 0xffffffff).trim() &&
            tableWidth >= px((pd.landscape ? pd.height : pd.width) - pd.marginLeft - pd.marginRight) * .8) {
            const props = JSON.parse(doc.getTableProperties(section, para, c.controlIndex)), position = controlPosition(section, para, c.controlIndex);
            const offset = px(Number(props.vertOffset)), outerTop = px(Number(props.outerTop));
            const firstLine = JSON.parse(doc.getParagraphLineMetrics(section, para, '[]')).lines?.[0];
            const sourceDistance = position && firstLine ? position.y - px(pd.marginTop + pd.marginHeader) - Number(firstLine.verticalPosition) - offset : NaN;
            const y = offset + (outerTop > 0 && Math.abs(sourceDistance - outerTop) <= .5 ? outerTop : 0);
            const x = position && (props.horzRelTo === 'Paper' ? position.x : position.x - px(pd.marginLeft));
            if (props.textWrap === 'TopAndBottom' && props.vertRelTo === 'Para' && props.vertAlign === 'Top' && !props.hasCaption &&
              ['Column', 'Para', 'Paper'].includes(props.horzRelTo) && Number(props.vertOffset) >= 0 && y >= 0 && x !== undefined && x >= 0)
              controlHtml = controlHtml.replace(/^<table\b/, `<table data-table-anchor="${escape(JSON.stringify({ x, y, bottom: px(props.outerBottom || 0), reference: 'paragraph', ...(props.horzRelTo !== 'Paper' ? { horizontal: 'margin' } : {}) }))}"`);
          }
          if (c.ctrlId === 'tbl' && !inlineTable && !/^<table\b[^>]*\bdata-table-anchor=/.test(controlHtml))
            warnings.add('일부 떠 있는 표는 본문 흐름 안으로 옮깁니다. 음수 위치·호스트 본문·단·겹침·캡션은 원본과 다를 수 있습니다.');
          if (fullWidth) { closeColumns(); append('<div data-column-count="1" data-column-space="0">'); }
          const tablePiece = pieces.length;
          append(controlHtml);
          inlineTail = undefined;
          if (inlineTable && !fullWidth && columnCount === 1 && tableWidth > 0) {
            const position = controlPosition(section, para, c.controlIndex);
            const pages = new Set(JSON.parse(doc.getTableCellBboxes(section, para, c.controlIndex)).map((b: any) => b.pageIndex));
            const left = position ? position.x - px(pd.marginLeft) : 0;
            const available = px((pd.landscape ? pd.height : pd.width) - pd.marginLeft - pd.marginRight) - left;
            if (pages.size === 1 && available - tableWidth >= 1)
              inlineTail = { piece: tablePiece, line: JSON.parse(doc.getLineInfo(section, para, pos)).lineIndex, width: tableWidth, left, available };
          }
          if (fullWidth) { append('</div>'); openColumns(); }
          content = ''; contentBytes = 0;
          lastInlineTable = inlineTable;
        } else { appendInline(controlHtml.replace(/<img /g, '<img data-inline-image="true" ')); lastInlineTable = false; }
      }
      appendInline(inline(range(cursor, 0xffffffff)));
      if (inlineTail && !onlyWhitespace(content) && !onlyFloating(content) &&
        !/<(?:br|table)\b/.test(content) && JSON.parse(doc.getLineInfo(section, para, cursor)).lineIndex === inlineTail.line &&
        JSON.parse(doc.getLineInfo(section, para, 0)).lineCount === inlineTail.line + 1) {
        // Inline tables and following text can share the same source line.
        // A borderless layout row preserves this editable relationship in
        // standard DOCX too, instead of inventing a paragraph below the table.
        // Restrict this to a source table on one page and its final text line.
        const { piece, width, left, available } = inlineTail;
        const table = pieces[piece];
        const gap = [...(table.match(/^<table\b[^>]*>/)?.[0] ?? '').matchAll(/margin-bottom:([\d.]+)px/g)].at(-1)?.[1] ?? '0';
        const inner = table.replace(/^<table\b[^>]*>/, tag => cssStyle(tag, 'margin-left:0;margin-right:0;margin-bottom:0;'));
        const cell = 'padding:0;border:none;vertical-align:bottom;';
        const tailWidth = Math.max(1, available - width);
        const tail = cssStyle(open, 'margin:0;text-indent:0;text-align:left;') + content + '</p>';
        const wrapped = `<table style="width:${available}px;margin-left:${left}px;margin-bottom:${gap}px;border:none;"><tr><td colwidth="${Math.round(width)}" style="${cell}width:${width}px;">${inner}</td><td colwidth="${Math.round(tailWidth)}" style="${cell}width:${tailWidth}px;">${tail}</td></tr></table>`;
        size += hwpHtmlBytes(wrapped) - hwpHtmlBytes(table); checkHwpHtmlSize(size);
        pieces[piece] = wrapped; content = ''; contentBytes = 0;
      }
      // Keep real blank paragraphs, but avoid an extra empty paragraph after a table.
      if (content.trim() || !(byParagraph.get(globalPara) ?? []).some(c => c.ctrlId === 'tbl')) {
        if (lastInlineTable && onlyWhitespace(content)) open = cssStyle(open, 'line-height:0.1px;margin-top:0px;');
        else if (onlyFloating(content))
          open = cssStyle(open, 'line-height:1px;margin-top:0;margin-bottom:0;');
        append(open + content + '</p>');
      }
      onProgress?.({ label: '본문·표·그림을 변환하는 중…', completed: globalPara + 1, total, unit: '문단' });
    }
    closeColumns();
  }
  for (const note of noteStories.notes) append(`<div data-document-note="${note.id}" data-note-kind="${note.kind}" data-note-label="${escape(note.label)}">${note.html}</div>`);
  onProgress?.({ label: '변환한 본문과 그림을 정리하는 중…' });
  let html = pieces.join('\n').replace(/\sdata-hwp-cell-paragraph="\d+"/g, '').replace(/<!--\s*rhwp:\s*셀 안 (ColumnDef|Bookmark|PageHide)\b[\s\S]*?-->/g, (_, type: string) => {
    warnings.add(type === 'ColumnDef' ? '표 안의 단 설정은 단일 흐름으로 가져옵니다. 원본의 셀 안 배치를 확인하세요.'
      : type === 'Bookmark' ? '책갈피 위치 정보는 편집본에 전달되지 않습니다. 본문 내용은 유지합니다.'
      : '원본의 쪽 숨김 설정은 편집본에 전달되지 않습니다. 본문 내용은 유지합니다.');
    return '';
  }).replace(/<div data-column-count="[^"]*"[^>]*>\s*<\/div>/g, '').replace(/<!--\s*rhwp:([\s\S]*?)-->/g, () => {
    warnings.add('표 안에 변환하지 못한 개체가 있습니다. 해당 위치의 표시와 원본을 확인하세요.');
    return '<span>[변환하지 못한 개체 — 원본 보기에서 확인]</span>';
  });
  // Verified Hancom glyph meanings, also used by the core reader's text paint.
  // Limit this to the two forms symbols; never guess arbitrary PUA meanings.
  // This Hancom line glyph advances half an em. A plain Unicode substitution
  // advances a full em in the fallback face and doubles the cut-line's width.
  html = html.replace(/<span style="([^"]*)">([^<]*\u{F081C}[^<]*)<\/span>/gu, (tag, css: string, text: string) => {
    const font = [...css.matchAll(/(?:^|;)font-size:([\d.]+)(pt|px)/g)].at(-1);
    if (!font) return tag;
    const size = Number(font[1]) * (font[2] === 'pt' ? 4 / 3 : 1);
    const content = text.split(/(\u{F081C}+)/gu).map(part => part.includes('\u{F081C}')
      ? `<span style="${css};letter-spacing:${-size / 2}px;">${part.replace(/\u{F081C}/gu, '┈')}</span>` : part).join('');
    return `<span style="${css}">${content}</span>`;
  }).replace(/\u{F081C}/gu, '┈').replace(/\u{F012B}/gu, '(인)');
  // Some HWPX manifests label embedded BMPs as application/octet-stream.
  // Sniff only known raster signatures; never permit arbitrary data resources.
  html = html.replace(/src="data:application\/octet-stream;base64,([A-Za-z0-9+/=\s]+)"/g, (source, base64: string) => {
    const b = base64.replace(/\s/g, '');
    const mime = b.startsWith('Qk') ? 'bmp' : b.startsWith('iVBORw0KGgo') ? 'png' : b.startsWith('/9j/') ? 'jpeg' : b.startsWith('R0lGOD') ? 'gif' : '';
    return mime ? `src="data:image/${mime};base64,${b}"` : source;
  });
  // A paragraph's CSS strut otherwise inherits Word's 11pt font even when every
  // HWP run is 9pt, inflating line-height percentages throughout long tables.
  html = html.replace(/<p\b([^>]*?)style="([^"]*)"([^>]*)>(\s*<span style="([^"]*)">)/g, (_, before: string, css: string, after: string, span: string, runCss: string) => {
    const font = ['font-size', 'font-family'].map(key => runCss.match(new RegExp(`(?:^|;)${key}:([^;]+)`))?.[0]).filter(Boolean).join(';');
    return `<p${before}style="${css}${font};"${after}>${span}`;
  });
  // HWP percentage leading is based on the character em. Word's AUTO leading
  // multiplies a font's full line box instead, so preserve the physical pitch
  // for these imported paragraphs. Ordinary authored HTML keeps its own rule.
  html = html.replace(/<p\b[^>]*>/g, tag => {
    // Tabs and other source attributes must not bypass line normalization.
    // CSS percentages use the em; DOCX AUTO uses the larger font line box.
    const css = tag.match(/\sstyle="([^"]*)"/)?.[1] ?? '';
    const line = [...css.matchAll(/(?:^|;)line-height:([\d.]+)%;/g)].at(-1);
    const font = [...css.matchAll(/(?:^|;)font-size:([\d.]+)(pt|px)(?=;|$)/g)].at(-1);
    // Do not consume the semicolon: adjacent declarations can override one
    // another, and skipping the second would replace measured exact leading
    // with an earlier percentage again.
    if (!line || !font || [...css.matchAll(/(?:^|;)line-height:([^;]+)(?=;|$)/g)].at(-1)?.[1] !== line[1] + '%') return tag;
    const pitch = Number(font[1]) * (font[2] === 'pt' ? 4 / 3 : 1) * Number(line[1]) / 100;
    return cssStyle(tag, `line-height:${Math.max(.1, Math.round(pitch * 100) / 100)}px;`);
  });
  checkHwpHtmlSize(hwpHtmlBytes(html));
  warnings.add('복잡한 표 테두리·셀 크기, 글상자·수식·각주·필드 등 고급 서식은 완전히 호환되지 않습니다. 편집본과 원본을 비교하세요.');
  return { html: html || '<p></p>', page: firstPage, warnings: [...warnings] };
}

export function setHwpPage(doc: HwpDocument, page: PageSettings, section = 0) {
  doc.setSectionDef(section, JSON.stringify({ hideEmptyLine: page.hideEmptyLines === true }));
  const m = page.margins ?? { top: page.margin, right: page.margin, bottom: page.margin, left: page.margin };
  const letter = page.size === 'Letter';
  const header = Math.min(m.top, page.headerDistance ?? Math.max(0, m.top / 2 - 3));
  const footer = Math.min(m.bottom, page.footerDistance ?? Math.max(0, m.bottom / 2 - 3));
  doc.setPageDef(section, JSON.stringify({ width: page.size === 'Custom' ? hu(page.width!) : letter ? 61200 : 59528, height: page.size === 'Custom' ? hu(page.height!) : letter ? 79200 : 84188, landscape: page.orientation === 'landscape',
    // HWP top/bottom margins locate the furniture at the paper edge; the
    // header/footer areas are the additional distance to the body boundary.
    marginTop: hu(header), marginRight: hu(m.right), marginBottom: hu(footer), marginLeft: hu(m.left),
    marginHeader: hu(m.top - header), marginFooter: hu(m.bottom - footer), marginGutter: hu(page.gutter ?? 0) }));
  if (page.columns) doc.setColumnDef(section, page.columns.count, page.columns.balance ? 1 : 0, 1, Math.round(page.columns.space * 75));
  for (const isHeader of [true, false]) {
    const text = page[isHeader ? 'header' : 'footer'];
    if (text) { doc.createHeaderFooter(section, isHeader, 0); doc.insertTextInHeaderFooter(section, isHeader, 0, 0, 0, text); }
  }
}

export function serializeHwp(Core: typeof HwpDocument, html: string, page: PageSettings, format: 'hwp' | 'hwpx', onProgress?: HwpProgressReporter) {
  if (typeof html !== 'string' || !html) throw new Error('저장할 편집 문서가 없습니다.');
  checkHwpHtmlSize(hwpHtmlBytes(html));
  if (!['hwp', 'hwpx'].includes(format)) throw new Error('지원하지 않는 한글 저장 형식입니다.');
  let doc = Core.createEmpty();
  try {
    // A real Hancom template supplies valid HWP 5.1 DocInfo and default styles.
    doc.createBlankDocument();
    const breakToken = 'WORDSTUDIO-PAGE-BREAK-' + crypto.randomUUID();
    const sections = stageHwpSections(html, page);
    const furniture = stageHwpFurniture(sections.html, sections.settings);
    const notes = stageHwpNotes(furniture.html);
    const tables = stageHwpTables(notes.html);
    const pictures = stageHwpPictures(tables.html);
    const fields = stageHwpFields(pictures.html);
    const navigation = stageHwpTextNavigation(fields.html);
    const columns = stageHwpColumns(navigation.html, page);
    const converted = columns.html.replace(/<div\b[^>]*data-page-break=["'][^"']*["'][^>]*>[\s\S]*?<\/div>/gi, `<p>${breakToken}</p>`);
    checkHwpHtmlSize(hwpHtmlBytes(converted));
    // The first pasted paragraph inherits the host's paragraph shape even on
    // the multi-paragraph path. Sacrifice a leading blank, then remove both
    // sentinels. The writer rebuilds section controls from the section model.
    onProgress?.({ label: '본문·표·그림을 한글 문서로 변환하는 중…' });
    const pasted = JSON.parse(doc.pasteHtml(0, 0, 0, '<p></p>' + converted + '<p></p>'));
    if (!pasted.ok) throw new Error('문서 내용을 한글 형식으로 변환하지 못했습니다.');
    onProgress?.({ label: '한글 문서의 쪽 나누기를 반영하는 중…' });
    const last = doc.getParagraphCount(0) - 1;
    if (last > 0 && doc.getParagraphLength(0, last) === 0) doc.deleteParagraph(0, last);
    if (doc.getParagraphCount(0) > 1) doc.deleteParagraph(0, 0);
    for (let p = 0; p < doc.getParagraphCount(0); p++) {
      if (doc.getTextRange(0, p, 0, 100) !== breakToken) continue;
      if (p + 1 < doc.getParagraphCount(0) && doc.getTextRange(0, p + 1, 0, 100) !== breakToken) doc.deleteParagraph(0, p);
      else doc.deleteText(0, p, 0, breakToken.length);
      doc.applyParaFormat(0, p, JSON.stringify({ pageBreakBefore: true }));
    }
    setHwpPage(doc, page);
    const withPictures = pictures.restore(Core, doc);
    if (withPictures !== doc) { doc.free(); doc = withPictures; }
    const withFields = fields.restore(Core, doc);
    if (withFields !== doc) { doc.free(); doc = withFields; }
    const withNavigation = navigation.restore(Core, doc);
    if (withNavigation !== doc) { doc.free(); doc = withNavigation; }
    const withTables = tables.restore(Core, doc);
    if (withTables !== doc) { doc.free(); doc = withTables; }
    const withColumns = columns.restore(Core, doc);
    if (withColumns !== doc) { doc.free(); doc = withColumns; }
    const withNotes = notes.restore(Core, doc);
    if (withNotes !== doc) { doc.free(); doc = withNotes; }
    const withSections = sections.restore(Core, doc, setHwpPage);
    if (withSections !== doc) { doc.free(); doc = withSections; }
    const withFurniture = furniture.restore(Core, doc);
    if (withFurniture !== doc) { doc.free(); doc = withFurniture; }
    onProgress?.({ label: `${format.toUpperCase()} 파일을 만드는 중…` });
    const output = format === 'hwp' ? doc.exportHwpWithReport() : doc.exportHwpxWithReport();
    try {
      const report = JSON.parse(output.contentLoss());
      const bytes = output.takeBytes();
      if (bytes.byteLength > documentLimits.documentFileBytes) throw new Error(`변환한 한글 파일이 ${documentLimits.documentFileBytes / 1024 / 1024} MiB를 넘습니다.`);
      // A successful writer alone is insufficient: actually parse its result.
      onProgress?.({ label: '변환한 한글 파일을 다시 읽어 확인하는 중…' });
      const check = new Core(bytes);
      let pages = 0;
      try { pages = check.pageCount(); if (!check.getSectionCount() || !pages) throw new Error('저장한 한글 문서를 다시 읽지 못했습니다.'); }
      finally { check.free(); }
      return { bytes, pages, warnings: report.count ? [`한글 저장 엔진에서 내용 손실 ${report.count}건을 보고했습니다. ${JSON.stringify(report.losses).slice(0, 1600)}`] : [] };
    } finally { output.free(); }
  } finally { doc.free(); }
}

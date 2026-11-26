import type { HwpDocument } from '@rhwp/core';
import { hwpWrapFingerprint, normalizeHwpWrapText } from '../hwp-text-advances';
import type { AuthoredNativeTableFragment } from '../table-anchor-data';
import { nativeBodyRuns, nativePaintedControls } from './export-binary-pages';

type Box = { page: number; x: number; y: number; w: number; h: number;
  rowCount: number; colCount: number; cells?: { row: number; col: number }[] };
type Props = { pageBreak: number; repeatHeader: boolean; hasCaption: boolean; allowOverlap: boolean;
  horzRelTo: string; horzAlign: string; outerTop: number; outerBottom: number; outerLeft: number; outerRight: number };

export function sourceNativeTableFragment(doc: HwpDocument, section: number, paragraph: number, control: number,
  box: Box, properties: Props): AuthoredNativeTableFragment | undefined {
  if (box.page < 1 || properties.pageBreak !== 2 || !properties.repeatHeader || properties.hasCaption ||
      properties.allowOverlap || properties.horzRelTo !== 'Para' || properties.horzAlign !== 'Left' ||
      box.colCount < 2 || box.colCount > 6 || !box.cells?.length) return;
  try {
    const rows = new Set(box.cells.map(cell => cell.row)).size;
    if (rows < 2 || rows > 12 || rows >= box.rowCount || box.cells.length !== rows * box.colCount ||
        !box.cells.every(cell => cell.row >= 0 && cell.row < rows)) return;
    const paper = JSON.parse(doc.getPageInfo(box.page)), width = paper.bodyRight - paper.bodyLeft;
    const room = paper.footerArea.y - box.y - box.h;
    if (paper.columns.length !== 1 || box.x < paper.bodyLeft || box.x + box.w > paper.bodyRight ||
        room < 0 || room > 12 || box.y < paper.marginTop + paper.marginHeader + 24) return;
    const sameTable = (shape: { secIdx: number; paraIdx: number; controlIdx: number }) =>
      shape.secIdx === section && shape.paraIdx === paragraph && shape.controlIdx === control;
    const controls = nativePaintedControls(doc, box.page);
    if (controls.length < 2 || controls.length > 3 || !controls.some(sameTable) ||
        controls.some(item => !sameTable(item) && (item.type !== 'table' || item.y + item.h >= box.y))) return;
    const next = nativePaintedControls(doc, box.page + 1);
    if (!next.some(sameTable)) return;
    const runs = nativeBodyRuns(doc, box.page).filter(run => normalizeHwpWrapText(run.text));
    if (runs.some(run => !run.cellPath?.length)) return;
    const ink = runs.filter(run => run.secIdx === section && run.parentParaIdx === paragraph &&
      run.cellPath?.length === 1 && (run.cellPath[0] as { controlIndex: number }).controlIndex === control);
    const text = ink.map(run => run.text).join('');
    if (!normalizeHwpWrapText(text) || ink.some(run => run.x < box.x - .15 || run.x + run.w > box.x + box.w + .15 ||
        run.y < box.y - .15 || run.y + run.h > paper.footerArea.y + .15)) return;
    const { outerTop: top, outerBottom: bottom, outerLeft: left, outerRight: right } = properties;
    return { page: box.page + 1, x: box.x, y: box.y, width: box.w, bodyWidth: width, paperHeight: paper.height,
      rows, columns: box.colCount, top, bottom, left, right, fingerprint: hwpWrapFingerprint(text) };
  } catch { /* Preserve normal placement without authoritative layout. */ }
}

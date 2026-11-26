import type { HwpDocument } from '@rhwp/core';
import { cssPixels, parseHwpCellLine } from '../hwp-cell-line';
import { nativeBodyRuns, nativePaintedControls, type NativeBodyRun, type NativePaintedControl } from './export-binary-pages';

type EmptyRow = { row: number; font: number; height: number };
export type EmptyTableRows = { preceding: string; cells: string[][]; rows: EmptyRow[]; width: number; height: number };
type NativeCell = { row: number; col: number; x: number; y: number; w: number; h: number };
type Table = Omit<NativePaintedControl,'cells'> & { cells: NativeCell[] };
type Target = { page: number; table: Table; proof: EmptyTableRows; cells: number[] };
const tolerance = .16;
const defaultEmptyFont = 1000;
const units = (pixels: number) => Math.round(pixels * 75);
const normalizedSpace = (text: string) => text.replace(/\u00a0/gu,' ');
const same = (a: unknown,b: unknown) => JSON.stringify(a) === JSON.stringify(b);

function authoredEmptyRow(row: HTMLTableRowElement,index: number,cells: HTMLTableCellElement[]): EmptyRow | undefined {
  if (row.dataset.rowSplit !== 'avoid' || !row.style.height) return;
  const height = row.getBoundingClientRect().height;
  if (!(height > 0 && units(height) < defaultEmptyFont)) return;
  const fonts = cells.map(cell => {
    const paragraph = cell.firstElementChild;
    if (cell.children.length !== 1 || !(paragraph instanceof HTMLParagraphElement) ||
        paragraph.childNodes.length || paragraph.dataset.wordLineRule !== 'exact') return;
    const declared = cssPixels(paragraph.dataset.hwpCellAuthoredSize);
    const measured = cssPixels(getComputedStyle(paragraph).fontSize);
    const policy = parseHwpCellLine(paragraph.dataset.hwpCellLine);
    if (declared === null || measured === null || !policy || Math.abs(declared-measured) > .05 ||
        declared > height + tolerance) return;
    return units(declared);
  });
  const font = fonts[0];
  if (!font || font >= defaultEmptyFont || !fonts.every(value=>value===font)) return;
  return {row:index,font,height};
}

/** A whole table already fits in the printed lane. Save its explicit
 * empty cell sizes when a default paragraph strut stretches a spacer row. */
export function measuredEmptyTableRows(body: Element): EmptyTableRows[] {
  const frame = body.getBoundingClientRect(), output: EmptyTableRows[] = [];
  for (const table of body.querySelectorAll<HTMLTableElement>('table[data-word-table-layout="fixed"]')) {
    if (table.parentElement?.closest('table') || table.dataset.tableAnchor ||
        table.querySelector('table,img,svg,br,canvas,[data-note-reference],[data-docx-tab],[data-document-field]') ||
        table.rows.length < 2 || table.rows.length > 16) continue;
    const previous = table.previousElementSibling, literal = previous?.textContent ?? '';
    if (!(previous instanceof HTMLParagraphElement) || previous.dataset.printParagraphFragment ||
        !literal.trim() || literal.length > 512) continue;
    const {width,height,bottom,top} = table.getBoundingClientRect(), remaining = frame.bottom-bottom;
    if (!(width > 0 && height > 0 && top >= frame.top && remaining >= -tolerance &&
        units(remaining) < defaultEmptyFont)) continue;
    const rows = [...table.rows], cells = rows.map(row=>[...row.cells]), columns = cells[0].length;
    if (columns < 2 || rows.length * columns > 128 ||
        cells.some(row=>row.length!==columns || row.some(cell=>cell.colSpan!==1 || cell.rowSpan!==1))) continue;
    const fits = rows.flatMap((row,i)=>authoredEmptyRow(row,i,cells[i]) ?? []);
    if (!fits.length) continue;
    const texts = cells.map(row=>row.map(cell=>cell.textContent ?? ''));
    if (!texts.flat().some(text=>text.trim()) || texts.flat().some(text=>text.length>128 || /[\t\r\n\x00-\x1f\ufffc]/u.test(text))) continue;
    output.push({preceding:literal,cells:texts,rows:fits,width,height});
  }
  return output;
}

const tableKey = (table: NativePaintedControl) => [table.secIdx,table.paraIdx,table.controlIdx] as const;
const belongs = (run: NativeBodyRun,table: Table) =>
  run.secIdx === table.secIdx && run.parentParaIdx === table.paraIdx && run.cellPath?.length === 1 &&
  (run.cellPath[0] as {controlIndex:number}).controlIndex === table.controlIdx;
const ink = (doc: HwpDocument,page: number) => nativeBodyRuns(doc,page).filter(run=>run.text.trim());

function targetOnExtraSheet(doc: HwpDocument, proofs: readonly EmptyTableRows[]): Target | undefined {
  for (let page=1; page+1<doc.pageCount(); page++) {
    const [part,...others] = nativePaintedControls(doc,page);
    if (!part || others.length || part.type!=='table' || part.cellPath?.length ||
        part.paraIdx < 1 || part.paraIdx+1 >= doc.getParagraphCount(part.secIdx)) continue;
    const table = part as Table, key = tableKey(table), section = table.secIdx, index = table.paraIdx;
    const {rowCount,colCount} = JSON.parse(doc.getTableDimensions(...key));
    const proof = proofs.find(proof=>proof.cells.length===rowCount && proof.cells[0]?.length===colCount &&
      Math.abs(proof.width-table.w)<=tolerance && table.h > proof.height &&
      normalizedSpace(doc.getTextRange(section,index-1,0,0xffffffff))===normalizedSpace(proof.preceding));
    if (!proof || !table.cells || table.cells.length!==rowCount*colCount ||
        (table.rowStart ?? 0)!==0 || (table.rowEnd ?? rowCount)!==rowCount ||
        !JSON.parse(doc.getTableProperties(...key)).treatAsChar ||
        !JSON.parse(doc.getParaPropertiesAt(section,index+1)).pageBreakBefore ||
        JSON.parse(doc.getPageOfPosition(section,index+1)).page!==page+1 ||
        !ink(doc,page).length || ink(doc,page).some(run=>!belongs(run,table))) continue;
    const shorts = new Set(proof.rows.map(row=>row.row)), cells: number[] = [];
    for (let cell=0; cell<table.cells.length; cell++) {
      const loc = JSON.parse(doc.getCellInfo(...key,cell)), expected = proof.cells[loc.row]?.[loc.col];
      if (loc.colSpan!==1 || loc.rowSpan!==1 || expected===undefined ||
          normalizedSpace(doc.getTextInCell(...key,cell,0,0,0xffffffff))!==normalizedSpace(expected)) break;
      cells.push(cell);
    }
    if (cells.length!==table.cells.length) continue;
    const empty = cells.filter(cell=>shorts.has(JSON.parse(doc.getCellInfo(...key,cell)).row));
    if (empty.some(cell=>doc.getCellParagraphCount(...key,cell)!==1 || doc.getCellParagraphLength(...key,cell,0)!==0 ||
        JSON.parse(doc.getCellCharPropertiesAt(...key,cell,0,0)).fontSize!==defaultEmptyFont)) continue;
    return {page,table,proof,cells:empty};
  }
}

function verifiedLayout(prior: HwpDocument,fresh: HwpDocument,target: Target) {
  const {table,proof} = target, pages = prior.pageCount();
  if (fresh.pageCount()!==pages) return false;
  let found = false;
  for (let page = 0; page < pages; page++) {
    const old = ink(prior,page), next = ink(fresh,page);
    const controls = nativePaintedControls(prior,page), changed = nativePaintedControls(fresh,page);
    const belongsToTarget = (part: NativePaintedControl)=>same(tableKey(part),tableKey(table)) && part.type==='table';
    const before = controls.find(belongsToTarget) as Table | undefined;
    const after = changed.find(belongsToTarget) as Table | undefined;
    if (!same(controls.filter(part=>!belongsToTarget(part)),changed.filter(part=>!belongsToTarget(part))) ||
        !same(old.filter(run=>!belongs(run,table)),next.filter(run=>!belongs(run,table)))) return false;
    if (!before && !after) continue;
    if (!before || !after || found || after.h >= before.h || !same([before.x,before.y,before.w],[after.x,after.y,after.w]) ||
        Math.abs(after.h-proof.height) >= Math.abs(before.h-proof.height) ||
        !same(before.cells.map(({row,col,x,w})=>[row,col,x,w]),after.cells.map(({row,col,x,w})=>[row,col,x,w])) ||
        proof.rows.some(row=>{
          const a=before.cells.find(cell=>cell.row===row.row), b=after.cells.find(cell=>cell.row===row.row);
          return !a || !b || b.h>=a.h || Math.abs(b.h-row.height)>=Math.abs(a.h-row.height);
        })) return false;
    const source = old.filter(run=>belongs(run,table)), placed = next.filter(run=>belongs(run,table));
    const withoutVerticalPosition = ({y,charShapeId,...run}: NativeBodyRun) => run;
    const paper = JSON.parse(fresh.getPageInfo(page));
    if (!same(source.map(withoutVerticalPosition),placed.map(withoutVerticalPosition)) ||
        placed.some((run,index)=>{
          const cell = (run.cellPath![0] as {cellIndex:number}).cellIndex;
          return !before.cells[cell] || !after.cells[cell] ||
            Math.abs((run.y-after.cells[cell].y) - (source[index].y-before.cells[cell].y)) > tolerance ||
            run.y<0 || run.y+run.h>paper.height+tolerance || run.x<0 || run.x+run.w>paper.width+tolerance;
        })) return false;
    found = true;
  }
  return found;
}

export function restoreNativeEmptyTableRows(Core: typeof HwpDocument, doc: HwpDocument,
  proofs: readonly EmptyTableRows[] | undefined, printedPages: number | undefined, format: 'hwp'|'hwpx') {
  if (!proofs?.length || !printedPages || doc.pageCount()<=printedPages) return doc;
  let prior: HwpDocument | undefined, candidate: HwpDocument | undefined, binaryPrior: HwpDocument | undefined, binaryAfter: HwpDocument | undefined;
  try {
    const bytes = doc.exportHwpx();
    prior = new Core(bytes);
    if (prior.pageCount()!==doc.pageCount()) return doc;
    const target = targetOnExtraSheet(prior,proofs);
    if (!target) return doc;
    const {table,proof,cells} = target, key = tableKey(table);
    candidate = new Core(bytes);
    for (const cell of cells) {
      const {row} = JSON.parse(candidate.getCellInfo(...key,cell));
      const fontSize = proof.rows.find(spec=>spec.row===row)?.font;
      if (!fontSize || !JSON.parse(candidate.applyCharFormatInCell(...key,cell,0,0,0,JSON.stringify({fontSize}))).ok ||
          candidate.getCellParagraphLength(...key,cell,0)!==0 ||
          JSON.parse(candidate.getCellCharPropertiesAt(...key,cell,0,0)).fontSize!==fontSize) return doc;
    }
    if (!verifiedLayout(prior,candidate,target)) return doc;
    if (format==='hwp') {
      binaryPrior = new Core(prior.exportHwp());
      binaryAfter = new Core(candidate.exportHwp());
      if (!verifiedLayout(binaryPrior,binaryAfter,target)) return doc;
    }
    const result = candidate;
    candidate = undefined;
    return result;
  } catch {
    return doc;
  } finally {
    binaryAfter?.free();binaryPrior?.free();candidate?.free();prior?.free();
  }
}

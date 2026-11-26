'use strict';

const { createHash } = require('node:crypto');
const { unzipSync, strFromU8 } = require('fflate');

function readPages(document) {
  return Array.from({ length: document.pageCount() }, (_, page) => {
    const paper = JSON.parse(document.getPageInfo(page));
    const runs = JSON.parse(document.getPageTextLayout(page)).runs;
    const controls = JSON.parse(document.getPageControlLayout(page)).controls;
    const visible = runs.filter(run => !run.headerFooter && run.paraIdx !== 0xffffffff && run.text.trim());
    return {
      index: page, section: paper.sectionIndex, printedNumber: paper.pageNumber,
      bodyRuns: visible.filter(run => !run.cellPath?.length).length,
      cellRuns: visible.filter(run => run.cellPath?.length).length,
      controls: controls.filter(control => !control.headerFooter).length,
      storyRuns: runs.filter(run => run.headerFooter && run.text.trim()).length,
      top: paper.marginTop + paper.marginHeader, bottom: paper.footerArea.y,
    };
  });
}

function paragraphPositions(document) {
  const rows = [], unmapped = [];
  for (let section = 0; section < document.getSectionCount(); section++) {
    for (let paragraph = 0; paragraph < document.getParagraphCount(section); paragraph++) {
      const { stored, controlCount, lines } = JSON.parse(document.getParagraphLineMetrics(section, paragraph, '[]'));
      const visibleText = !!document.getTextRange(section, paragraph, 0, 0xffffffff).trim();
      let location;
      try {
        location = JSON.parse(document.getPageOfPosition(section, paragraph));
      } catch (error) {
        if (typeof error !== 'string' || !/^렌더링 오류: 문단 .*이 페이지에 없습니다$/.test(error)) throw error;
      }
      if (!location?.ok) {
        unmapped.push({ section, paragraph, visibleText, controlCount });
        continue;
      }
      const { pageBreakBefore, keepWithNext, keepLines, spacingBefore, spacingAfter } =
        JSON.parse(document.getParaPropertiesAt(section, paragraph));
      rows.push({
        section, paragraph, page: location.page,
        units: document.getParagraphLength(section, paragraph),
        visibleText,
        controlCount, pageBreakBefore, keepWithNext, keepLines, spacingBefore, spacingAfter, stored,
        lines: lines.map(line => [line.verticalPosition, line.height, line.spacing]),
      });
    }
  }
  return { rows, unmapped };
}

function empty(page) {
  return !page.bodyRuns && !page.cellRuns && !page.controls;
}

function logicalBody(document) {
  const sections = [];
  for (let section = 0; section < document.getSectionCount(); section++) {
    const paragraphs = [];
    for (let paragraph = 0; paragraph < document.getParagraphCount(section); paragraph++) {
      const { controlCount } = JSON.parse(document.getParagraphLineMetrics(section, paragraph, '[]'));
      paragraphs.push([document.getTextRange(section, paragraph, 0, 0xffffffff), controlCount]);
    }
    sections.push(paragraphs);
  }
  return JSON.stringify(sections);
}

function packagePayload(document) {
  const archive = unzipSync(document.exportHwpx());
  const elements = [], pictures = [];
  for (const path of Object.keys(archive).sort()) {
    if (/^Contents\/section\d+\.xml$/.test(path)) {
      const xml = strFromU8(archive[path]);
      const body = [...xml.matchAll(/<(hp:t)(?:\s[^>]*)?\/>|<(hp:t)(?:\s[^>]*)?>[\s\S]*?<\/hp:t>/g)].map(match => match[0]);
      elements.push([path, body, [...xml.matchAll(/<hp:(tbl|tc|pic|lineBreak|tab)\b/g)].map(match => match[1])]);
    } else if (path.startsWith('BinData/')) {
      pictures.push([path, createHash('sha256').update(archive[path]).digest('hex')]);
    }
  }
  return { elements: JSON.stringify(elements), pictures: JSON.stringify(pictures) };
}

module.exports = (Core, original, saved, bytes) => {
  const sourcePages = readPages(original), savedPages = readPages(saved);
  const missingPaint = savedPages.filter(empty);
  const { rows: owned, unmapped: savedUnmapped } = paragraphPositions(saved);
  const { rows: sourceOwners, unmapped: sourceUnmapped } = paragraphPositions(original);
  const candidates = new Map();
  const authoredBreaks = sourceOwners.filter(row => row.pageBreakBefore);
  const boundaries = missingPaint.map(site => {
    const before = owned.filter(row => row.page === site.index - 1).slice(-2);
    const onSheet = owned.filter(row => row.page === site.index);
    const after = owned.filter(row => row.page === site.index + 1).slice(0, 2);
    if (!authoredBreaks.length && after[0]?.pageBreakBefore) {
      const row = after[0];
      candidates.set(row.section + ':' + row.paragraph, row);
    }
    for (const row of onSheet.filter(row => row.pageBreakBefore && !row.visibleText && row.controlCount === 0))
      candidates.set(row.section + ':' + row.paragraph, row);
    return { previous: savedPages[site.index - 1], empty: site, next: savedPages[site.index + 1], before, onSheet, after };
  });
  const base = logicalBody(saved), payload = packagePayload(saved);
  const selected = [...candidates.values()];
  const trials = [];
  const groups = [[], ...selected.map(row => [row])];
  if (selected.length > 1) groups.push(selected);
  for (const rows of groups) {
    const doc = new Core(bytes);
    try {
      doc.beginBatch();
      const changed = [];
      try {
        for (const row of rows) {
          const result = JSON.parse(doc.applyParaFormat(row.section, row.paragraph, JSON.stringify({ pageBreakBefore: false })));
          if (result.ok) changed.push([row.section, row.paragraph]);
        }
      } finally { doc.endBatch(); }
      const pages = readPages(doc);
      const sampledPayload = packagePayload(doc);
      trials.push({ changed, pages: pages.length, emptyPages: pages.filter(empty).map(page => page.index + 1),
        sameBodyAndControlSlots: logicalBody(doc) === base, sameStoryContentAndControls: sampledPayload.elements === payload.elements,
        samePictureParts: sampledPayload.pictures === payload.pictures,
        totalVisibleRuns: pages.reduce((count, page) => count + page.bodyRuns + page.cellRuns, 0),
      });
    } finally { doc.free(); }
  }
  return {
    sourcePages: sourcePages.length, savedPages: savedPages.length,
    sourceEmptyPages: sourcePages.filter(empty).map(page => page.index + 1),
    savedUnmapped, sourceUnmapped,
    sourceExplicitBodyBreaks: authoredBreaks.map(row => [row.section, row.paragraph, row.page + 1, row.visibleText]),
    totalVisibleRuns: savedPages.reduce((count, page) => count + page.bodyRuns + page.cellRuns, 0),
    boundaries, trials,
  };
};

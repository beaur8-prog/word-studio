'use strict';

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
  const rows = [];
  for (let section = 0; section < document.getSectionCount(); section++) {
    for (let paragraph = 0; paragraph < document.getParagraphCount(section); paragraph++) {
      const location = JSON.parse(document.getPageOfPosition(section, paragraph));
      if (!location.ok) continue;
      const { stored, controlCount, lines } = JSON.parse(document.getParagraphLineMetrics(section, paragraph, '[]'));
      const { pageBreakBefore, keepWithNext, keepLines, spacingBefore, spacingAfter } =
        JSON.parse(document.getParaPropertiesAt(section, paragraph));
      rows.push({
        section, paragraph, page: location.page,
        units: document.getParagraphLength(section, paragraph),
        visibleText: !!document.getTextRange(section, paragraph, 0, 0xffffffff).trim(),
        controlCount, pageBreakBefore, keepWithNext, keepLines, spacingBefore, spacingAfter, stored,
        lines: lines.map(line => [line.verticalPosition, line.height, line.spacing]),
      });
    }
  }
  return rows;
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

module.exports = (Core, original, saved, bytes) => {
  const sourcePages = readPages(original), savedPages = readPages(saved);
  const missingPaint = savedPages.filter(empty);
  const owned = paragraphPositions(saved);
  const candidates = new Map();
  const boundaries = missingPaint.map(site => {
    const before = owned.filter(row => row.page === site.index - 1).slice(-2);
    const onSheet = owned.filter(row => row.page === site.index);
    const after = owned.filter(row => row.page === site.index + 1).slice(0, 2);
    for (const row of [...onSheet, ...after].filter(row => row.pageBreakBefore && !row.visibleText && row.controlCount === 0))
      candidates.set(row.section + ':' + row.paragraph, row);
    return { previous: savedPages[site.index - 1], empty: site, next: savedPages[site.index + 1], before, onSheet, after };
  });
  const base = logicalBody(saved);
  const selected = [...candidates.values()];
  const trials = [];
  for (const rows of selected.length ? [...selected.map(row => [row]), ...(selected.length > 1 ? [selected] : [])] : []) {
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
      trials.push({ changed, pages: pages.length, emptyPages: pages.filter(empty).map(page => page.index + 1),
        sameBodyAndControlSlots: logicalBody(doc) === base,
        totalVisibleRuns: pages.reduce((count, page) => count + page.bodyRuns + page.cellRuns, 0),
      });
    } finally { doc.free(); }
  }
  return {
    sourcePages: sourcePages.length, savedPages: savedPages.length,
    sourceEmptyPages: sourcePages.filter(empty).map(page => page.index + 1),
    sourceExplicitBodyBreaks: paragraphPositions(original).filter(row => row.pageBreakBefore).map(row => [row.section, row.paragraph, row.page + 1, row.visibleText]),
    totalVisibleRuns: savedPages.reduce((count, page) => count + page.bodyRuns + page.cellRuns, 0),
    boundaries, trials,
  };
};

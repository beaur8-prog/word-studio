import { expect, test } from '@playwright/test';
import { hwpWrapFingerprint, parseHwpSourceWrap } from '../src/hwp-text-advances';

test('the complete HWP row profile preserves the prior face proof', () => {
  const face = { size: 15, pitch: 20, rowHeight: 110, hanging: 24, continued: true };
  const prior = { fingerprint: hwpWrapFingerprint('A literal for the measured paragraph tail'), ends: [28, 30],
    cellWidth: { outer: 320, text: 305 }, authoredCellFace: face };
  const scoped = { ...prior, expandedRowHeight: 110, expandedRowLine: { hanging: 24, pitch: 20, scale: 95 } };
  expect(parseHwpSourceWrap(JSON.stringify(prior))).toEqual(prior);
  expect(parseHwpSourceWrap(JSON.stringify(scoped))).toEqual(scoped);
  for (const incompatible of [
    { ...scoped, expandedRowHeight: 109 },
    { ...scoped, expandedRowLine: { ...scoped.expandedRowLine, hanging: 25 } },
    { ...scoped, expandedRowLine: { ...scoped.expandedRowLine, scale: 100 } },
  ]) expect(parseHwpSourceWrap(JSON.stringify(incompatible))).toBeUndefined();
});

test('measured cell widths preserve the original hard break and reject a mismatched proof', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { hwpWrapFingerprint, normalizeHwpWrapText } = await import('/src/hwp-text-advances.ts');
    const { clearHwpNativeCellWidthProjection, layoutHwpCellWidthFallbacks } = await import('/src/hwp-source-cell-lines.ts');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
    const first = '자연스럽게 이어지는 표의 문장입니다';
    const second = '둘째 줄';
    const size = 14.6667, pitch = size * 1.3;
    const family = '"KoPub돋움체 Light", "Noto Sans KR"';
    await document.fonts.load(`${size}px "Noto Sans KR"`, first + second);
    const host = document.createElement('div');
    host.className = 'tiptap';
    host.style.cssText = 'position:absolute;left:24px;top:24px;width:700px;';
    document.body.append(host);
    try {
      const measure = document.createElement('span');
      measure.style.cssText = `font-family:${family};font-size:${size}px;white-space:pre;`;
      measure.textContent = first;
      host.append(measure);
      const width = Math.floor(measure.getBoundingClientRect().width * .975);
      let intactHeight = 0;
      const modes = ['original', 'edited text', 'different column', 'intact', 'edited row height', 'row permits splitting', 'native', 'native intact'];
      return modes.map(mode => {
        const intact = ['intact', 'edited row height', 'row permits splitting', 'native intact'].includes(mode);
        const native = mode.startsWith('native');
        const columns = intact ? [Math.floor(width / 2), Math.ceil(width / 2)] : [width];
        const grid = columns.map(value => '<col style="width:' + value + 'px">').join('');
        host.innerHTML = `<table data-hwp-centered-borders="continuation" style="width:${width}px;table-layout:fixed;border-collapse:collapse">
          <colgroup>${grid}</colgroup><tbody><tr>
          <td colspan="${columns.length}" style="border:0;padding:0;vertical-align:top">
          <p style="border:0;padding:0;margin:0;text-indent:0px;text-align:justify;line-height:${pitch}px">
          </p></td></tr></tbody></table>`;
        const table = host.querySelector<HTMLTableElement>('table')!;
        const paragraph = table.querySelector<HTMLParagraphElement>('p')!;
        const [row] = table.rows;
        if (mode === 'edited row height') row.style.height = (intactHeight + 5) + 'px';
        if (mode === 'row permits splitting') row.dataset.rowSplit = 'allow';
        const text = document.createElement('span');
        text.style.cssText = `font-family:${family};font-size:${size}px;--word-character-scale:95`;
        text.append(mode === 'edited text' ? first + '다' : first, document.createElement('br'), second);
        paragraph.replaceChildren(text);
        const literal = first + '\n' + second;
        const authoredWidth = mode === 'different column' ? width + 8 : width;
        paragraph.dataset.hwpSourceWrap = JSON.stringify({
          fingerprint: hwpWrapFingerprint(literal), ends: [normalizeHwpWrapText(first).length, normalizeHwpWrapText(literal).length],
          cellWidth: { outer: authoredWidth, text: authoredWidth },
          cellWidthFallback: { scale: 95, hanging: 0, size, pitch, hardBreak: true,
            ...(intact && { intactRow: { height: intactHeight, columns: columns.length } }) },
        });
        const before = paragraphLineOffsets(paragraph).length + 1;
        const originalFace = text.style.fontFamily;
        layoutHwpCellWidthFallbacks(table, native);
        if (mode === 'original') intactHeight = row.getBoundingClientRect().height;
        const measured = { before, after: paragraphLineOffsets(paragraph).length + 1,
          text: paragraph.textContent, breaks: paragraph.querySelectorAll('br').length, scale: paragraph.style.transform };
        if (!native) return measured;
        const copy = host.cloneNode(true) as HTMLElement;
        clearHwpNativeCellWidthProjection(copy);
        const restored = copy.querySelector('p')!;
        return { ...measured, restored: { transform: restored.style.transform, width: restored.style.width,
          indent: restored.style.textIndent, breaks: restored.querySelectorAll('br').length,
          softBreaks: restored.querySelectorAll('wbr').length, sameFace: restored.querySelector('span')!.style.fontFamily === originalFace } };
      });
    } finally { host.remove(); }
  });
  expect(result[3]).toEqual(result[0]);
  expect(result[0]).toEqual({ before:3, after:2, text:'자연스럽게 이어지는 표의 문장입니다둘째 줄', breaks:1, scale:'scaleX(0.95)' });
  for (const rejected of [result[1],result[2],result[4],result[5]]) {
    expect(rejected.after).toBe(rejected.before);
    expect(rejected.breaks).toBe(1);
    expect(rejected.scale).toBe('');
  }
  const restored = { transform: '', width: '', indent: '0px', breaks: 1, softBreaks: 0, sameFace: true };
  expect(result[6]).toEqual({ ...result[0], restored });
  expect(result[7]).toEqual({ ...result[0], after: result[7].before, scale: '', restored });
});

test('a rejected soft line-end trial rolls back to its original intact paragraph', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { layoutHwpCellWidthFallbacks } = await import('/src/hwp-source-cell-lines.ts');
    const { hwpWrapFingerprint, normalizeHwpWrapText } = await import('/src/hwp-text-advances.ts');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
    const size = 14.6667, pitch = 20.5333, width = 90;
    const literal = '표의 첫 단락에서 원래 공백을 유지합니다. 이어지는 문장은 한 번 더 검증합니다.';
    const root = document.createElement('div');
    root.className = 'tiptap';
    root.innerHTML = '<table data-hwp-centered-borders="continuation"><colgroup><col></colgroup><tbody><tr><td><p><span></span></p></td></tr></tbody></table>';
    const table = root.querySelector<HTMLTableElement>('table')!;
    table.style.cssText = `width:${width}px;table-layout:fixed;border-collapse:collapse`;
    table.querySelector('col')!.style.width = width + 'px';
    const cell = table.rows[0].cells[0], p = cell.querySelector<HTMLParagraphElement>('p')!;
    cell.style.cssText = 'padding:0;border:0;vertical-align:top';
    p.style.cssText = `margin:0;padding:0;text-align:justify;line-height:${pitch}px;overflow-wrap:anywhere`;
    p.dataset.hwpCellLine = `1 Percent ${size} ${pitch} ${size} ${pitch}`;
    const run = p.querySelector('span')!;
    run.style.cssText = `font-family:"KoPub돋움체 Light", "Noto Sans KR";font-size:${size}px;--word-character-scale:95`;
    run.textContent = literal;
    p.dataset.hwpSourceWrap = JSON.stringify({
      fingerprint: hwpWrapFingerprint(literal), ends: [19, normalizeHwpWrapText(literal).length],
      cellWidth: { outer: width, text: width },
      cellWidthFallback: { scale: 95, hanging: 0, size, pitch, intactRow: { height: pitch * 2, columns: 1 } },
    });
    document.body.append(root);
    try {
      const lines = paragraphLineOffsets(p).length + 1;
      const before = p.outerHTML;
      layoutHwpCellWidthFallbacks(table);
      return { lines, original: cell.firstElementChild === p, sameMarkup: p.outerHTML === before,
        text: cell.textContent, hardBreaks: cell.querySelectorAll('br').length };
    } finally { root.remove(); }
  });
  expect(result.lines).toBeGreaterThan(2);
  expect(result).toMatchObject({ original: true, sameMarkup: true, hardBreaks: 0,
    text: '표의 첫 단락에서 원래 공백을 유지합니다. 이어지는 문장은 한 번 더 검증합니다.' });
});

test('a complete row keeps its whole content width separate from an existing hanging measurement', () => {
  const proof = {
    fingerprint: hwpWrapFingerprint('Original measured cell paragraph'),
    ends: [19, 30],
    cellWidth: { outer: 335, text: 291 },
    cellWidthFallback: { scale: 95, hanging: 19, size: 14.6667, pitch: 20.5,
      intactRow: { height: 50.3, columns: 1, textWidth: 311.25 } },
  };
  const previous = { ...proof, cellWidthFallback: {
    ...proof.cellWidthFallback, intactRow: { height: 50.3, columns: 1 },
  } };
  expect(parseHwpSourceWrap(JSON.stringify(proof))).toEqual(proof);
  expect(parseHwpSourceWrap(JSON.stringify(previous))).toEqual(previous);
  for (const invalid of [0, 19, 336, null, '311.25']) {
    const altered = { ...proof, cellWidthFallback: { ...proof.cellWidthFallback, intactRow: {
      ...proof.cellWidthFallback.intactRow, textWidth: invalid,
    } } };
    expect(parseHwpSourceWrap(JSON.stringify(altered))).toBeUndefined();
  }
});

test('a long cell paragraph requires evidence for both siblings in the same intact row', () => {
  const proof = {
    fingerprint: hwpWrapFingerprint('Paragraph with five independently observed source line ends'),
    ends: [8, 16, 24, 32, 48],
    cellWidth: { outer: 200, text: 174 },
    cellWidthFallback: { scale: 95, hanging: 12, size: 14, pitch: 19,
      intactRow: { height: 168, columns: 1, textWidth: 186, paragraphs: 2 } },
  };
  const { intactRow, ...fallback } = proof.cellWidthFallback;
  const legacy = { ...intactRow, paragraphs: undefined };
  expect(parseHwpSourceWrap(JSON.stringify(proof))).toEqual(proof);
  for (const altered of [legacy, { ...intactRow, paragraphs: 3 }, { ...intactRow, height: 80 },
    { ...intactRow, textWidth: undefined }]) {
    expect(parseHwpSourceWrap(JSON.stringify({ ...proof, cellWidthFallback: { ...fallback, intactRow: altered } })))
      .toBeUndefined();
  }
});

test('a two-character repeated HWP header has its own bounded source proof', () => {
  const proof = {
    fingerprint: hwpWrapFingerprint('표제'), ends: [2],
    cellWidth: { outer: 30.2, text: 30.2 }, singleHeaderLine: true,
  };
  expect(parseHwpSourceWrap(JSON.stringify(proof))).toEqual(proof);
  expect(parseHwpSourceWrap(JSON.stringify({ ...proof, ends: [3] }))).toBeUndefined();
  expect(parseHwpSourceWrap(JSON.stringify({ ...proof, singleHeaderLine: undefined, singleCellLine: true }))).toBeUndefined();
});

test('repeated headings and authored hard breaks retain their proofs through two DOCX saves', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { hwpWrapFingerprint, normalizeHwpWrapText, parseHwpSourceWrap } = await import('/src/hwp-text-advances.ts');
    const root = document.createElement('div');
    root.innerHTML = '<table><tbody><tr><th><p>표제</p></th></tr><tr><td><p>원래 첫 줄<br>원래 둘째 줄</p></td></tr></tbody></table>';
    const heading = root.querySelector<HTMLParagraphElement>('th p')!;
    const body = root.querySelector<HTMLParagraphElement>('td p')!;
    const first = body.firstChild!.textContent!;
    const text = first + '\n' + body.lastChild!.textContent!;
    const headerProof = { fingerprint: hwpWrapFingerprint('표제'), ends: [2],
      cellWidth: { outer: 30.2, text: 30.2 }, singleHeaderLine: true };
    const bodyProof = { fingerprint: hwpWrapFingerprint(text),
      ends: [normalizeHwpWrapText(first).length, normalizeHwpWrapText(text).length],
      cellWidth: { outer: 180, text: 165 },
      cellWidthFallback: { scale: 95, hanging: 12, size: 14.6667, pitch: 19.0667, hardBreak: true } };
    heading.dataset.hwpSourceWrap = JSON.stringify(headerProof);
    body.dataset.hwpSourceWrap = JSON.stringify(bodyProof);
    let doc = importEditableHwp({ html: root.innerHTML, warnings: [],
      page: { size:'A4', orientation:'portrait', margin:20, header:'', footer:'', pageNumbers:false } }, 'Measured proof.hwp');
    const stages = [];
    for (let pass = 0; pass < 2; pass++) {
      const bytes = await exportDocx(doc);
      const loaded = await importFile(new File([bytes], 'Measured proof.docx'));
      root.innerHTML = loaded.html;
      const savedHead = root.querySelector<HTMLElement>('th p')!;
      const savedBody = root.querySelector<HTMLElement>('td p')!;
      stages.push({
        header: parseHwpSourceWrap(savedHead?.dataset.hwpSourceWrap),
        body: parseHwpSourceWrap(savedBody?.dataset.hwpSourceWrap),
        savedText: [savedHead?.textContent,savedBody?.textContent],
        hardBreaks: savedBody?.querySelectorAll('br').length,
      });
      doc = { ...loaded.document, content: parseDocumentHtml(loaded.html) };
    }
    return { stages, headerProof, bodyProof };
  });
  for (const stage of result.stages) {
    expect(stage).toEqual({ header:result.headerProof, body:result.bodyProof,
      savedText:['표제','원래 첫 줄원래 둘째 줄'], hardBreaks:1 });
  }
});

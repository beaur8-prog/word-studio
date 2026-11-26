import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { Document, Packer, Paragraph, Table, TableCell, TableRow, TextRun, WidthType } from 'docx';

declare global {
  interface Window {
    hwpProbeInterrupted?: Promise<number>;
    resumeHwpProbe?: () => void;
  }
}

const content = '한글 문자 공백과 글꼴';
const source = `<p data-hwp-source-spaces="true" style='font-family:"Malgun Gothic";font-size:14pt;line-height:28px'>${content}</p>
<p style='font-family:"Malgun Gothic";font-size:14pt;line-height:28px'>${content}</p>`;

test('HWP source advances keep aligned editor and print measurements through typing and two DOCX saves', async ({ page }, info) => {
  await page.goto('/');
  const original = await page.evaluate(async () => {
    try {
      await new FontFace('Native HWP font test', 'local("Malgun Gothic"),local("MalgunGothic"),local("맑은 고딕")').load();
      return true;
    } catch { return false; }
  });
  test.skip(original, 'An installed source face does not need correction');
  let bytes = Buffer.from(source);
  for (let stage = 0; stage < 3; stage++) {
    const extension = stage ? 'docx' : 'html';
    await page.locator('input[type=file]').first().setInputFiles({
      name: `hwp-fixed-space-${stage}.${extension}`,
      mimeType: stage ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' : 'text/html',
      buffer: bytes,
    });
    await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(`hwp-fixed-space-${stage}`);
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
    const marked = editor.locator('p[data-hwp-source-spaces="true"]');
    await expect(marked).toHaveText(content);
    await expect(marked.locator('[data-hwp-source-advance-run]').first()).toBeVisible();
    const widths = async (locator: typeof marked) => locator.evaluateAll(items => items.map(p => {
      const range = document.createRange(); range.selectNodeContents(p);
      return { width: range.getBoundingClientRect().width, height: p.getBoundingClientRect().height };
    }));
    const painted = await widths(editor.locator('p'));
    expect(painted).toHaveLength(2);
    expect(painted[0].width).toBeGreaterThan(painted[1].width);
    expect(painted[0].height).toBe(painted[1].height);
    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const output = page.locator('.print-layout.preview .print-content p');
    await expect(output).toHaveCount(2);
    const printed = await widths(output);
    expect(printed[0].width).toBeCloseTo(painted[0].width, 1);
    expect(printed[1].width).toBeCloseTo(painted[1].width, 1);
    await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
    if (stage === 0) {
      await marked.click();
      await page.keyboard.press('End');
      await page.keyboard.insertText(' 입력');
      await expect(marked).toHaveText(content + ' 입력');
      await page.keyboard.press('ControlOrMeta+z');
      await expect(marked).toHaveText(content);
    }
    if (stage === 2) break;
    await page.getByRole('button', { name: '파일', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
    const saved = info.outputPath(`textbox-save-hwp-source-advances-${stage+1}.docx`);
    await (await pending).saveAs(saved);
    bytes = await readFile(saved);
    const documentXml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
    expect(documentXml.match(/wsa:fixed="1"/g)).toHaveLength(1);
    expect(documentXml).not.toContain('data-hwp-source-advance-run');
  }
});

test('HWP source advances finish across multiple batches after a document replacement', async ({ page }) => {
  await page.goto('/');
  const installed = await page.evaluate(async () => {
    try {
      await new FontFace('HWP batch test native face', 'local("한양중고딕")').load();
      return true;
    } catch { return false; }
  });
  test.skip(installed, 'An installed source face does not need correction');
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  const title = page.getByRole('textbox', { name: '문서 제목', exact: true });
  const upload = page.locator('input[type=file]').first();
  const paragraph = (text: string) =>
    `<p data-hwp-source-spaces="true" style='font-family:"한양중고딕";font-size:14pt;line-height:28px'>${text}</p>`;
  await upload.setInputFiles({
    name: 'hwp-batch-old.html', mimeType: 'text/html',
    buffer: Buffer.from(paragraph('기존 문단').repeat(72)),
  });
  await expect(title).toHaveValue('hwp-batch-old');
  await upload.setInputFiles({
    name: 'hwp-batch-new.html', mimeType: 'text/html',
    buffer: Buffer.from(paragraph(content).repeat(70)),
  });
  await expect(title).toHaveValue('hwp-batch-new');
  const paragraphs = editor.locator('p[data-hwp-source-spaces="true"]');
  await expect(paragraphs).toHaveText(Array(70).fill(content));
  await expect(paragraphs.filter({ has: page.locator('[data-hwp-source-advance-run]') })).toHaveCount(70);
});

test('a print measurement interrupts an active HWP editor probe and the editor restarts afterwards', async ({ page }) => {
  await page.goto('/');
  const installed = await page.evaluate(async () => {
    try {
      await new FontFace('HWP probe test native face', 'local("한양중고딕")').load();
      return true;
    } catch { return false; }
  });
  test.skip(installed, 'An installed source face does not need correction');
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  await editor.waitFor();
  await page.evaluate(async () => {
    const { suspendEditorReflow } = await import('/src/editor-layout-task.ts');
    window.hwpProbeInterrupted = new Promise<number>(resolve => {
      const observed = new MutationObserver(records => {
        for (const record of records) for (const node of record.addedNodes) {
          if (!(node instanceof HTMLElement) || node.getAttribute('aria-hidden') !== 'true' ||
              node.style.left !== '-100000px' || !node.querySelector('p[data-hwp-source-spaces]')) continue;
          observed.disconnect();
          const resume = suspendEditorReflow();
          const original = Element.prototype.getBoundingClientRect;
          let measurements = 0;
          Element.prototype.getBoundingClientRect = function () {
            if (node.contains(this)) measurements++;
            return original.call(this);
          };
          const removed = new MutationObserver(() => {
            if (node.isConnected) return;
            removed.disconnect();
            Element.prototype.getBoundingClientRect = original;
            window.resumeHwpProbe = resume;
            resolve(measurements);
          });
          removed.observe(document.body, { childList: true });
          return;
        }
      });
      observed.observe(document.body, { childList: true });
    });
  });
  const paragraph = `<p data-hwp-source-spaces="true" style='font-family:"한양중고딕";font-size:14pt;line-height:28px'>${content}</p>`;
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'hwp-probe-interrupted.html', mimeType: 'text/html', buffer: Buffer.from(paragraph.repeat(70)),
  });
  await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue('hwp-probe-interrupted');
  expect(await page.evaluate(() => window.hwpProbeInterrupted)).toBe(0);
  await expect(editor.locator('[data-hwp-source-advance-run]')).toHaveCount(0);
  await page.evaluate(() => window.resumeHwpProbe?.());
  const paragraphs = editor.locator('p[data-hwp-source-spaces="true"]');
  await expect(paragraphs).toHaveText(Array(70).fill(content));
  await expect(paragraphs.filter({ has: page.locator('[data-hwp-source-advance-run]') })).toHaveCount(70);
});

test('a cell continuation keeps a verified line above the cut when only its trailing space extends past it', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { tableFragmenter } = await import('/src/table-fragments.ts');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
    const { hwpWrapFingerprint, normalizeHwpWrapText } = await import('/src/hwp-text-advances.ts');
    const text = 'Alpha bravo charlie delta echo foxtrot golf hotel';
    function measure(stale: boolean) {
      const host = document.createElement('div');
      host.className = 'tiptap';
      host.style.cssText = 'position:absolute;left:0;top:0;padding:0;width:140px';
      host.innerHTML = `<table style="width:140px;table-layout:fixed;border-collapse:collapse">
        <tbody><tr data-row-split="allow"><td style="padding:2px;border:0;vertical-align:top">
        <p data-exact-line-height="true" style="font:16px/28px Arial;margin:0">${text}</p>
        <p style="font:16px/28px Arial;margin:0">Continuation paragraph</p>
        </td></tr></tbody></table>`;
      document.body.append(host);
      try {
        const row = host.querySelector('tr')!, cell = row.cells[0], first = cell.querySelector('p')!;
        const offsets = paragraphLineOffsets(first);
        const ends = [...offsets, text.length].map(offset => normalizeHwpWrapText(text.slice(0, offset)).length);
        first.dataset.hwpSourceWrap = JSON.stringify({ fingerprint: hwpWrapFingerprint(stale ? text + ' edited' : text), ends });
        const range = document.createRange(); range.selectNodeContents(first);
        const inkBottom = Math.max(...[...range.getClientRects()].filter(rect => rect.width > 0).map(rect => rect.bottom));
        const full = first.getBoundingClientRect().bottom;
        const css = getComputedStyle(cell);
        const cut = (full + inkBottom) / 2 + parseFloat(css.paddingBottom) + parseFloat(css.borderBottomWidth);
        const fragments = tableFragmenter({ bodyHeight: () => 300, hasDeferredNote: () => false, deferredNotes: () => new Set() });
        const tail = fragments.splitElement(cell, cut);
        return { hasLines: ends.length > 1, unpainted: full - inkBottom, head: row.querySelector('p')?.textContent,
          tail: tail?.textContent?.trim(), margin: row.querySelector('p')?.style.marginBottom };
      } finally { host.remove(); }
    }
    return { unchanged: measure(false), edited: measure(true) };
  });
  expect(result.unchanged.hasLines).toBe(true);
  expect(result.unchanged.unpainted).toBeGreaterThan(1);
  expect(result.unchanged.head).toBe('Alpha bravo charlie delta echo foxtrot golf hotel');
  expect(result.unchanged.tail).toBe('Continuation paragraph');
  expect(parseFloat(result.unchanged.margin!)).toBeLessThan(0);
  expect(result.edited.head).not.toBe(result.unchanged.head);
  expect(result.edited.tail).toContain('Continuation paragraph');
});

test('source HWP cell lines keep styled text and its geometry with view-only wrap opportunities', async ({ page }) => {
  await page.goto('/');
  const measured = await page.evaluate(async () => {
    const { layoutHwpSourceCellLines } = await import('/src/hwp-source-cell-lines.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
    const text = 'abcdefghijklmnopqrstuv';
    const host = document.createElement('div');
    host.className = 'tiptap';
    host.style.cssText = 'position:absolute;top:0;left:0;padding:0;width:500px';
    document.body.append(host);
    try {
      function measure(mode: 'valid' | 'edited' | 'resized') {
        const painted = mode === 'edited' ? text.replace('m', 'x') : text;
        host.innerHTML = `<div class="print-page"><div class="print-content">
          <table style="width:84px;table-layout:fixed;border-collapse:collapse"><tbody><tr><td style="padding:0;border:0">
          <p data-hwp-source-spaces="true" style="margin:0;line-height:28px;overflow-wrap:anywhere">
            <span style="font:700 16px/28px Courier New,monospace;color:rgb(4,10,20)">${painted}</span></p>
          </td></tr></tbody></table></div></div>`;
        const paragraph = host.querySelector('p')!, cell = host.querySelector('td')!;
        paragraph.firstChild?.remove(); // Keep only the authored run in the fixture.
        const geometry = () => [paragraph, cell, cell.parentElement!].map(el => {
          const bounds = el.getBoundingClientRect();
          return [bounds.x, bounds.y, bounds.width, bounds.height];
        });
        const before = geometry();
        paragraph.dataset.hwpSourceWrap = JSON.stringify({ fingerprint: hwpWrapFingerprint(text), ends: [7, 14, 22],
          cellWidth: { outer: before[1][2] + (mode === 'resized' ? 8 : 0), text: before[0][2] } });
        const oldEnds = paragraphLineOffsets(paragraph);
        layoutHwpSourceCellLines(host);
        return { before, after: geometry(), oldEnds, ends: paragraphLineOffsets(paragraph), content: paragraph.textContent,
          opportunities: paragraph.querySelectorAll('wbr').length,
          styles: [...paragraph.querySelectorAll('span[style*="color"]')].map(el => getComputedStyle(el).fontWeight) };
      }
      return { unchanged: measure('valid'), edited: measure('edited'), resized: measure('resized') };
    } finally { host.remove(); }
  });
  expect(measured.unchanged.oldEnds).not.toEqual([7, 14]);
  expect(measured.unchanged.ends).toEqual([7, 14]);
  expect(measured.unchanged.content).toBe('abcdefghijklmnopqrstuv');
  expect(measured.unchanged.opportunities).toBe(2);
  expect(measured.unchanged.styles).toEqual(['700', '700', '700']);
  expect(measured.unchanged.after).toEqual(measured.unchanged.before);
  expect(measured.edited.opportunities).toBe(0);
  expect(measured.resized.opportunities).toBe(0);
});

test('source HWP terminal cells keep their compact last line when the recorded wraps move', async ({ page }) => {
  await page.goto('/');
  const measured = await page.evaluate(async () => {
    const { layoutHwpCellLines } = await import('/src/hwp-cell-line-metrics.ts');
    const { layoutHwpSourceCellLines } = await import('/src/hwp-source-cell-lines.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
    const text = 'abcdefghijklmnopqrstuv';
    const host = document.createElement('div');
    host.className = 'tiptap';
    host.style.cssText = 'position:absolute;top:0;left:0;padding:0;width:500px';
    document.body.append(host);
    try {
      return ['unchanged', 'edited', 'resized', 'continued'].map(mode => {
        const painted = mode === 'edited' ? text.replace('m', 'x') : text;
        host.innerHTML = `<div class="print-page"><div class="print-content">
          <table style="width:84px;table-layout:fixed;border-collapse:collapse"><tbody><tr><td style="padding:0;border:0">
          <p data-hwp-source-spaces="true" data-hwp-cell-line="1 Fixed 16 28 16 28"
            style="margin:0;font:16px/28px 'Courier New',monospace;overflow-wrap:anywhere"><span style="color:rgb(4, 10, 20)">${painted}</span></p>
          </td></tr></tbody></table></div></div>`;
        const paragraph = host.querySelector('p')!, cell = host.querySelector('td')!;
        layoutHwpCellLines(host);
        if (mode === 'continued') paragraph.dataset.hwpCellContinuation = 'true';
        const geometry = () => [paragraph, cell, cell.parentElement!, cell.closest('table')!].map(el => {
          const { x, y, width, height } = el.getBoundingClientRect();
          return [x, y, width, height];
        });
        const snapshot = () => ({ geometry: geometry(), ends: paragraphLineOffsets(paragraph),
          prefix: paragraph.querySelector('[data-hwp-cell-prefix]')?.textContent,
          lastHeight: paragraph.style.lineHeight, text: paragraph.textContent,
          opportunities: paragraph.querySelectorAll('wbr').length });
        const before = snapshot();
        paragraph.dataset.hwpSourceWrap = JSON.stringify({ fingerprint: hwpWrapFingerprint(text), ends: [7, 14, 22],
          cellWidth: { outer: before.geometry[1][2], text: before.geometry[0][2] + (mode === 'resized' ? 8 : 0) } });
        layoutHwpSourceCellLines(host);
        return { mode, before, after: snapshot() };
      });
    } finally { host.remove(); }
  });
  const [unchanged, ...controls] = measured;
  expect(unchanged.before.ends).toEqual([8, 16]);
  expect(unchanged.before.prefix).toBe('abcdefghijklmnop');
  expect(unchanged.after.ends).toEqual([7, 14]);
  expect(unchanged.after.prefix).toBe('abcdefghijklmn');
  expect(unchanged.after.lastHeight).toBe('16px');
  expect(unchanged.after.opportunities).toBe(2);
  expect(unchanged.after.geometry).toEqual(unchanged.before.geometry);
  expect(unchanged.after.text).toBe(unchanged.before.text);
  for (const result of controls) expect(result.after, result.mode).toEqual(result.before);
});

test('a recorded HWP body origin tolerates only layout rounding in the live editor', async ({ page }) => {
  await page.goto('/');
  const allowance = await page.evaluate(async () => {
    const { sourceFlowAllowance } = await import('/src/hwp-text-advance-layout.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const host = document.createElement('div');
    host.className = 'tiptap';
    host.style.cssText = 'position:absolute;left:0;top:0;width:400px;padding:0';
    host.innerHTML = '<p data-hwp-source-spaces="true" data-hwp-body-line="2 170 17.3333 29.5" ' +
      'style="position:relative;top:30px;margin:0 0 0 20px;text-indent:-8px;line-height:29.5px">원본 조판에 기록한 문장</p>';
    document.body.append(host);
    try {
      const p = host.querySelector('p')!, dimensions = p.getBoundingClientRect();
      const paper = { index: 2, box: new DOMRect(0, 0, 400, 650), bottom: 600, scale: 1, liveEditor: true };
      const record = { fingerprint: hwpWrapFingerprint(p.textContent!), ends: [4, 10] };
      function measure(y: number, x = 0, index = 2, liveEditor = true) {
        p.dataset.hwpSourceWrap = JSON.stringify({ ...record, origin: { page: 2, x: dimensions.left - 8 + x, y: dimensions.top + y } });
        return sourceFlowAllowance(p, { ...paper, index, liveEditor }) ?? null;
      }
      const result = { available: 600 - dimensions.top, sameOrigin: measure(.13), shiftedY: measure(.3),
        shiftedX: measure(.13, .13), otherPaper: measure(.13, 0, 3), print: measure(.13, 0, 2, false) };
      const prior = document.createElement('p');
      prior.style.cssText = 'position:absolute;margin:0;top:-20px;height:31.7px';
      const gutter = document.createElement('span');
      gutter.className = 'pagination-cell-gap ProseMirror-widget';
      gutter.setAttribute('aria-hidden', 'true');
      gutter.style.cssText = 'display:block;height:29.7px';
      prior.append(gutter); host.prepend(prior);
      const calibrated = { ...paper, topMargin: 10 };
      p.dataset.hwpSourceWrap = JSON.stringify({ ...record, origin: { page:2, x:dimensions.left - 8, y:dimensions.top + .3 } });
      const anchored = sourceFlowAllowance(p, calibrated) ?? null;
      const differentPaper = sourceFlowAllowance(p, { ...calibrated, index:3 }) ?? null;
      gutter.style.height = '27px';
      const wrongGutter = sourceFlowAllowance(p, calibrated) ?? null;
      p.append(' 변경');
      return { ...result, anchored, differentPaper, wrongGutter, changedText: measure(.13) };
    } finally { host.remove(); }
  });
  expect(allowance.sameOrigin).toBe(allowance.available);
  expect(allowance.shiftedY).toBeNull();
  expect(allowance.shiftedX).toBeNull();
  expect(allowance.otherPaper).toBeNull();
  expect(allowance.print).toBeNull();
  expect(allowance.anchored).toBe(allowance.available);
  expect(allowance.differentPaper).toBeNull();
  expect(allowance.wrongGutter).toBeNull();
  expect(allowance.changedText).toBeNull();
});

test('recorded HWP body wraps fit only at the authored measure without changing page geometry', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { layoutHwpSourceCellLines } = await import('/src/hwp-source-cell-lines.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
    const text = 'abcdefghijklmnopqrstuv';
    const host = document.createElement('div');
    host.className = 'tiptap';
    host.style.cssText = 'position:absolute;top:0;left:0;padding:0';
    document.body.append(host);
    try {
      return ['same', 'resized', 'fragment', 'columns'].map(mode => {
        host.innerHTML = `<div class="print-page" style="width:100px;height:200px;padding:0">
          <div class="print-content" style="width:${mode === 'resized' ? 90 : 84}px" ${mode === 'columns' ? 'data-column-count="2"' : ''}>
          <p data-hwp-source-spaces="true" style="margin:0;line-height:28px;text-align:justify;overflow-wrap:anywhere" ${mode === 'fragment' ? 'data-print-soft-end="true"' : ''}><span style="font:700 16px/28px Courier New,monospace">${text}</span></p>
          </div></div>`;
        const paragraph = host.querySelector('p')!;
        paragraph.dataset.hwpSourceWrap = JSON.stringify({ fingerprint: hwpWrapFingerprint(text), ends: [7, 14, 22], bodyWidth: 84 });
        const geometry = () => [...host.querySelectorAll('.print-page,.print-content,p')].map(element => {
          const { x, y, width, height } = element.getBoundingClientRect();
          return [x, y, width, height];
        });
        const before = geometry(), oldEnds = paragraphLineOffsets(paragraph);
        layoutHwpSourceCellLines(host);
        return { before, after: geometry(), oldEnds, ends: paragraphLineOffsets(paragraph), text: paragraph.textContent,
          opportunities: paragraph.querySelectorAll('wbr').length };
      });
    } finally { host.remove(); }
  });
  const [same, ...controls] = result;
  expect(same.oldEnds).not.toEqual([7, 14]);
  expect(same.ends).toEqual([7, 14]);
  expect(same.opportunities).toBe(2);
  expect(same.text).toBe('abcdefghijklmnopqrstuv');
  for (const control of controls) {
    expect(control.opportunities).toBe(0);
    expect(control.ends).toEqual(control.oldEnds);
  }
  for (const observation of result) expect(observation.after).toEqual(observation.before);
});

test('recorded HWP column lines require the original literal and a stable unbroken column', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { layoutHwpSourceCellLines } = await import('/src/hwp-source-cell-lines.ts');
    const { preserveHwpColumnLineEnds } = await import('/src/hwp/script-fonts.ts');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
    const text = 'abcdefghijklmnopqrstuv';
    const original = '<p style="margin:0;line-height:28px;text-align:justify;overflow-wrap:anywhere">\n' +
      '<u><span style="font:700 16px/28px Courier New,monospace">' + text + '</span></u>\n</p>';
    const makeProof = (width: number, lastY = 56) => preserveHwpColumnLineEnds(original, text, () => JSON.stringify({
      stored: true, controlCount: 0,
      lines: [0, 7, 14].map((textStart, index) => ({ textStart, width, verticalPosition: [0, 28, lastY][index] })),
    }));
    const host = document.createElement('div');
    host.className = 'tiptap';
    host.style.cssText = 'position:absolute;top:0;left:0;padding:0';
    document.body.append(host);
    try {
      const cases = ['same', 'changed width', 'edited text', 'fragment', 'single column'].map(mode => {
        host.innerHTML = '<div class="print-page" style="width:100px;height:200px;padding:0">' +
          '<div class="print-content" style="width:' + (mode === 'changed width' ? '90' : '84') + 'px"' +
          (mode === 'single column' ? '' : ' data-column-count="2"') + '>' + makeProof(84) + '</div></div>';
        const paragraph = host.querySelector('p')!;
        // The HWP importer drops outside-span indentation in bridge markup.
        for (const node of [...paragraph.childNodes]) if (node.nodeType === Node.TEXT_NODE) node.remove();
        if (mode === 'edited text') paragraph.querySelector('span')!.textContent = text.replace('a', 'X');
        if (mode === 'fragment') paragraph.dataset.printParagraphFragment = 'true';
        const geometry = () => [...host.querySelectorAll('.print-page,.print-content,p')].map(element => {
          const { x, y, width, height } = element.getBoundingClientRect();
          return [x, y, width, height];
        });
        const before = geometry(), oldEnds = paragraphLineOffsets(paragraph), oldText = paragraph.textContent;
        layoutHwpSourceCellLines(host);
        return { before, after: geometry(), oldEnds, ends: paragraphLineOffsets(paragraph), oldText,
          text: paragraph.textContent, opportunities: paragraph.querySelectorAll('wbr').length,
          wrapperDisplays: [...paragraph.querySelectorAll(':scope > span')].map(span => getComputedStyle(span).display),
          underlined: [...paragraph.querySelectorAll('u')].map(span => span.textContent).join('') };
      });
      return { cases, decreasingYIgnored: makeProof(84, 20) === original };
    } finally { host.remove(); }
  });
  const [same, ...controls] = result.cases;
  expect(same.oldEnds).not.toEqual([7, 14]);
  expect(same.ends).toEqual([7, 14]);
  expect(same.opportunities).toBe(2);
  expect(same.wrapperDisplays).toEqual(['contents', 'contents', 'contents']);
  expect(result.decreasingYIgnored).toBe(true);
  for (const candidate of result.cases) {
    expect(candidate.text).toBe(candidate.oldText);
    expect(candidate.underlined).toBe(candidate.oldText);
    expect(candidate.after).toEqual(candidate.before);
  }
  for (const control of controls) {
    expect(control.opportunities).toBe(0);
    expect(control.ends).toEqual(control.oldEnds);
  }
});

test('a continuing HWP cell can recover a missing line without changing its text or grid width', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { restoreHwpContinuedTableLines } = await import('/src/hwp-source-cell-lines.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
    const author = 'abcdefghijklmnopqrstuv';
    const host = document.createElement('div');
    host.style.cssText = 'position:absolute;top:0;left:0';
    document.body.append(host);
    try {
      return ['matching', 'unwitnessed', 'resized', 'edited', 'exact row'].map(mode => {
        const width = mode === 'resized' ? 117 : 112;
        host.innerHTML = '<table style="table-layout:fixed;border-collapse:collapse;width:' + (width + 56) +
          'px"><colgroup><col style="width:' + width + 'px"><col style="width:56px"></colgroup>' +
          '<tr><td style="padding:0"><p data-hwp-source-spaces="true" style="margin:0;overflow-wrap:anywhere;' +
          'font:700 16px/28px Courier New,monospace">' + (mode === 'edited' ? author.replace('a', 'X') : author) +
          '</p></td><td style="padding:0">Peer</td></tr></table>';
        const row = host.querySelector('tr')!;
        if (mode === 'exact row') row.dataset.wordRowRule = 'exact';
        const paragraph = host.querySelector('p')!;
        paragraph.dataset.hwpSourceWrap = JSON.stringify({ fingerprint: hwpWrapFingerprint(author), ends: [7, 14, 22],
          cellWidth: { outer: 112, text: 112 }, splitCell: mode === 'unwitnessed' ? undefined : true });
        const box = () => paragraph.getBoundingClientRect();
        const before = { ends: paragraphLineOffsets(paragraph), width: box().width, height: box().height };
        const text = row.textContent;
        restoreHwpContinuedTableLines(row, false);
        return { before, after: { ends: paragraphLineOffsets(paragraph), width: box().width, height: box().height },
          stableText: row.textContent === text, opportunities: paragraph.querySelectorAll('wbr').length, hardBreaks: paragraph.querySelectorAll('br').length };
      });
    } finally { host.remove(); }
  });
  const [matching, ...controls] = result;
  expect(matching.after.ends).toEqual([7, 14]);
  expect(matching.after.height).toBeGreaterThan(matching.before.height);
  expect(matching.opportunities).toBe(2);
  for (const candidate of result) {
    expect(candidate.stableText).toBe(true);
    expect(candidate.after.width).toBe(candidate.before.width);
    expect(candidate.hardBreaks).toBe(0);
  }
  for (const control of controls) {
    expect(control.after).toEqual(control.before);
    expect(control.opportunities).toBe(0);
  }
});

test('a continuing HWP cell keeps a measured final line when the earlier breaks already agree', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { restoreHwpContinuedTableLines } = await import('/src/hwp-source-cell-lines.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
    const text = 'ABCDEFGH IJKLMNOP QRST UV';
    const host = document.createElement('div');
    document.body.append(host);
    try {
      return ['authored', 'unknown row', 'changed line ends', 'wide cell', 'edited text', 'unbounded ink'].map(mode => {
        const width = mode === 'wide cell' ? 117 : 112;
        const padding = mode === 'unbounded ink' ? 0 : 4;
        host.innerHTML = '<table style="table-layout:fixed;border-collapse:collapse;width:' + (width + 56) +
          'px"><colgroup><col style="width:' + width + 'px"><col style="width:56px"></colgroup>' +
          '<tr><td style="padding:0 ' + padding + 'px"><p data-hwp-source-spaces="true" style="margin:0;overflow-wrap:anywhere;' +
          'font:700 16px/28px Courier New,monospace">' + (mode === 'edited text' ? text + 'Z' : text) +
          '</p></td><td style="padding:0">Peer</td></tr></table>';
        const paragraph = host.querySelector('p')!, cell = paragraph.closest('td')!, row = cell.closest('tr')!;
        const raw = paragraph.textContent;
        paragraph.dataset.hwpSourceWrap = JSON.stringify({ fingerprint: hwpWrapFingerprint(text),
          ends: mode === 'changed line ends' ? [6, 16, 20, 22] : [8, 16, 20, 22],
          cellWidth: { outer: 112, text: 112 - padding * 2 }, splitCell: mode !== 'unknown row' || undefined });
        const measure = () => ({ ends: paragraphLineOffsets(paragraph), width: cell.getBoundingClientRect().width,
          height: paragraph.getBoundingClientRect().height });
        const before = measure();
        restoreHwpContinuedTableLines(row, false);
        return { before, after: measure(), literal: paragraph.textContent === raw,
          hardBreaks: paragraph.querySelectorAll('br').length };
      });
    } finally { host.remove(); }
  });
  const [authored, ...controls] = result;
  expect(authored.before.ends).toHaveLength(2);
  expect(authored.after.ends).toEqual([9, 18, 24]);
  expect(authored.after.height).toBe(authored.before.height + 28);
  expect(authored.hardBreaks).toBe(1);
  for (const candidate of result) {
    expect(candidate.literal).toBe(true);
    expect(candidate.after.width).toBe(candidate.before.width);
  }
  for (const control of controls) expect(control.hardBreaks).toBe(0);
});

test('HWP body wraps retain an existing first-line overhang only inside the available ink space', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { layoutHwpSourceCellLines } = await import('/src/hwp-source-cell-lines.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
    const text = 'abcdefghijklmnopqrstuv';
    const host = document.createElement('div');
    host.className = 'tiptap';
    host.style.cssText = 'position:absolute;top:0;left:0;padding:0';
    document.body.append(host);
    try {
      return ['space', 'body start', 'previous ink'].map(mode => {
        const previous = mode === 'body start' ? '' :
          `<p style="margin:0;line-height:${mode === 'previous ink' ? 18 : 44}px"><span
          style="position:relative;top:${mode === 'previous ink' ? 8 : 0}px">Previous</span></p>`;
        host.innerHTML = `<div class="print-page" style="width:100px;height:250px;padding:0">
          <div class="print-content" style="width:84px">${previous}
          <p data-hwp-source-spaces="true" data-hwp-body-layout="mixed-top"
            style="margin:0;line-height:0;text-align:justify;overflow-wrap:anywhere;--hwp-body-ratio:1.75;--hwp-body-ink-offset:-8px">
            <span data-hwp-body-run style="font-family:Courier New,monospace;font-size:16px">${text}</span></p>
          </div></div>`;
        const paragraph = host.querySelector<HTMLParagraphElement>('[data-hwp-source-spaces]')!;
        for (const child of [...paragraph.childNodes]) if (child.nodeType === Node.TEXT_NODE) child.remove();
        paragraph.dataset.hwpSourceWrap = JSON.stringify({ fingerprint: hwpWrapFingerprint(text), ends: [7, 14, 22], bodyWidth: 84 });
        const geometry = () => [...host.querySelectorAll('.print-page,.print-content,p')].map(element => {
          const { x, y, width, height } = element.getBoundingClientRect();
          return [x, y, width, height];
        });
        const range = document.createRange();
        range.selectNodeContents(paragraph);
        const top = Math.min(...[...range.getClientRects()].filter(r => r.width > 0).map(r => r.top));
        const firstOverhang = paragraph.getBoundingClientRect().top - top;
        const prior = paragraph.previousElementSibling;
        if (prior) range.selectNodeContents(prior);
        if (mode === 'previous ink') {
          const witness = prior!.querySelector('span')!;
          const shift = top - range.getBoundingClientRect().bottom + 2;
          witness.style.top = (parseFloat(witness.style.top) + shift) + 'px';
        }
        const previousBottom = prior ? range.getBoundingClientRect().bottom : null;
        const before = geometry(), oldEnds = paragraphLineOffsets(paragraph);
        layoutHwpSourceCellLines(host);
        return { mode, before, after: geometry(), firstOverhang, top, previousBottom, oldEnds, ends: paragraphLineOffsets(paragraph),
          text: paragraph.textContent, opportunities: paragraph.querySelectorAll('wbr').length };
      });
    } finally { host.remove(); }
  });
  const [available, ...controls] = result;
  expect(available.firstOverhang).toBeGreaterThan(.5);
  expect(available.ends).toEqual([7, 14]);
  expect(available.opportunities).toBe(2);
  expect(controls[1].previousBottom).toBeGreaterThan(controls[1].top);
  for (const control of controls) {
    expect(control.opportunities, control.mode).toBe(0);
    expect(control.ends).toEqual(control.oldEnds);
  }
  for (const observation of result) {
    expect(observation.after).toEqual(observation.before);
    expect(observation.text).toBe('abcdefghijklmnopqrstuv');
  }
});

test('HWP body wrap ignores a fitting glyph box when the actual letter extends past the body', async ({ page }) => {
  await page.goto('/');
  const rows = await page.evaluate(async () => {
    const { layoutHwpSourceCellLines } = await import('/src/hwp-source-cell-lines.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
    const canvas = document.createElement('canvas').getContext('2d')!;
    canvas.font = 'italic 256px "Courier New", monospace';
    const text = 'abcdefghijklmnopqrstuf';
    const advance = canvas.measureText('xxxxxxxx').width / 16;
    const lastInkWidth = canvas.measureText('f').actualBoundingBoxRight / 16;
    const run = (value: string) => `<span data-hwp-body-run style="font:italic 16px 'Courier New',monospace">${value}</span>`;
    const host = document.createElement('div');
    host.className = 'tiptap';
    host.style.cssText = 'position:absolute;top:0;left:0;padding:0';
    document.body.append(host);
    try {
      return [.25, 3].map(room => {
        const width = advance + room;
        host.innerHTML = `<div class="print-page" style="width:100px;height:250px;padding:0">
          <div class="print-content" style="width:${width}px"><p style="margin:0;line-height:44px">Previous</p>
          <p data-hwp-source-spaces="true" data-hwp-body-layout="mixed-top"
            style="margin:0;line-height:0;text-align:justify;overflow-wrap:anywhere;--hwp-body-ratio:1.75;--hwp-body-ink-offset:-8px">${run(text)}</p>
          </div></div>`;
        const paragraph = host.querySelector<HTMLParagraphElement>('p[data-hwp-source-spaces]')!;
        paragraph.dataset.hwpSourceWrap = JSON.stringify({
          fingerprint: hwpWrapFingerprint(text), ends: [7, 14, 22], bodyWidth: width,
        });
        layoutHwpSourceCellLines(host);
        const ends = paragraphLineOffsets(paragraph);
        const opportunities = paragraph.querySelectorAll('wbr').length;
        paragraph.innerHTML = [text.slice(0, 7), text.slice(7, 14), text.slice(14)].map(run).join('<wbr>');
        for (const inline of paragraph.querySelectorAll('span')) inline.style.whiteSpace = 'pre';
        const last = paragraph.lastElementChild!.firstChild as Text;
        const range = document.createRange();
        range.setStart(last, last.length - 1);
        range.setEnd(last, last.length);
        const glyph = range.getBoundingClientRect(), right = paragraph.getBoundingClientRect().right;
        return { ends, opportunities, trialEnds: paragraphLineOffsets(paragraph),
          boxRoom: right - glyph.right, inkRoom: right - glyph.left - lastInkWidth };
      });
    } finally { host.remove(); }
  });
  const [outside, inside] = rows;
  expect(outside.trialEnds).toEqual([7, 14]);
  expect(outside.boxRoom).toBeGreaterThanOrEqual(0);
  expect(outside.inkRoom).toBeLessThan(0);
  expect(outside.opportunities).toBe(0);
  expect(inside.inkRoom).toBeGreaterThan(.25);
  expect(inside.ends).toEqual([7, 14]);
  expect(inside.opportunities).toBe(2);
});

for (const target of ['cell', 'body origin', 'body width'] as const) {
  test(`formatting a recorded HWP ${target} retires its visual line starts through undo and redo`, async ({ page }) => {
    await page.goto('/');
    const wrap = await page.evaluate(async target => {
      const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
      return JSON.stringify({ fingerprint: hwpWrapFingerprint('abcdefghijklmnopqrstuv'), ends: [7, 14, 22],
        ...(target === 'cell' ? { cellWidth: { outer: 84, text: 84 } } : target === 'body width' ? { bodyWidth: 84 }
          : { origin: { page: 1, x: 96, y: 96 } }) });
    }, target);
    const source = `<p data-hwp-source-spaces="true" data-hwp-source-wrap='${wrap}'>abcdefghijklmnopqrstuv</p>`;
    await page.locator('input[type=file]').first().setInputFiles({
      name: 'hwp-source-wrap-format.html', mimeType: 'text/html',
      buffer: Buffer.from(target === 'cell' ? `<table><tbody><tr><td>${source}</td></tr></tbody></table>` : source),
    });
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
    const paragraph = editor.locator('p[data-hwp-source-spaces]').first();
    await expect(paragraph).toHaveAttribute('data-hwp-source-wrap', wrap);
    const formatted = await editor.evaluate(el => {
      const instance = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
      let position = -1;
      instance.state.doc.descendants((node, at) => {
        if (node.type.name === 'paragraph' && node.attrs.hwpSourceWrap) position = at;
      });
      return position >= 0 && instance.chain().setTextSelection({ from: position + 1, to: position + 23 }).toggleBold().run();
    });
    expect(formatted).toBe(true);
    await expect(paragraph.locator('strong')).not.toHaveCount(0);
    await expect(paragraph).not.toHaveAttribute('data-hwp-source-wrap');
    await expect(paragraph).toHaveText('abcdefghijklmnopqrstuv');
    await editor.evaluate(el => (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor.commands.undo());
    await expect(paragraph).toHaveAttribute('data-hwp-source-wrap', wrap);
    await expect(paragraph.locator('strong')).toHaveCount(0);
    await editor.evaluate(el => (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor.commands.redo());
    await expect(paragraph).not.toHaveAttribute('data-hwp-source-wrap');
    await expect(paragraph.locator('wbr')).toHaveCount(0);
  });
}

test('verified hanging HWP body lines survive print-only breaks and reject a changed width', async ({ page }) => {
  await page.goto('/');
  const results = await page.evaluate(async () => {
    const { layoutHwpBodyLines, restoreHwpAuthoredBodyLines } = await import('/src/hwp-body-line-metrics.ts');
    const { encodeHwpBodyLine } = await import('/src/hwp-body-line.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const literal = 'AAAA BBBB CCCC DDDD';
    const sheet = document.createElement('div');
    sheet.className = 'print-page';
    sheet.style.cssText = 'position:fixed;top:0;left:0;width:400px;height:400px';
    sheet.dataset.page = '1';
    sheet.innerHTML = '<div class="tiptap print-content" style="position:absolute;top:10px;left:10px;width:240px;height:360px;padding:0;min-height:0;box-sizing:border-box"></div>';
    document.body.append(sheet);
    try {
      const body = sheet.querySelector<HTMLElement>('.print-content')!;
      return ['plain', 'mixed', 'width', 'late'].map(mode => {
        body.style.width = mode === 'width' ? '220px' : '240px';
        const p = document.createElement('p');
        p.style.cssText = 'font-family:monospace;font-size:16px;line-height:24px;min-height:0;margin:0 0 0 30px;text-indent:-30px;text-align:justify;position:relative;left:4px';
        p.innerHTML = '<span style="font-size:16px">AAAA BBBB </span><span style="font-size:' +
          (mode === 'mixed' ? '14px' : '16px') + '">CCCC DDDD</span>';
        if (mode === 'mixed') p.dataset.hwpBodyLine = encodeHwpBodyLine({percent:150,font:16,base:24});
        body.replaceChildren(p);
        layoutHwpBodyLines(body, true);
        const paper = sheet.getBoundingClientRect(), box = p.getBoundingClientRect();
        p.dataset.hwpSourceWrap = JSON.stringify({ fingerprint:hwpWrapFingerprint(literal), ends:[8,16],
          authoredBodyLines:{hanging:30,size:16,pitch:24,width:240,
            origin:{page:1,x:box.left-paper.left-30,y:box.top-paper.top-(mode === 'late' ? 2 : 0)}} });
        const restore = restoreHwpAuthoredBodyLines(p);
        const addedHeight = p.getBoundingClientRect().height - box.height;
        const printed = p.querySelectorAll('br').length, retainedText = p.textContent === literal;
        restore?.();
        return {mode,printed,addedHeight,retainedText,restored:p.querySelectorAll('br').length, restoredHeight:p.getBoundingClientRect().height};
      });
    } finally { sheet.remove(); }
  });
  expect(results).toEqual([
    {mode:'plain',printed:1,addedHeight:24,retainedText:true,restored:0,restoredHeight:24},
    {mode:'mixed',printed:1,addedHeight:24,retainedText:true,restored:0,restoredHeight:24},
    {mode:'width',printed:0,addedHeight:0,retainedText:true,restored:0,restoredHeight:24},
    {mode:'late',printed:0,addedHeight:0,retainedText:true,restored:0,restoredHeight:24},
  ]);
});

test('hanging HWP line recovery keeps the placed sheet and rejects overflow or a note anchor', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { layoutHwpAuthoredBodyLines } = await import('/src/hwp-body-line-metrics.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const literal = 'AAAA BBBB CCCC DDDD';
    const paper = document.createElement('div');
    paper.className = 'print-page'; paper.dataset.page = '3';
    paper.style.cssText = 'position:fixed;top:0;left:0;width:300px;height:300px';
    paper.innerHTML = '<div class="tiptap print-content" style="position:absolute;top:10px;left:10px;width:240px;height:80px;padding:0;min-height:0;box-sizing:border-box"></div>';
    document.body.append(paper);
    try {
      const body = paper.querySelector<HTMLElement>('.print-content')!;
      return ['fit', 'overflow', 'note'].map(mode => {
        body.style.height = mode === 'overflow' ? '60px' : '80px';
        body.innerHTML = '<p style="font-family:monospace;font-size:16px;line-height:24px;min-height:0;margin:0 0 0 30px;text-indent:-30px;text-align:justify;position:relative;left:4px">' +
          literal + '</p><p style="margin:0;height:24px;line-height:24px">AFTER</p>';
        const [p, tail] = [...body.children] as HTMLElement[];
        if (mode === 'note') tail.dataset.noteReference = 'note-1';
        const box = p.getBoundingClientRect();
        p.dataset.hwpSourceWrap = JSON.stringify({ fingerprint: hwpWrapFingerprint(literal), ends:[8,16],
          authoredBodyLines:{ hanging:30,size:16,pitch:24,width:240,
            origin:{ page:3, x:box.left-paper.getBoundingClientRect().left-30, y:box.top-paper.getBoundingClientRect().top } } });
        const oldTop = tail.getBoundingClientRect().top;
        layoutHwpAuthoredBodyLines(paper);
        return { mode, after: tail.getBoundingClientRect().top-oldTop, breaks:p.querySelectorAll('br').length,
          text:p.textContent, paper:paper.dataset.page };
      });
    } finally { paper.remove(); }
  });
  expect(result).toEqual([
    { mode:'fit', after:24, breaks:1, text:'AAAA BBBB CCCC DDDD', paper:'3' },
    { mode:'overflow', after:0, breaks:0, text:'AAAA BBBB CCCC DDDD', paper:'3' },
    { mode:'note', after:0, breaks:0, text:'AAAA BBBB CCCC DDDD', paper:'3' },
  ]);
});

test('source HWP line opportunities survive an aligned paper cut and reject a changed seam', async ({ page }) => {
  await page.goto('/');
  const measured = await page.evaluate(async () => {
    const { layoutHwpSourceCellLines, markHwpSourceCellLineGroups } = await import('/src/hwp-source-cell-lines.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const { paragraphLineOffsets, splitText } = await import('/src/paragraph-fragments.ts');
    const text = 'abcdefghijklmnopqrstuvwxyz1234';
    const host = document.createElement('div');
    host.className = 'tiptap';
    host.style.cssText = 'position:absolute;top:0;left:0;padding:0;width:500px';
    document.body.append(host);
    try {
      async function measure(mode: 'valid' | 'seam' | 'width' | 'edited', cooperative = false) {
        host.innerHTML = '<div class="print-page" data-page="1"><div class="print-content"></div></div>' +
          '<div class="print-page" data-page="2"><div class="print-content"></div></div>';
        const source = document.createElement('div');
        source.innerHTML = '<table><tbody><tr><td><p data-hwp-source-spaces="true" ' +
          'style="margin:0;line-height:28px;overflow-wrap:anywhere"><span style="font:700 16px/28px Courier New,monospace">' +
          (mode === 'edited' ? text.replace('f','z') : text) + '</span></p></td></tr></tbody></table>';
        const paragraph = source.querySelector('p')!;
        paragraph.dataset.hwpSourceWrap = JSON.stringify({ fingerprint: hwpWrapFingerprint(text),
          ends: mode === 'seam' ? [7, 14, 22, 30] : [7, 15, 22, 30], cellWidth: { outer: 84, text: 84 } });
        markHwpSourceCellLineGroups(source);
        const fragments = splitText(paragraph, 15);
        for (const [index, body] of [...host.querySelectorAll('.print-content')].entries()) {
          const table = source.querySelector('table')!.cloneNode(true) as HTMLTableElement;
          table.style.cssText = 'width:' + (index === 1 && mode === 'width' ? 88 : 84) + 'px;table-layout:fixed;border-collapse:collapse';
          const cell = table.rows[0].cells[0];
          cell.style.cssText = 'padding:0;border:0';
          cell.replaceChildren(fragments[index]);
          body.append(table);
        }
        const geometry = () => [...host.querySelectorAll('table,td,p')].map(node => {
          const { x, y, width, height } = node.getBoundingClientRect();
          return [x, y, width, height];
        });
        const before = geometry(), original = fragments.map(fragment => paragraphLineOffsets(fragment));
        let checkpoints = 0;
        await layoutHwpSourceCellLines(host, cooperative ? () => {
          checkpoints++;
          return new Promise<void>(resolve => setTimeout(resolve, 0));
        } : undefined);
        if (cooperative && checkpoints < 3) throw Error('Both page fragments must be checked cooperatively');
        return { before, after: geometry(), original, ends: fragments.map(fragment => paragraphLineOffsets(fragment)),
          text: host.textContent, opportunities: host.querySelectorAll('wbr').length,
          internalIds: host.querySelectorAll('[data-hwp-print-wrap-group]').length };
      }
      return { valid: await measure('valid'), seam: await measure('seam'),
        width: await measure('width'), edited: await measure('edited'),
        validCooperative: await measure('valid', true), widthCooperative: await measure('width', true) };
    } finally { host.remove(); }
  });
  expect(measured.validCooperative).toEqual(measured.valid);
  expect(measured.widthCooperative).toEqual(measured.width);
  expect(measured.valid.original).not.toEqual([[7], [7]]);
  expect(measured.valid.ends).toEqual([[7], [7]]);
  expect(measured.valid.text).toBe('abcdefghijklmnopqrstuvwxyz1234');
  expect(measured.valid.opportunities).toBe(2);
  for (const trial of Object.values(measured)) {
    expect(trial.after).toEqual(trial.before);
    expect(trial.internalIds).toBe(0);
  }
  for (const name of ['seam', 'width', 'edited'] as const) {
    expect(measured[name].opportunities).toBe(0);
    expect(measured[name].ends).toEqual(measured[name].original);
  }
});

test('a continuing HWP row fits against its lowest retained cell instead of the earliest cut', async ({ page }) => {
  await page.goto('/');
  const results = await page.evaluate(async () => {
    const { paginateCellRow } = await import('/src/editor-cell-pagination.ts');
    const { tableFragmenter } = await import('/src/table-fragments.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const host = document.createElement('div');
    host.className = 'tiptap';
    host.style.cssText = 'position:absolute;left:0;top:108px;min-height:0;padding:0;width:200px';
    document.body.append(host);
    try {
      return [{ source: true, rightPitch: 25, top: 108 }, { source: false, rightPitch: 25, top: 108 },
        { source: true, rightPitch: 20, top: 108 }, { source: true, rightPitch: 25, top: 150 }]
        .map(({ source, rightPitch, top }) => {
          host.style.top = `${top}px`;
          const paragraphs = (name: string, pitch: number) => [1, 2, 3].map(i =>
            `<p style="margin:0;font:10px/${pitch}px Arial;white-space:nowrap">${name} ${i}</p>`).join('');
          const cell = 'style="border:0;padding:0;vertical-align:top"';
          host.innerHTML = '<table style="width:200px;border-collapse:collapse;table-layout:fixed;margin:0">' +
            `<tbody><tr data-row-split="allow" style="height:120px"><td ${cell}>${paragraphs('green', 20)}</td>` +
            `<td ${cell}>${paragraphs('blue', rightPitch)}</td></tr></tbody></table>`;
          const row = host.querySelector('tr')!, paragraph = row.cells[0].querySelector('p')!;
          if (source) {
            paragraph.dataset.hwpSourceSpaces = 'true';
            paragraph.dataset.hwpSourceWrap = JSON.stringify({ fingerprint: hwpWrapFingerprint(paragraph.textContent!), ends: [2, 6] });
          }
          const view = { state: { doc: { nodeAt: () => ({ nodeSize: 10 }) } }, posAtDOM: () => 2 } as unknown as import('@tiptap/pm/view').EditorView;
          const flow = paginateCellRow(view, row, 1, { origin: 0, scale: 1, top: 20, bottom: 40, height: 200, extra: 0, removed: [] });
          const authored = Array.from(row.cells, cell => cell.textContent);
          const divider = tableFragmenter({ bodyHeight: () => 140, hasDeferredNote: () => false, deferredNotes: () => new Set() });
          const tail = divider.splitElement(row, 160) as HTMLTableRowElement | null;
          return { extra: flow?.extra ?? null, printedExtra: tail ? 84 + 160 - row.getBoundingClientRect().bottom : null,
            widgets: flow?.signatures.filter(key => !key.startsWith('row:')).length ?? 0,
            text: Array.from(row.cells, (cell, i) => (cell.textContent || '') + (tail?.cells[i].textContent || '')), authored };
        });
    } finally { host.remove(); }
  });
  expect(results.map(r => [r.extra, r.printedExtra])).toEqual([[86, 86], [96, 86], [96, 96], [null, null]]);
  expect(results.map(r => r.widgets)).toEqual([2, 2, 2, 0]);
  for (const result of results) expect(result.text).toEqual(result.authored);
});

test('a retained HWP cell line keeps the same cut after the editor inserts its spacer', async ({ page }) => {
  await page.goto('/');
  const results = await page.evaluate(async () => {
    const { paginateCellRow } = await import('/src/editor-cell-pagination.ts');
    const { hwpWrapFingerprint, normalizeHwpWrapText } = await import('/src/hwp-text-advances.ts');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
    const host = document.createElement('div');
    host.className = 'tiptap';
    host.style.cssText = 'position:absolute;left:0;top:62px;min-height:0;padding:0;width:290px';
    document.body.append(host);
    const text = 'Alpha bravo charlie delta echo foxtrot golf hotel';
    try {
      return ['original', 'with-widget', 'changed-text', 'outside'].map(mode => {
        const cell = 'style="padding:2px;border:0;vertical-align:top"';
        host.innerHTML = '<table style="width:290px;table-layout:fixed;border:0;border-collapse:collapse;margin:0">' +
          `<tbody><tr data-row-split="allow" style="height:155px"><td ${cell}>` +
          `<p data-exact-line-height="true" data-hwp-source-spaces="true" style="font:16px/28px Arial;margin:0px">${text}</p>` +
          `<p style="font:16px/28px Arial;margin:0px">Continuation paragraph</p></td><td ${cell}>` +
          '<p style="font:16px/28px Arial;margin:0px">Other column</p></td></tr></tbody></table>';
        const row = host.querySelector('tr')!, [first, next] = row.cells[0].querySelectorAll('p');
        if (mode === 'with-widget') {
          const widget = document.createElement('span');
          widget.className = 'ProseMirror-widget pagination-cell-gap';
          widget.style.display = 'none';
          first.after(widget);
        }
        const ends = [...paragraphLineOffsets(first), text.length].map(end => normalizeHwpWrapText(text.slice(0, end)).length);
        first.dataset.hwpSourceWrap = JSON.stringify({ fingerprint: hwpWrapFingerprint(mode === 'changed-text' ? text + 'x' : text), ends });
        const range = document.createRange(); range.selectNodeContents(first);
        const inkEnd = Math.max(...[...range.getClientRects()].filter(rect => rect.width > 0).map(rect => rect.bottom));
        const box = first.getBoundingClientRect();
        const cellEnd = mode === 'outside' ? inkEnd - 1 : (inkEnd + box.bottom) / 2;
        const view = { state: { doc: { nodeAt: () => ({ nodeSize: 10 }) } }, posAtDOM: (node: Node, offset: number) =>
          node === first ? 11 : node === next ? 101 : first.contains(node) ? 20 + offset : 200 + offset } as unknown as import('@tiptap/pm/view').EditorView;
        const flow = paginateCellRow(view, row, 1, { origin: 0, scale: 1, top: 20, bottom: 40, height: cellEnd + 42, extra: 0, removed: [] });
        return { extra: flow?.extra, widgets: flow?.signatures.filter(key => !key.startsWith('row:')).map(key => Number(key.split(':')[0])),
          glyphRoom: cellEnd - inkEnd, strutExcess: box.bottom - cellEnd, text: first.textContent };
      });
    } finally { host.remove(); }
  });
  expect(results[0].glyphRoom).toBeGreaterThan(1);
  expect(results[0].strutExcess).toBeGreaterThan(1);
  expect(results[0].widgets).toEqual([100]);
  expect(results[1]).toEqual(results[0]);
  for (const result of results.slice(2)) {
    expect(result.widgets).toHaveLength(1);
    expect(result.widgets![0]).toBeLessThan(100);
    expect(result.text).toBe(results[0].text);
  }
});

test('a short source label recovers one line without changing its table width', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { layoutShortHwpCellWidths } = await import('/src/hwp-text-advance-layout.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
    const label = '예제문구';
    const host = document.createElement('div');
    host.className = 'tiptap';
    host.style.cssText = 'position:fixed;left:0;top:0;font:16px/20px sans-serif';
    document.body.append(host);
    try {
      const sample = document.createElement('span');
      sample.style.whiteSpace = 'nowrap';
      sample.textContent = label;
      host.append(sample);
      const advance = sample.getBoundingClientRect().width;
      const measure = (style: 'original' | 'edited' | 'resized' | 'would overflow') => {
        const moved = style === 'resized' ? 6 : 0;
        host.innerHTML = `<table style="width:${advance + 4.1 + moved}px;table-layout:fixed;border-collapse:collapse">
          <tbody><tr><td style="padding:2px;border:1px solid black"><p style="margin:0;line-height:20px;font-size:16px;">
          </p></td></tr></tbody></table>`;
        const table = host.querySelector('table')!, cell = table.rows[0].cells[0], paragraph = cell.querySelector('p')!;
        paragraph.textContent = style === 'edited' ? '예제문장' : label;
        const before = cell.getBoundingClientRect();
        paragraph.dataset.hwpSourceWrap = JSON.stringify({ fingerprint: hwpWrapFingerprint(label), ends: [label.length],
          cellWidth: { outer: before.width - moved, text: advance + (style === 'would overflow' ? 4 : .05) } });
        const original = paragraphLineOffsets(paragraph).length;
        layoutShortHwpCellWidths(table);
        const after = cell.getBoundingClientRect();
        const range = document.createRange();
        range.selectNodeContents(paragraph);
        return { original, wrapping: paragraphLineOffsets(paragraph).length,
          rowHeight: [before.height, after.height], widthUnchanged: before.width === after.width,
          inkInside: [...range.getClientRects()].every(r => r.left >= after.left && r.right <= after.right),
          changed: !!paragraph.style.width, text: paragraph.textContent };
      };
      return { original: measure('original'), edited: measure('edited'), resized: measure('resized'), overflow: measure('would overflow') };
    } finally { host.remove(); }
  });
  expect(result.original.original).toBe(1);
  expect(result.original.wrapping).toBe(0);
  expect(result.original.rowHeight[1]).toBeLessThan(result.original.rowHeight[0]);
  expect(result.original.widthUnchanged).toBe(true);
  expect(result.original.inkInside).toBe(true);
  expect(result.original.text).toBe('예제문구');
  for (const control of [result.edited, result.resized, result.overflow]) {
    expect(control.changed).toBe(false);
    expect(control.widthUnchanged).toBe(true);
  }
});

test('a scaled source suffix uses its measured width through editing and DOCX saves', async ({ page }) => {
  await page.goto('/');
  const prefix = 'MM MMM ', suffix = 'WWWWWWWWWW', value = prefix + suffix;
  const metadata = await page.evaluate(async value => {
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    return JSON.stringify({ fingerprint: hwpWrapFingerprint(value), ends: [15],
      cellWidth: { outer: 140, text: 139 }, scaleRun: { start: 7, scale: 80, width: 60, size: 16, font: 'Courier New' } });
  }, value);
  const width = { size: 2100, type: WidthType.DXA } as const;
  const content = new Paragraph({ indent: { left: 180, hanging: 180 },
    children: [new TextRun({ text: prefix, size: 24, font: 'Courier New' }),
      new TextRun({ text: suffix, size: 24, font: 'Courier New', scale: 80 })] });
  const table = new Table({ width, layout: 'fixed', columnWidths: [2100], rows: [new TableRow({
    children: [new TableCell({ width, margins: { top: 0, left: 0, bottom: 0, right: 0 }, children: [content] })],
  })] });
  const zip = await JSZip.loadAsync(await Packer.toBuffer(new Document({ sections: [{ children: [table] }] })));
  const xmlPath = 'word/document.xml', raw = await zip.file(xmlPath)!.async('string');
  zip.file(xmlPath, raw.replace(/<w:p(?=[ >])/, `<w:p xmlns:wsa="urn:wordstudio:hwp-source-spaces:1" mc:Ignorable="wsa" mc:PreserveAttributes="wsa:wrap" wsa:wrap="${metadata.replaceAll('"', '&quot;')}"`));
  let bytes = await zip.generateAsync({ type: 'nodebuffer' });
  for (let stage = 0; stage < 3; stage++) {
    const name = `mixed-source-run-${stage}`;
    await page.locator('input[type=file]').first().setInputFiles({
      name: name + '.docx',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      buffer: bytes,
    });
    await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(name);
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
    const paragraph = editor.locator('td > p');
    await expect(paragraph).toHaveText(value);
    const scaled = paragraph.locator('[data-hwp-source-advance-run]');
    const expected = async () => {
      await expect(scaled).toHaveText(suffix);
      await expect.poll(() => scaled.evaluate(node => node.getBoundingClientRect().width)).toBeCloseTo(60, 1);
    };
    await expected();
    if (stage === 0) {
      await scaled.click({ position: { x: 25, y: 7 } });
      await page.keyboard.type('X');
      await expect(paragraph).toContainText('X');
      await expect(scaled).toHaveCount(0);
      await page.keyboard.press('ControlOrMeta+z');
      await expect(paragraph).toHaveText(value);
      await expected();
      await page.keyboard.press('ControlOrMeta+Shift+z');
      await expect(scaled).toHaveCount(0);
      await page.keyboard.press('ControlOrMeta+z');
      await expected();
    }
    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const printed = page.locator('.print-layout.preview .print-content td > p');
    await expect(printed).toHaveText(value);
    const measured = await printed.evaluate(async p => {
      const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
      const span = p.querySelector<HTMLElement>('span[style*="transform: scaleX"]');
      return { width: span?.getBoundingClientRect().width, breaks: paragraphLineOffsets(p).length };
    });
    expect(measured.width).toBeCloseTo(60, 1);
    expect(measured.breaks).toBe(0);
    await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
    if (stage === 2) break;
    await page.getByRole('button', { name: '파일', exact: true }).click();
    const download = page.waitForEvent('download');
    await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
    bytes = await readFile((await (await download).path())!);
    const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
    expect(xml).toContain('wsa:wrap');
    expect(xml).toContain('<w:w w:val="80"');
    expect(xml).not.toContain('scaleX');
  }
});

test('unchanged rounded HWP cell grids retain original line ends only when their painted lines fit', async ({ page }) => {
  await page.goto('/');
  const cases = await page.evaluate(async () => {
    const { layoutHwpSourceCellLines } = await import('/src/hwp-source-cell-lines.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
    const text = 'abcdefghijklmnopqrstuv';
    const host = document.createElement('div');
    host.className = 'tiptap';
    host.lang = 'en';
    host.style.cssText = 'position:fixed;top:0;left:-10000px;width:300px';
    document.body.append(host);
    try {
      return ['source', 'wider column', 'more padding', 'horizontal overflow', 'vertical overflow', 'changed text'].map(mode => {
        const width = mode === 'wider column' ? 101 : 100;
        const padding = mode === 'more padding' ? 1 : 0;
        const lineHeight = mode === 'vertical overflow' ? 10 : 24;
        const breakAt = mode === 'horizontal overflow' ? 16 : 11;
        host.innerHTML = `<section class="print-page" style="margin:0;padding:0">
          <div class="print-content"><table data-word-table-layout="fixed"
            style="width:${width}px;table-layout:fixed;border-collapse:collapse">
            <colgroup><col style="width:${width}px"></colgroup>
            <tbody><tr><td style="padding:${padding}px;border:none">
              <p data-hwp-source-spaces="true"
                style="margin:0;text-align:justify;font:12px/${lineHeight}px monospace;overflow-wrap:anywhere">${text}</p>
            </td></tr></tbody>
          </table></div></section>`;
        const paragraph = host.querySelector('p')!;
        paragraph.dataset.hwpSourceWrap = JSON.stringify({
          fingerprint: hwpWrapFingerprint(text), ends: [breakAt, text.length],
          cellWidth: { outer: 99.6, text: 99.6 },
        });
        if (mode === 'changed text') paragraph.textContent = 'z' + text.slice(1);
        const literal = paragraph.textContent;
        const geometry = () => [...host.querySelectorAll('table,tr,td,p')].map(element => element.getBoundingClientRect().toJSON());
        const before = geometry(), oldEnds = paragraphLineOffsets(paragraph);
        layoutHwpSourceCellLines(host);
        return { mode, before, after: geometry(), oldEnds, ends: paragraphLineOffsets(paragraph),
          breaks: paragraph.querySelectorAll('br').length, literal, actual: paragraph.textContent };
      });
    } finally { host.remove(); }
  });
  const [source, ...controls] = cases;
  expect(source.oldEnds).toHaveLength(1);
  expect(source.oldEnds).not.toEqual([11]);
  expect(source.ends).toEqual([12]);
  expect(source.breaks).toBe(1);
  for (const control of controls) {
    expect(control.breaks, control.mode).toBe(0);
    expect(control.ends, control.mode).toEqual(control.oldEnds);
  }
  for (const observation of cases) {
    expect(observation.after, observation.mode).toEqual(observation.before);
    expect(observation.actual, observation.mode).toBe(observation.literal);
  }
});

test('a measured two-line cell in a column uses only its unchanged rounded grid and row height', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { layoutHwpSourceCellLines } = await import('/src/hwp-source-cell-lines.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
    const host = document.createElement('div');
    host.className = 'tiptap';
    host.style.cssText = 'position:fixed;top:0;left:-10000px';
    document.body.append(host);
    const literal = 'abcdefghijklmnopqrstuv';
    const proof = { fingerprint: hwpWrapFingerprint(literal), ends: [11, literal.length],
      cellWidth: { outer: 99.6, text: 99.6 } };
    try {
      return ['valid', 'width', 'height', 'text', 'single-column'].map(mode => {
        const width = mode === 'width' ? 103 : 100;
        host.innerHTML = `<section class="print-page"><div data-column-count="${mode === 'single-column' ? 1 : 2}">
          <table data-word-table-layout="fixed" style="width:${width}px;table-layout:fixed;border-collapse:collapse">
            <colgroup><col style="width:${width}px"></colgroup>
            <tbody><tr style="height:60px"><td style="width:${width}px;padding:0;border:none">
              <p style="margin:0;text-align:justify;font:12px/24px monospace;overflow-wrap:anywhere">${mode === 'text' ? literal + 'z' : literal}</p>
            </td></tr></tbody>
          </table></div></section>`;
        const text = host.querySelector('p')!, cell = host.querySelector('td')!;
        const dimensions = () => [text, cell].flatMap(element => {
          const { width, height } = element.getBoundingClientRect();
          return [width, height];
        });
        const rectBefore = dimensions();
        text.dataset.hwpSourceWrap = JSON.stringify({ ...proof, columnCellHeight: rectBefore[3] + (mode === 'height' ? 3 : 0) });
        const before = paragraphLineOffsets(text);
        layoutHwpSourceCellLines(host);
        return { mode, before, after: paragraphLineOffsets(text), markers: text.querySelectorAll('wbr').length,
          text: text.textContent, rectBefore, rectAfter: dimensions() };
      });
    } finally { host.remove(); }
  });
  await test.info().attach('synthetic cell geometry', { body: JSON.stringify(result), contentType: 'application/json' });
  const [valid, ...controls] = result;
  expect(valid.before).toHaveLength(1);
  expect(valid.before).not.toEqual([11]);
  expect(valid.after).toEqual([11]);
  expect(valid.markers).toBe(1);
  expect(valid.text).toBe('abcdefghijklmnopqrstuv');
  for (const item of result) expect(item.rectAfter, item.mode).toEqual(item.rectBefore);
  for (const item of controls) {
    expect(item.markers, item.mode).toBe(0);
    expect(item.after, item.mode).toEqual(item.before);
  }
});

test('intact HWP column cells recover three stored line ends at their original pitch', async ({ page }) => {
  await page.goto('/');
  const cases = await page.evaluate(async () => {
    const { layoutHwpSourceCellLines } = await import('/src/hwp-source-cell-lines.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
    const host = document.createElement('div');
    host.className = 'tiptap';
    host.style.cssText = 'position:fixed;top:0;left:-10000px';
    document.body.append(host);
    const literal = 'abcdefghijklmnopqrstuvwxyz0123456789';
    const ends = [12, 24, literal.length];
    try {
      return ['same', 'pitch', 'margin', 'neighbor'].map(mode => {
        host.innerHTML = `<section class="print-page"><div class="print-content" style="height:${mode === 'margin' ? 85 : 200}px">
          <div data-column-count="2"><table data-word-table-layout="fixed" style="width:100px;table-layout:fixed;border-collapse:collapse">
            <colgroup><col style="width:100px"></colgroup>
            <tbody><tr style="height:90px"><td style="padding:0;border:none">
              <p style="margin:0;text-align:justify;font-family:monospace;font-size:12px;line-height:${mode === 'pitch' ? 23 : 24}px;overflow-wrap:anywhere">${literal}</p>
              ${mode === 'neighbor' ? '<p style="margin:0">Another paragraph</p>' : ''}
            </td></tr></tbody>
          </table></div>
        </div></section>`.replace(/>\s+</g, '><');
        const paragraph = host.querySelector('p')!, cell = host.querySelector('td')!;
        const geometry = () => [paragraph, cell].map(element => {
          const { x, y, width, height } = element.getBoundingClientRect();
          return { x, y, width, height };
        });
        paragraph.dataset.hwpSourceWrap = JSON.stringify({ fingerprint: hwpWrapFingerprint(literal), ends,
          cellWidth: { outer: 99.6, text: 99.6 }, columnCellHeight: cell.getBoundingClientRect().height,
          columnCellPitch: 24 });
        const before = paragraphLineOffsets(paragraph), boxes = geometry();
        layoutHwpSourceCellLines(host);
        return { mode, before, after: paragraphLineOffsets(paragraph), boxes, afterBoxes: geometry(),
          literal: paragraph.textContent, breaks: paragraph.querySelectorAll('wbr').length };
      });
    } finally { host.remove(); }
  });
  const [matching, ...controls] = cases;
  expect(matching.before).toHaveLength(2);
  expect(matching.before).not.toEqual([12, 24]);
  expect(matching.after).toEqual([12, 24]);
  expect(matching.breaks).toBe(2);
  for (const result of cases) {
    expect(result.afterBoxes, result.mode).toEqual(result.boxes);
    expect(result.literal).toBe('abcdefghijklmnopqrstuvwxyz0123456789');
  }
  for (const result of controls) {
    expect(result.after, result.mode).toEqual(result.before);
    expect(result.breaks, result.mode).toBe(0);
  }
});

test('restores a witnessed HWP body boundary and rolls back when the next paragraph still overflows', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { restoreHwpBodyBeforePageBreak } = await import('/src/hwp-body-line-metrics.ts');
    const { hwpSingleBodyFace, hwpWrapFingerprint, normalizeHwpWrapText } = await import('/src/hwp-text-advances.ts');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
    const fallback = await new FontFace('Noto Sans KR', 'url(/src/assets/fonts/NotoSansMono.ttf)',
      { sizeAdjust: '75%', unicodeRange: 'U+0020,U+0041-005A' }).load();
    document.fonts.add(fallback);
    const paper = document.createElement('div');
    paper.className = 'print-page';
    paper.dataset.page = '1';
    paper.style.cssText = 'position:fixed;top:0;left:0;width:500px;height:500px';
    const body = document.createElement('div');
    body.className = 'tiptap print-content';
    body.style.cssText = 'position:absolute;top:10px;left:10px;padding:0;min-height:0';
    paper.append(body);
    document.body.append(paper);
    try {
      const literal = 'WWW WWW WWWWW WWW WWW', family = JSON.stringify(hwpSingleBodyFace);
      const paragraphStyle = 'font-size:16px;line-height:24px;min-height:0;margin:0 0 0 10px;' +
        'text-indent:-10px;text-align:justify;overflow-wrap:normal;';
      const probe = document.createElement('p');
      probe.style.cssText = paragraphStyle + 'font-family:' + family;
      body.append(probe);
      // The first face falls back to the operating system until Noto is
      // requested. Calibrate the fixture against that platform's glyph widths.
      const width = Array.from({ length: 51 }, (_, index) => index + 90).find(measure => {
        body.style.width = measure + 'px';
        probe.textContent = literal;
        probe.style.fontFamily = family;
        if (paragraphLineOffsets(probe).length !== 2) return false;
        probe.style.fontFamily += ', "Noto Sans KR"';
        const natural = paragraphLineOffsets(probe);
        if (natural.length !== 1 || normalizeHwpWrapText(literal.slice(0, natural[0])).length === 6) return false;
        probe.replaceChildren(literal.slice(0, 8), document.createElement('br'), literal.slice(8));
        return paragraphLineOffsets(probe).length === 1;
      });
      if (!width) throw new Error('No common synthetic width for the two fonts');
      const makeParagraph = (text: string, ends: number[], style: string) => {
        const paragraph = document.createElement('p');
        paragraph.style.cssText = paragraphStyle + style;
        paragraph.textContent = text;
        body.append(paragraph);
        const box = paragraph.getBoundingClientRect(), sheet = paper.getBoundingClientRect();
        paragraph.dataset.hwpSourceWrap = JSON.stringify({ fingerprint:hwpWrapFingerprint(text), ends,
          authoredBodyLines:{ hanging:10, size:16, pitch:24, width,
            origin:{ page:1, x:box.left-sheet.left-10, y:box.top-sheet.top } } });
        return paragraph;
      };
      return [
        { height: 100, permitPrintBreak: true, useNaturalBreak: false },
        { height: 60, permitPrintBreak: true, useNaturalBreak: false },
        { height: 100, permitPrintBreak: false, useNaturalBreak: false },
        { height: 100, permitPrintBreak: false, useNaturalBreak: true },
      ].map(({ height, permitPrintBreak, useNaturalBreak }) => {
        body.style.height = height + 'px';
        body.replaceChildren();
        const prior = makeParagraph(literal, [6,17], 'font-family:' + family);
        const last = makeParagraph('LAST END', [4,7], 'font-family:monospace;word-spacing:200px');
        const originalLines = paragraphLineOffsets(prior).length + 1;
        prior.style.fontFamily += ', "Noto Sans KR"';
        const [naturalBreak] = paragraphLineOffsets(prior);
        prior.style.fontFamily = family;
        const naturalEnd = normalizeHwpWrapText(literal.slice(0, naturalBreak)).length;
        if (useNaturalBreak) {
          const proof = JSON.parse(prior.dataset.hwpSourceWrap!);
          proof.ends[0] = naturalEnd;
          prior.dataset.hwpSourceWrap = JSON.stringify(proof);
        }
        const before = prior.outerHTML;
        const fits = () => last.getBoundingClientRect().bottom <= body.getBoundingClientRect().bottom;
        const initiallyFits = fits();
        const restored = restoreHwpBodyBeforePageBreak(last, fits, permitPrintBreak);
        return { originalLines, naturalEnd,
          initiallyFits, restored, fits:fits(), synthetic:prior.querySelectorAll('br').length,
          literal:prior.textContent === literal, sameMarkup:prior.outerHTML === before };
      });
    } finally {
      document.fonts.delete(fallback);
      paper.remove();
    }
  });
  expect(result.map(({naturalEnd, ...row}) => ({...row, explicit:naturalEnd !== 6}))).toEqual([
    { originalLines:3, explicit:true, initiallyFits:false, restored:true, fits:true, synthetic:1, literal:true, sameMarkup:false },
    { originalLines:3, explicit:true, initiallyFits:false, restored:false, fits:false, synthetic:0, literal:true, sameMarkup:true },
    { originalLines:3, explicit:true, initiallyFits:false, restored:false, fits:false, synthetic:0, literal:true, sameMarkup:true },
    { originalLines:3, explicit:true, initiallyFits:false, restored:true, fits:true, synthetic:0, literal:true, sameMarkup:false },
  ]);
});

test('HWPX rows use stored line ends without changing authored character scaling', async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async () => {
    const { layoutHwpExpandedCellRows } = await import('/src/hwp-source-cell-lines.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const host = document.createElement('div');
    host.className = 'tiptap';
    host.style.cssText = 'position:absolute;left:10px;top:10px;width:160px';
    document.body.append(host);
    try {
      return ['valid', 'edited', 'width', 'exact'].map(mode => {
        host.innerHTML = '<table style="width:150px;table-layout:fixed;border-collapse:collapse">' +
          '<colgroup><col style="width:30px"><col style="width:120px"></colgroup><tbody><tr>' +
          '<td style="border:0;padding:0;line-height:24px;font-size:16px">표</td><td style="border:0;padding:2px 4px;vertical-align:middle">' +
          '<p style="font:16px/24px monospace;min-height:0;margin:0"><span style="--word-character-scale:95">AA </span>BB CC</p>' +
          '</td></tr></tbody></table>';
        const table = host.querySelector('table')!;
        const row = table.rows[0], p = row.querySelector('p')!;
        if (mode === 'exact') row.dataset.wordRowRule = 'exact';
        const before = row.getBoundingClientRect().height;
        const paragraphWidth = p.getBoundingClientRect().width;
        const cellWidth = p.closest('td')!.getBoundingClientRect().width;
        const different = mode === 'width' ? 5 : 0;
        p.dataset.hwpSourceWrap = JSON.stringify({
          fingerprint:hwpWrapFingerprint(mode === 'edited' ? 'AA BB DC' : 'AA BB CC'),
          ends:[4,6], cellWidth:{ outer:cellWidth+different, text:paragraphWidth+different }, expandedRowHeight:before + 24,
        });
        layoutHwpExpandedCellRows(table);
        return { mode, lines:p.querySelectorAll('br').length, delta:row.getBoundingClientRect().height - before,
          text:p.textContent, scale:p.querySelector('span')!.style.getPropertyValue('--word-character-scale') };
      });
    } finally { host.remove(); }
  });
  expect(states).toEqual(['valid', 'edited', 'width', 'exact'].map(mode => ({
    mode, lines:mode === 'valid' ? 1 : 0, delta:mode === 'valid' ? 24 : 0, text:'AA BB CC', scale:'95',
  })));
});

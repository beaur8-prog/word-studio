import { expect, test } from '@playwright/test';
import JSZip from 'jszip';

test('a measured HWP line before a forced column break fits in the editor and through two DOCX saves', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const paragraph = (text: string, attributes = '') =>
      `<p ${attributes} style="margin:0;line-height:32px;font-family:Arial;font-size:15pt"><span style="font-size:15pt;font-family:Arial">${text}</span></p>`;
    const html = '<div data-column-count="2" data-column-space="12" data-column-balance="false">' +
      ['FIRST', 'SECOND', 'THIRD'].map(text => paragraph(text)).join('') +
      paragraph('LAST GLYPH', 'data-hwp-body-line="4 160 20 32"') +
      '<div data-column-break="true"></div>' + paragraph('RIGHT COLUMN') + '</div>';
    let doc: any = {
      format: 'word-studio', version: 1, id: crypto.randomUUID(), title: 'Column terminal', comments: [],
      updatedAt: new Date().toISOString(), content: parseDocumentHtml(html),
      page: { size: 'Custom', width: 132.29167, height: 38.1, orientation: 'portrait',
        margin: 2.64583, header: '', footer: '', pageNumbers: false },
    };
    const original = doc;
    const render = async (d: any) => {
      const printed = await createPrintLayout(generateHTML(d.content, editorExtensions), d.page);
      try {
        const blocks = [...printed.element.querySelectorAll<HTMLElement>('.print-content p')];
        const last = blocks.find(p => p.textContent === 'LAST GLYPH');
        const right = blocks.find(p => p.textContent === 'RIGHT COLUMN');
        return {
          pages: printed.pages, text: blocks.map(p => p.textContent), terminal: !!last?.dataset.hwpColumnEnd,
          lastLeft: last?.getBoundingClientRect().left, rightLeft: right?.getBoundingClientRect().left,
        };
      } finally { printed.dispose(); }
    };
    const plain = { ...doc, content: parseDocumentHtml(html.replace('data-hwp-body-line="4 160 20 32"', '')) };
    const withoutBreak = await render({ ...doc, content: parseDocumentHtml(html.replace('<div data-column-break="true"></div>', '')) });
    const withoutFollowingBlock = await render({ ...doc, content: parseDocumentHtml(html.replace(
      '<div data-column-break="true"></div>' + paragraph('RIGHT COLUMN'), '',
    )) });
    const largerRun = { ...doc, content: parseDocumentHtml(html.replace(
      'font-size:15pt;font-family:Arial">LAST GLYPH', 'font-size:36pt;font-family:Arial">LAST GLYPH',
    )) };
    const largerRoundTrip = await importFile(new File([await exportDocx(largerRun)], 'column-enlarged.docx'));
    const largeColumn = parseDocumentHtml(largerRoundTrip.html!).content!.find(n => n.type === 'columnGroup');
    const largerParagraph = largeColumn?.content?.find(n => n.content?.[0]?.text === 'LAST GLYPH');
    const withoutSourceLine = await render(plain), stages = [];
    for (let save = 0; save < 3; save++) {
      stages.push(await render(doc));
      if (save === 2) break;
      const again = await importFile(new File([await exportDocx(doc)], 'source-column.docx'));
      doc = { ...again.document, content: parseDocumentHtml(again.html!) };
    }
    return { original, withoutSourceLine, withoutBreak, withoutFollowingBlock, largerAttrs: largerParagraph?.attrs, stages };
  });
  expect(result.withoutSourceLine.pages).toBe(2);
  expect(result.withoutBreak.terminal).toBe(false);
  expect(result.withoutFollowingBlock.terminal).toBe(false);
  expect(result.largerAttrs?.lineSpacing).toBe('57.6pt');
  expect(result.largerAttrs?.hwpBodyLine).toBeNull();
  for (const stage of result.stages) {
    expect(stage.pages).toBe(1);
    expect(stage.terminal).toBe(true);
    expect(stage.text).toEqual(['FIRST', 'SECOND', 'THIRD', 'LAST GLYPH', 'RIGHT COLUMN']);
    expect(stage.rightLeft).toBeGreaterThan(stage.lastLeft);
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'source-column.wordstudio', mimeType: 'application/json',
    buffer: Buffer.from(JSON.stringify(result.original)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  const terminal = editor.locator('p').filter({ hasText: 'LAST GLYPH' });
  await expect(terminal).toHaveAttribute('data-hwp-column-end', 'true');
  await expect(page.locator('.document-stack')).toHaveAttribute('data-pages', '1');
  await terminal.click();
  await page.keyboard.press('End');
  await page.keyboard.insertText(' LONGER WORDS REQUIRING THE ORIGINAL PARAGRAPH TO WRAP AGAIN');
  await expect(terminal).toContainText('WRAP AGAIN');
  await expect(terminal).not.toHaveAttribute('data-hwp-column-end', 'true');
  await page.keyboard.press('ControlOrMeta+z');
  await expect(terminal).toHaveText('LAST GLYPH');
  await expect(terminal).toHaveAttribute('data-hwp-column-end', 'true');
});

test('native column terminal paint requires the same words, column measure and remaining space', async ({ page }) => {
  await page.goto('/');
  const evidence = await page.evaluate(async () => {
    const { measureHwpColumnEnd } = await import('/src/hwp-column-end.ts');
    const { measureSourceColumnEnd, layoutSourceColumnTail } = await import('/src/hwp-column-source-layout.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const { encodeHwpBodyLine, parseHwpBodyLine } = await import('/src/hwp-body-line.ts');
    const root = document.createElement('div');
    root.className = 'tiptap';
    root.style.cssText = 'position:fixed;left:-10000px;top:0;width:420px';
    root.innerHTML = `<div data-column-count="2">
      <p data-exact-line-height="true" style="margin:0;font-family:Arial;font-size:15pt;--exact-line-height:32px;line-height:32px"><span>Preceding short line </span></p>
      <p data-exact-line-height="true" style="margin:0;font-family:Arial;font-size:15pt;--exact-line-height:32px;line-height:32px"><span>Ending with descenders pyg</span></p>
      <div data-column-break="true"></div>
    </div>`;
    document.body.append(root);
    try {
      await document.fonts.ready;
      const [prior, terminal] = root.querySelectorAll('p');
      const standard = { percent: 160, font: 20, base: 32, maxFont: 20, align: 'column-end' as const };
      const ordinary = measureHwpColumnEnd(terminal, standard, 1) ?? standard.base;
      const marker = document.createElement('span');
      marker.style.cssText = 'display:inline-block;vertical-align:baseline;width:0;height:0;line-height:0';
      terminal.dataset.hwpColumnEnd = 'true';
      terminal.style.setProperty('--hwp-column-em', '20px');
      terminal.append(marker);
      const measuredBase = marker.getBoundingClientRect().bottom - terminal.getBoundingClientRect().top;
      marker.remove();
      terminal.style.removeProperty('--hwp-column-em');
      delete terminal.dataset.hwpColumnEnd;

      const proof = { baseline: measuredBase - 2.5, remainder: 21,
        width: terminal.getBoundingClientRect().width, fingerprint: hwpWrapFingerprint(terminal.textContent!),
        preceding: hwpWrapFingerprint(prior.textContent!), hang: true as const };
      const encoded = encodeHwpBodyLine({ ...standard, columnFit: proof }), source = parseHwpBodyLine(encoded)!;
      const bodyBottom = prior.getBoundingClientRect().bottom + ordinary - .6;
      const baselineLiteral = prior.textContent;
      layoutSourceColumnTail(terminal, source);
      const tail = prior.querySelector('[data-hwp-column-tail]')?.textContent;
      const fitted = measureSourceColumnEnd(terminal, source, 1, bodyBottom, ordinary);
      const ordinaryFits = measureSourceColumnEnd(terminal, source, 1, bodyBottom + 3, ordinary);
      const tighterRoom = (fitted?.height ?? ordinary) - 1;
      const tighterBottom = prior.getBoundingClientRect().bottom + tighterRoom;
      const footerProof = { ...proof, remainder: tighterRoom + 1 };
      const footerFit = measureSourceColumnEnd(terminal, { ...source, columnFit: footerProof }, 1, tighterBottom, ordinary);
      const changedFooter = measureSourceColumnEnd(terminal, { ...source, columnFit: { ...footerProof, remainder: tighterRoom + 1.5 } },
        1, tighterBottom, ordinary);
      const currentProof = source.columnFit;
      const badMeasure = measureSourceColumnEnd(terminal, { ...source, columnFit: { ...proof, width: proof.width - 10 } }, 1, bodyBottom, ordinary);
      terminal.firstElementChild!.textContent += ' changed';
      const edited = measureSourceColumnEnd(terminal, source, 1, bodyBottom, ordinary);
      const badSource = parseHwpBodyLine(encoded.replace(proof.fingerprint, 'invalid'));
      return { tail, preserved: baselineLiteral === prior.textContent, ordinary, fitted, ordinaryFits, tighterRoom, footerFit, changedFooter, currentProof, badMeasure, edited, badSource };
    } finally { root.remove(); }
  });
  expect(evidence.tail).toBe(' ');
  expect(evidence.preserved).toBe(true);
  expect(evidence.fitted?.offset).toBeCloseTo(-2.5);
  expect(evidence.fitted?.height).toBeLessThan(evidence.ordinary);
  expect(evidence.ordinaryFits).toBeNull();
  expect(evidence.footerFit?.offset).toBeCloseTo(-3.5);
  expect(evidence.footerFit?.height).toBeLessThanOrEqual(evidence.tighterRoom);
  expect(evidence.changedFooter).toBeNull();
  expect(evidence.badMeasure).toBeNull();
  expect(evidence.edited).toBeNull();
  expect(evidence.badSource).toBeNull();
});

test('the exact HWP column tail keeps its space and standard DOCX text through two saves', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const { encodeHwpBodyLine } = await import('/src/hwp-body-line.ts');
    const penultimate = 'Prior line ', terminal = 'Terminal line';
    const line = encodeHwpBodyLine({ percent: 160, font: 20, base: 32, maxFont: 20, align: 'column-end',
      columnFit: { baseline: 15, remainder: 21, width: 234, fingerprint: hwpWrapFingerprint(terminal),
        preceding: hwpWrapFingerprint(penultimate), hang: true } });
    const para = (text: string, metadata = '') =>
      `<p ${metadata} style="margin:0;font-family:Arial;font-size:15pt;line-height:32px"><span>${text}</span></p>`;
    const html = '<div data-column-count="2" data-column-space="12" data-column-balance="false">' +
      para(penultimate) + para(terminal, `data-hwp-body-line="${line}"`) +
      '<div data-column-break="true"></div>' + para('Next column') + '</div>';
    let doc: any = { format: 'word-studio', version: 1, id: crypto.randomUUID(), title: 'Stored native end',
      comments: [], updatedAt: new Date().toISOString(), content: parseDocumentHtml(html),
      page: { size: 'Custom', width: 132.29167, height: 82.55, orientation: 'portrait', margin: 2.64583,
        header: '', footer: '', pageNumbers: false } };
    const inspect = async (candidate: typeof doc) => {
      const html = generateHTML(candidate.content, editorExtensions);
      const print = await createPrintLayout(html, candidate.page);
      try {
        const p = [...print.element.querySelectorAll('.print-content p')].find(p => p.textContent === penultimate);
        return { active: p?.querySelector('[data-hwp-column-tail]')?.textContent ?? null,
          standardParagraphs: [...print.element.querySelectorAll('.print-content p')].map(p => p.textContent),
          modelHasProof: html.includes('data-hwp-body-line="8 ') };
      } finally { print.dispose(); }
    };
    const moved = await inspect({ ...doc, content: parseDocumentHtml(html.replace('<div data-column-break="true"></div>', '')) });
    const edited = await inspect({ ...doc, content: parseDocumentHtml(html.replace(terminal, terminal + ' edited')) });
    const stages = [];
    for (let i = 0; i < 3; i++) {
      stages.push(await inspect(doc));
      if (i === 2) break;
      const load = await importFile(new File([await exportDocx(doc)], 'native-column.docx'));
      doc = { ...load.document, content: parseDocumentHtml(load.html!) };
    }
    return { stages, moved, edited, terminal, penultimate };
  });
  expect(result.moved.active).toBeNull();
  expect(result.edited.active).toBeNull();
  for (const stage of result.stages) {
    expect(stage.modelHasProof).toBe(true);
    expect(stage.active).toBe(' ');
    expect(stage.standardParagraphs).toEqual([result.penultimate, result.terminal, 'Next column']);
  }
});

test('fractional HWP leading keeps standard exact DOCX spacing and yields to an external edit', async ({ page }) => {
  await page.goto('/');
  const saves = await page.evaluate(async () => {
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const html = '<div data-column-count="2" data-column-space="12">' +
      '<p data-hwp-body-line="9 130 15.3333 19.9" style="font-family:Arial;font-size:11.5pt;line-height:19.9px">' +
      '<span style="font-size:12pt">Leading with the original exact measure</span></p>' +
      '<div data-column-break="true"></div><p>Following column</p></div>';
    let doc: any = { format: 'word-studio', version: 1, id: crypto.randomUUID(), title: 'Fractional leading',
      comments: [], updatedAt: new Date().toISOString(), content: parseDocumentHtml(html),
      page: { size: 'A4', orientation: 'portrait', margin: 12, header: '', footer: '', pageNumbers: false } };
    const result = [];
    for (let pass = 0; pass < 2; pass++) {
      const file = await exportDocx(doc);
      const opened = await importFile(new File([file], 'fractional.docx'));
      const content = parseDocumentHtml(opened.html!);
      const holder = document.createElement('div');
      holder.innerHTML = generateHTML(content, editorExtensions);
      const p = holder.querySelector('p[data-hwp-body-line]');
      result.push({ bytes: [...new Uint8Array(await file.arrayBuffer())], policy: p?.getAttribute('data-hwp-body-line'),
        pitch: p?.style.lineHeight, text: p?.textContent });
      doc = { ...opened.document, content };
    }
    return result;
  });
  for (const saved of saves) {
    expect(saved.policy).toBe('9 130 15.3333 19.9');
    expect(saved.pitch).toBe('19.9px');
    expect(saved.text).toBe('Leading with the original exact measure');
    const zip = await JSZip.loadAsync(Uint8Array.from(saved.bytes));
    const documentXml = await zip.file('word/document.xml')!.async('string');
    const paragraph = documentXml.match(/<w:p [^>]*wsb:bodyLine=[^>]*>.*?<\/w:p>/s)?.[0];
    expect(paragraph).toMatch(/<w:spacing[^>]*\bw:line="299"[^>]*\bw:lineRule="exact"/);
  }
  const changedZip = await JSZip.loadAsync(Uint8Array.from(saves[1].bytes));
  const documentXml = await changedZip.file('word/document.xml')!.async('string');
  expect(documentXml.match(/w:line="299"/g)).toHaveLength(1);
  changedZip.file('word/document.xml', documentXml.replace('w:line="299"', 'w:line="300"'));
  const data = [...await changedZip.generateAsync({ type: 'uint8array' })];
  const changed = await page.evaluate(async bytes => {
    const { importFile } = await import('/src/document-io.ts');
    const altered = await importFile(new File([Uint8Array.from(bytes)], 'external.docx'));
    const holder = document.createElement('div');
    holder.innerHTML = altered.html!;
    const original = [...holder.querySelectorAll<HTMLElement>('p')].find(p => p.textContent?.startsWith('Leading with'));
    return { pitch: original?.style.lineHeight, stored: original?.dataset.hwpBodyLine, text: original?.textContent };
  }, data);
  expect(changed).toEqual({ pitch: '15pt', stored: undefined, text: 'Leading with the original exact measure' });
});

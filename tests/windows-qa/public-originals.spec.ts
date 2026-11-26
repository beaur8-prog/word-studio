import { createHash } from 'node:crypto';
import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { readFile, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

type Input = { id: string; ext: 'docx' | 'hwp' | 'hwpx'; verified: boolean; path?: string };
const inventory = process.env.WINDOWS_PUBLIC_QA_INVENTORY;
if (!inventory) throw new Error('The public QA inventory is required');
const cases: Input[] = JSON.parse(readFileSync(inventory, 'utf8'));

for (const item of cases) test('Windows public original ' + item.id, async ({page}, info) => {
  test.skip(!item.verified, 'Public input was unavailable or failed the byte pin');
  test.setTimeout(process.env.WINDOWS_STAGED_ORIGINAL === item.id ? 1800000 : 900000);
  const source = await readFile(item.path!);
  const diagnostic = process.env.WINDOWS_QA_PAGE_HEAD_TRACE === '1';
  const selectedFormat = process.env.WINDOWS_DIAGNOSTIC_FORMAT;
  const formats: ('hwp' | 'hwpx')[] = selectedFormat === 'hwp' || selectedFormat === 'hwpx' ? [selectedFormat] :
    diagnostic || process.env.WINDOWS_PAGE_TWELVE_TRIAL === '1' || process.env.WINDOWS_PAGE_TWELVE_PRODUCT === '1' ? ['hwpx'] : ['hwp', 'hwpx'];
  if (diagnostic) page.on('console', message => {
    if (message.text().startsWith('QA_PAGE_HEAD_')) console.log(message.text());
  });
  await page.goto('/');
  if (process.env.WINDOWS_NATIVE_EXPORT_FLOW) await page.locator('#root .tiptap[role="textbox"]').first().waitFor();
  await page.exposeFunction('logOriginalStage', (stage: string, pages: number) => {
    console.log('WINDOWS_STAGED_PROGRESS ' + item.id + ' ' + stage + ' ' + pages);
  });
  await page.exposeFunction('logOriginalForward', (metrics: object) => {
    console.log('WINDOWS_FORWARD_RESULT ' + JSON.stringify({ id: item.id, ext: item.ext, ...metrics }));
  });
  await page.exposeFunction('writeActualDocx', async (stage: number, encoded: string) => {
    await writeFile(info.outputPath('textbox-save-' + item.id + '-' + stage + '.docx'), Buffer.from(encoded, 'base64'));
  });
  const verifyGeometry = process.env.WINDOWS_VERIFY_NATIVE_GEOMETRY === '1';
  const watched: { ordinal: number; literal: string }[] = [];
  if (process.env.WINDOWS_NATIVE_EXPORT_FLOW) {
    const require = createRequire(import.meta.url), core = path.dirname(require.resolve('@rhwp/core'));
    const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
    await init({ module_or_path: await readFile(path.join(core, 'rhwp_bg.wasm')) });
    const original = new HwpDocument(new Uint8Array(source));
    try {
      for (const ordinal of process.env.WINDOWS_NATIVE_EXPORT_FLOW.split(',').map(Number)) {
        const literal = original.getTextRange(0, ordinal, 0, 0xffffffff).replace(/\s/gu, '');
        if (literal) watched.push({ ordinal, literal });
      }
    } finally { original.free(); }
  }

  await page.exposeFunction('writeActualHangul', async (format: 'hwp' | 'hwpx', encoded: string) => {
    await writeFile(info.outputPath('hancom-saved.' + format), Buffer.from(encoded, 'base64'));
  });
  const result = await page.evaluate(async ({name, data, ext, verifyPictures, verifyGeometry, watched, flowOnly, skipPrint, formats, stripEmptyTabs}) => {
    const qa = window as Window & {
      logOriginalForward(metrics: object): Promise<void>;
      logOriginalStage(stage: string, pages: number): Promise<void>;
      writeActualDocx(stage: number, encoded: string): Promise<void>;
      writeActualHangul(format: 'hwp' | 'hwpx', encoded: string): Promise<void>;
    };
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { wordLayoutOptions } = await import('/src/word-trailing-spaces.ts');
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp, exportEditableHwp } = await import('/src/hwp/editing.ts');
    const documentFile = new File([new Uint8Array(data)], name);
    const loadDocx = async (file: File) => {
      const opened = await importFile(file);
      return { ...opened.document, content: parseDocumentHtml(opened.html!) };
    };
    let nativeOriginalPages: number | null = null;
    const loadSource = async () => {
      if (ext === 'docx') return loadDocx(documentFile);
      const reader = new HwpReaderSession();
      try {
        nativeOriginalPages = (await reader.open(documentFile)).count;
        const editable = await reader.editable();
        if (stripEmptyTabs) {
          const model = new DOMParser().parseFromString(editable.html, 'text/html');
          for (const p of model.querySelectorAll('p[data-tab-stops]')) {
            if (!p.textContent?.trim() && !p.querySelector('img,br,svg,[data-docx-tab],[data-document-field]'))
              p.removeAttribute('data-tab-stops');
          }
          return importEditableHwp({ ...editable, html:model.body.innerHTML }, name);
        }
        return importEditableHwp(editable, name);
      } finally { reader.close(); }
    };
    const signature = (content: any) => {
      const paragraphs: string[] = [], media: string[] = [], cells: number[][] = [];
      const read = (node: any): string => {
        if (node.type === 'text') return node.text;
        if (node.type === 'documentTab') return '\t';
        if (node.type === 'hardBreak') return '\n';
        if (node.type === 'image' || node.type === 'inlineImage') media.push(node.attrs.src);
        if (node.type === 'tableCell' || node.type === 'tableHeader') cells.push([node.attrs.colspan, node.attrs.rowspan]);
        const content = (node.content ?? []).map(read).join('');
        if (node.type === 'paragraph' || node.type === 'heading') paragraphs.push(content);
        return content;
      };
      read(content);
      return [JSON.stringify(paragraphs), JSON.stringify(media), JSON.stringify(cells)];
    };
    let doc = await loadSource();
    await qa.logOriginalStage("reader", nativeOriginalPages ?? 0);
    const printPages = [], preservation = [];
    let saves = 0;
    let previous: string[] | undefined;
    for (let stage = 0; stage < 3; stage++) {
      const current = signature(doc.content);
      if (previous) preservation.push({ textTabAndBreakOrder: previous[0] === current[0], mediaOrderAndBytes: previous[1] === current[1], tableSpans: previous[2] === current[2] });
      previous = current;
      const html = generateHTML(doc.content, editorExtensions);
      if (!flowOnly && !skipPrint) {
        const print = await createPrintLayout(html, doc.page, undefined, undefined, wordLayoutOptions(doc));
        try { printPages.push(print.pages); }
        finally { print.dispose(); }
        await qa.logOriginalStage("print-" + stage, printPages[stage]);
      }
      if (stage < 2) {
        const blob = await exportDocx(doc);
        const bytes = new Uint8Array(await blob.arrayBuffer()), chunks = [];
        for (let i = 0; i < bytes.length; i += 8192) chunks.push(String.fromCharCode(...bytes.subarray(i, i + 8192)));
        await qa.writeActualDocx(stage + 1, btoa(chunks.join('')));
        saves++;
        await qa.logOriginalStage('save-' + (stage + 1), saves);
        doc = await loadDocx(new File([blob], 'again.docx'));
      }
    }

    const nativeFlow = [];
    if (watched.length) {
      const { wordContentLanguage } = await import('/src/types.ts');
      const html = generateHTML(doc.content, editorExtensions);
      const layout = await createPrintLayout(html, doc.page, () => {}, undefined, {
        paintHwpCellLines: false, hwpNativeExport: true, interactive: true, language: wordContentLanguage(doc.wordDefaultLanguage),
      });
      try {
        const range = document.createRange(), paragraphs = [...layout.element.querySelectorAll<HTMLParagraphElement>('.print-content > p')];
        const compact = (value: string) => value.replace(/\s/gu, '');
        const entries = paragraphs.map(p => {
          const paper = p.closest('.print-page')!, body = p.closest('.print-content')!;
          const origin = paper.getBoundingClientRect(), box = p.getBoundingClientRect();
          range.selectNodeContents(p);
          const ink = [...range.getClientRects()].filter(rect => rect.width > .1 && rect.height > 1);
          const lineY = [...new Set(ink.map(rect => Math.round((rect.top - origin.top) * 10) / 10))];
          return { literal: compact(p.textContent ?? ''), details: { page: Number((paper as HTMLElement).dataset.page),
            y: box.top - origin.top, height: box.height, bottom: box.bottom - origin.top,
            bodyBottom: body.getBoundingClientRect().bottom - origin.top, lineY,
            inkBottom: ink.length ? Math.max(...ink.map(rect => rect.bottom - origin.top)) : null,
            fragment: p.hasAttribute('data-print-paragraph-fragment'), authored: p.dataset.hwpSourceWrap?.includes('authoredBodyLines') ?? false,
            tailProof: p.dataset.hwpSourceWrap?.includes('authoredTailInk') ?? false,
            headProof: p.dataset.hwpSourceWrap?.includes('authoredPageHead') ?? false,
            nextType: p.nextElementSibling?.tagName, nextLength: compact(p.nextElementSibling?.textContent ?? '').length,
            fontSize: getComputedStyle(p).fontSize, lineHeight: getComputedStyle(p).lineHeight } };
        }).filter(row => row.literal);
        for (const {ordinal, literal} of watched) {
          const results = [];
          for (let i = 0; i < entries.length; i++) {
            const first = entries[i], next = entries[i+1];
            const details = (entry: typeof first) => ({ literalLength: entry.literal.length, ...entry.details });
            if (first.literal === literal) results.push([details(first)]);
            else if (first.details.fragment && next?.details.fragment && first.literal + next.literal === literal)
              results.push([details(first), details(next)]);
          }
          nativeFlow.push({ ordinal, length: literal.length, results });
        }
      } finally { layout.dispose(); }
    }

    const reverse = [];
    await qa.logOriginalForward({ nativeOriginalPages, printPages, preservation, saves });
    if (!flowOnly && ext !== 'docx') for (const format of formats) {
      let reader: InstanceType<typeof HwpReaderSession> | undefined;
      try {
        const saved = await exportEditableHwp(doc, format, () => {});
        if (verifyGeometry) {
          const bytes = new Uint8Array(await saved.blob.arrayBuffer()), chunks = [];
          for (let i = 0; i < bytes.length; i += 8192) chunks.push(String.fromCharCode(...bytes.subarray(i, i + 8192)));
          await qa.writeActualHangul(format, btoa(chunks.join('')));
        }
        reader = new HwpReaderSession();
        const pages = (await reader.open(new File([saved.blob], 'roundtrip.' + format))).count;
        let pictureReadback: { count: number; exactByteOrder: boolean } | undefined;
        if (verifyPictures) {
          const opened = importEditableHwp(await reader.editable(), 'roundtrip.' + format);
          const reopened = JSON.parse(signature(opened.content)[1]) as string[];
          const expected = JSON.parse(previous![1]) as string[];
          pictureReadback = { count: reopened.length, exactByteOrder: reopened.length === expected.length &&
            reopened.every((image, index) => image.split(',')[1] === expected[index].split(',')[1]) };
        }
        reverse.push({format, pages, valid: pages > 0, blankWarning: saved.warnings.some(warning => warning.startsWith('일부 빈 문단')), warningCount: saved.warnings.length, pictureReadback});
      } catch (error) {
        const reason = String(error);
        reverse.push({format, valid: false,
          cause: /일부 떠 있는 그림을 한글로 저장하지 못했습니다/.test(reason) ? 'floating-picture-association' :
            /시간이 초과/.test(reason) ? 'timeout' : 'export-or-reopen-error'});
      } finally { reader?.close(); }
      await qa.logOriginalStage('reverse-' + format, reverse.at(-1)?.pages ?? 0);
    }
    return { nativeOriginalPages, printPages, preservation, reverse, saves, nativeFlow };
  }, { name: item.id + '.' + item.ext, data: [...source], ext: item.ext, verifyGeometry, verifyPictures: process.env.WINDOWS_REQUIRE_REVERSE === '1' || process.env.WINDOWS_VERIFY_REVERSE_MEDIA === '1', watched, flowOnly: process.env.WINDOWS_NATIVE_FLOW_ONLY === '1', skipPrint: process.env.WINDOWS_NATIVE_SKIP_PRINT === '1', formats, stripEmptyTabs:process.env.WINDOWS_DIAGNOSTIC_DROP_EMPTY_TABS === '1' });
  if (process.env.WINDOWS_NATIVE_FLOW_ONLY !== '1' && process.env.WINDOWS_NATIVE_SKIP_PRINT !== '1') {
    expect(result.printPages).toHaveLength(3);
    expect(result.printPages.every(pages => pages > 0)).toBe(true);
  }
  expect(result.saves).toBe(2);
  const { saves: _, nativeFlow, ...metrics } = result;
  if (watched.length) console.log('WINDOWS_HANCOM_PRE_EXPORT_FLOW ' + JSON.stringify(nativeFlow));
  if (process.env.WINDOWS_AUTHORED_TAIL_PROOF === '1') {
    const tail = nativeFlow.find(entry => entry.ordinal === 80);
    expect(tail?.results).toHaveLength(1);
    expect(tail?.results[0]).toHaveLength(1);
    expect(tail?.results[0][0]).toMatchObject({ page: 7, literalLength: 60, fragment: false, tailProof: true });
  }
  console.log('WINDOWS_PUBLIC_RESULT ' + JSON.stringify({ id: item.id, ext: item.ext, ...metrics }));
  if (verifyGeometry) {
    const require = createRequire(import.meta.url), core = path.dirname(require.resolve('@rhwp/core'));
    const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
    await init({ module_or_path: await readFile(path.join(core, 'rhwp_bg.wasm')) });
    const comparison = [], latin = [], preciseMargins = [];
    for (const format of formats) {
      const saved = new HwpDocument(new Uint8Array(await readFile(info.outputPath('hancom-saved.' + format))));
      try {
        let tables = 0, outside = 0, maxOutside = 0, firstSevenCellRow: number | undefined;
        const signatures = new Map<number,string>();
        let inkOutside = 0;
        for (let i = 0; i < saved.pageCount(); i++) {
          const paper = JSON.parse(saved.getPageInfo(i));
          const controls = JSON.parse(saved.getPageControlLayout(i)).controls;
          for (const table of controls.filter(c => c.type === 'table')) {
            tables++;
            if (item.id === '06-mixed-document' && table.rowCount === 1 && table.colCount === 7 &&
                table.w > 500 && !table.cellPath?.length) firstSevenCellRow ??= i + 1;
            const over = table.x + table.w - paper.width;
            if (over > .2) { outside++; maxOutside = Math.max(maxOutside, over); }
            if (process.env.WINDOWS_GRID_COMPARISON && item.id === '06-mixed-document' &&
                table.colCount === 7 && [8, 10].includes(table.rowCount) && !table.cellPath?.length) {
              const cells = new Map<number,string>();
              const ink = JSON.parse(saved.getPageTextLayout(i)).runs;
              for (const run of ink) {
                if (run.secIdx !== table.secIdx || run.parentParaIdx !== table.paraIdx ||
                    run.controlIdx !== table.controlIdx || run.cellPath?.length !== 1 || !run.text.trim()) continue;
                cells.set(run.cellIdx, (cells.get(run.cellIdx) ?? '') + run.text.replace(/\s/gu, ''));
                if (run.x + run.w > paper.width + .2) inkOutside++;
              }
              expect(signatures.has(table.rowCount)).toBe(false);
              expect(cells.size).toBeGreaterThan(0);
              const signature = createHash('sha256').update(JSON.stringify([...cells].sort((a, b) => a[0] - b[0]))).digest('hex');
              signatures.set(table.rowCount, signature);
            }
          }
        }
        console.log('WINDOWS_NATIVE_GRID ' + JSON.stringify({ id:item.id, format, pages:saved.pageCount(),
          tables, outsideRight:outside, maxOutsidePx:Math.round(maxOutside*10)/10, firstSevenCellRow }));
        if (process.env.WINDOWS_HANCOM_FIXED_AGREEMENT === '1') {
          const original = new HwpDocument(new Uint8Array(source));
          try {
            const stats = require('../../scripts/windows-qa/hancom-page-agreement.cjs')(original, saved);
            console.log('WINDOWS_HANCOM_FIXED_AGREEMENT ' + JSON.stringify({ id:item.id, format, ...stats }));
          } finally { original.free(); }
        }
        if (item.id === '04-language-plan' && process.env.WINDOWS_LATIN_WITNESS) {
          const original = new HwpDocument(new Uint8Array(source));
          const witness = require('../../scripts/windows-qa/hancom-line-witness.cjs');
          try {
            const value = { format, pages: saved.pageCount(), ...witness(original, saved),
              pageHead: process.env.WINDOWS_AUTHORED_PAGE_HEAD
                ? require('../../scripts/windows-qa/hancom-page-head-witness.cjs')(original, saved) : undefined,
              authoredLine: process.env.WINDOWS_AUTHORED_LINE_WITNESS
                ? require('../../scripts/windows-qa/hancom-single-line-witness.cjs')(original, saved) : undefined };
            latin.push(value);
            console.log('WINDOWS_HANCOM_LINE_WITNESS ' + JSON.stringify(value));
            if (process.env.WINDOWS_PAGE_OFFSETS) {
              const measure = require('../../scripts/windows-qa/hancom-source-page-offsets.cjs');
              const offsets = measure(original, saved);
              console.log('WINDOWS_HANCOM_PAGE_OFFSETS ' + JSON.stringify({ format, ...offsets }));
              if (process.env.WINDOWS_AUTHORED_TAIL_PROOF === '1') {
                const tail = offsets.bodyTail;
                console.log('WINDOWS_HANCOM_NATIVE_TAIL ' + JSON.stringify({ format, ...tail }));
                expect(tail?.saved?.map(line => [line.page, line.glyphs]).reduce((totals, [page, glyphs]) => {
                  totals[page] = (totals[page] ?? 0) + glyphs;
                  return totals;
                }, {})).toEqual({ 7: 60 });
              }
            }
          } finally { original.free(); }
        }
        if (item.id === '06-mixed-document' && process.env.WINDOWS_MARGIN_WITNESS) {
          const original = new HwpDocument(new Uint8Array(source));
          const witness = require('../../scripts/windows-qa/hancom-line-witness.cjs');
          try {
            const {marginLeft, marginRight} = JSON.parse(saved.getPageDef(0));
            const table = process.env.WINDOWS_TABLE_INSET_WITNESS
              ? require('../../scripts/windows-qa/hancom-table-insets.cjs')(original, saved) : undefined;
            const value = { format, pages: saved.pageCount(), outsideRight: outside, tableCount: tables,
              marginLeft, marginRight, ...witness(original, saved, 39, null), table,
              terminal: process.env.WINDOWS_TERMINAL_WITNESS
                ? require('../../scripts/windows-qa/hancom-terminal-witness.cjs')(original, saved) : undefined,
              compact: process.env.WINDOWS_COMPACT_ROW_WITNESS
                ? require('../../scripts/windows-qa/hancom-compact-row-witness.cjs')(original, saved) : undefined,
              adjacent: process.env.WINDOWS_ADJACENT_TABLE_WITNESS
                ? require('../../scripts/windows-qa/hancom-adjacent-tables-witness.cjs')(original, saved) : undefined };
            preciseMargins.push(value);
            console.log('WINDOWS_HANCOM_MARGIN_WITNESS ' + JSON.stringify(value));
            if (process.env.WINDOWS_HANCOM_SINGLE_BODY_LINE === '1') {
              const line = witness(original, saved, 122, null);
              console.log('WINDOWS_HANCOM_SINGLE_BODY_LINE ' + JSON.stringify({ format, ...line }));
              expect(line).toMatchObject({ page: 11, lines: 1 });
            }
            if (process.env.WINDOWS_PAGE_OFFSETS) {
              const measure = require('../../scripts/windows-qa/hancom-source-page-offsets.cjs');
              console.log('WINDOWS_HANCOM_PAGE_OFFSETS ' + JSON.stringify({ format, ...measure(original, saved) }));
            }
            if (process.env.WINDOWS_PAGE_TWELVE_TABLE === '1') {
              const packaged = format === 'hwp' ? new HwpDocument(saved.exportHwpx()) : undefined;
              try {
                const measure = require('../../scripts/windows-qa/hancom-table-fragment-position.cjs');
                const items = measure(original, saved, packaged);
                console.log('WINDOWS_HANCOM_PAGE_TWELVE_TABLE ' + JSON.stringify({ format, items }));
                if (format === 'hwpx' && process.env.WINDOWS_PAGE_TWELVE_REQUIRED === '1') {
                  const matches = items.filter(item => item.source.glyphs === 1080);
                  expect(matches).toHaveLength(1);
                  expect(matches[0].saved).toMatchObject({ page: 12, glyphs: 1080 });
                }
                if (format === 'hwpx' && process.env.WINDOWS_PAGE_TWELVE_TRIAL === '1') {
                  const { elementRanges } = await import('../../src/hwp/markup-ranges');
                  const tryInsets = require('../../scripts/windows-qa/hancom-table-insets-trial.cjs');
                  console.log('WINDOWS_HANCOM_PAGE_TWELVE_PLACEMENT ' + JSON.stringify(tryInsets.placement(HwpDocument, original, saved, elementRanges)));
                }
              } finally { packaged?.free(); }
            }
          } finally { original.free(); }
        }
        if (process.env.WINDOWS_GRID_COMPARISON) {
          expect([...signatures.keys()].sort((a, b) => a - b)).toEqual([8, 10]);
          comparison.push({ format, pages: saved.pageCount(), tables, outside, inkOutside,
            cellText: [8, 10].map(rows => signatures.get(rows)).join(',') });
        }
      } finally { saved.free(); }
    }
    if (process.env.WINDOWS_GRID_COMPARISON)
      await writeFile(process.env.WINDOWS_GRID_COMPARISON, JSON.stringify(comparison));
    if (process.env.WINDOWS_LATIN_WITNESS && latin.length)
      await writeFile(process.env.WINDOWS_LATIN_WITNESS, JSON.stringify(latin));
    if (process.env.WINDOWS_MARGIN_WITNESS && preciseMargins.length)
      await writeFile(process.env.WINDOWS_MARGIN_WITNESS, JSON.stringify(preciseMargins));
  }
  if (process.env.WINDOWS_REQUIRE_REOPEN === '1') {
    expect(result.reverse.map(out => ({ format: out.format, valid: out.valid })))
      .toEqual(formats.map(format => ({ format, valid: true })));
  }
  if (process.env.WINDOWS_EXPECT_ONE_SHEET === '1') {
    expect(result.printPages).toEqual([1, 1, 1]);
    expect(result.reverse.map(out => out.pages)).toEqual([1, 1]);
    expect(result.reverse.map(out => out.pictureReadback)).toEqual([
      { count:3, exactByteOrder:true }, { count:3, exactByteOrder:true },
    ]);
  }
  if (process.env.WINDOWS_REQUIRE_REVERSE === '1') {
    expect(result.reverse).toHaveLength(2);
    expect(result.reverse.map(out => ({ format: out.format, valid: out.valid, pictures: out.pictureReadback?.count }))).toEqual([
      { format: 'hwp', valid: true, pictures: 543 }, { format: 'hwpx', valid: true, pictures: 543 },
    ]);
    if (process.env.WINDOWS_EXACT_REVERSE_PAGES) {
      expect(result.reverse.map(out => out.pages)).toEqual([20, 20]);
      expect(result.reverse.every(out => out.pictureReadback?.exactByteOrder)).toBe(true);
    }
  }
});

import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { readFile, writeFile } from 'node:fs/promises';

type Input = { id: string; ext: 'docx' | 'hwp' | 'hwpx'; verified: boolean; path?: string };
const inventory = process.env.WINDOWS_PUBLIC_QA_INVENTORY;
if (!inventory) throw new Error('The public QA inventory is required');
const cases: Input[] = JSON.parse(readFileSync(inventory, 'utf8'));

for (const item of cases) test('Windows public original ' + item.id, async ({page}, info) => {
  test.skip(!item.verified, 'Public input was unavailable or failed the byte pin');
  test.setTimeout(process.env.WINDOWS_STAGED_ORIGINAL === item.id ? 1800000 : 900000);
  const source = await readFile(item.path!);
  await page.goto('/');
  await page.exposeFunction('logOriginalStage', (stage: string, pages: number) => {
    console.log('WINDOWS_STAGED_PROGRESS ' + item.id + ' ' + stage + ' ' + pages);
  });
  await page.exposeFunction('logOriginalForward', (metrics: object) => {
    console.log('WINDOWS_FORWARD_RESULT ' + JSON.stringify({ id: item.id, ext: item.ext, ...metrics }));
  });
  await page.exposeFunction('writeActualDocx', async (stage: number, encoded: string) => {
    await writeFile(info.outputPath('textbox-save-' + item.id + '-' + stage + '.docx'), Buffer.from(encoded, 'base64'));
  });
  const result = await page.evaluate(async ({name, data, ext, verifyPictures}) => {
    const qa = window as Window & {
      logOriginalForward(metrics: object): Promise<void>;
      logOriginalStage(stage: string, pages: number): Promise<void>;
      writeActualDocx(stage: number, encoded: string): Promise<void>;
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
        return importEditableHwp(await reader.editable(), name);
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
      const print = await createPrintLayout(html, doc.page, undefined, undefined, wordLayoutOptions(doc));
      try { printPages.push(print.pages); }
      finally { print.dispose(); }
      await qa.logOriginalStage("print-" + stage, printPages[stage]);
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

    const reverse = [];
    await qa.logOriginalForward({ nativeOriginalPages, printPages, preservation, saves });
    if (ext !== 'docx') for (const format of ['hwp', 'hwpx'] as const) {
      let reader: InstanceType<typeof HwpReaderSession> | undefined;
      try {
        const saved = await exportEditableHwp(doc, format, () => {});
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
    return { nativeOriginalPages, printPages, preservation, reverse, saves };
  }, { name: item.id + '.' + item.ext, data: [...source], ext: item.ext, verifyPictures: process.env.WINDOWS_REQUIRE_REVERSE === '1' || process.env.WINDOWS_VERIFY_REVERSE_MEDIA === '1' });
  expect(result.printPages).toHaveLength(3);
  expect(result.printPages.every(pages => pages > 0)).toBe(true);
  expect(result.saves).toBe(2);
  const { saves: _, ...metrics } = result;
  console.log('WINDOWS_PUBLIC_RESULT ' + JSON.stringify({ id: item.id, ext: item.ext, ...metrics }));
  if (process.env.WINDOWS_REQUIRE_REOPEN === '1') {
    expect(result.reverse.map(out => ({ format: out.format, valid: out.valid }))).toEqual([
      { format: 'hwp', valid: true }, { format: 'hwpx', valid: true },
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

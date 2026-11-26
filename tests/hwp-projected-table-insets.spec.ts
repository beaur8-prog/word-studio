import { writeFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import { strFromU8, unzipSync } from 'fflate';

test('a paragraph-anchored table keeps its offset and vertical reserve when saved into Hancom flow', async ({ page }, info) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { exportEditableHwp } = await import('/src/hwp/editing.ts');
    const grid = (text: string, anchor?: object) => {
      const metadata = anchor ? ` data-table-anchor="${JSON.stringify(anchor).replaceAll('"', '&quot;')}"` : '';
      return `<table${metadata}><tr><td colwidth="580"><p>${text}</p></td></tr></table>`;
    };
    const anchor = { x: 2, y: .8, bottom: 1.88, reference: 'paragraph', horizontal: 'margin' };
    const fixture = '<p>Start</p>' + grid('Projected table', anchor) + '<p>Between</p>' +
      grid('Ordinary table') + '<p>Next</p>' +
      grid('Larger reserve', { ...anchor, y: 12 }) + '<p>End</p>';
    const { defaultPage } = await import('/src/types.ts');
    let doc = { format: 'word-studio' as const, version: 1 as const, title: 'Projected table fixture',
      updatedAt: new Date().toISOString(), comments: [], page: defaultPage, content: parseDocumentHtml(fixture) };
    const docx = [];
    for (let i = 0; i < 2; i++) {
      const saved = await exportDocx(doc);
      docx.push([...new Uint8Array(await saved.arrayBuffer())]);
      const reopened = await importFile(new File([saved], 'projected.docx'));
      doc = { ...reopened.document, content: parseDocumentHtml(reopened.html!) };
    }
    const exported = await exportEditableHwp(doc, 'hwpx');
    return { docx, hwpx: [...new Uint8Array(await exported.blob.arrayBuffer())] };
  });

  for (const [index, bytes] of result.docx.entries())
    await writeFile(info.outputPath(`textbox-save-projected-insets-${index + 1}.docx`), Buffer.from(bytes));
  const section = strFromU8(unzipSync(Uint8Array.from(result.hwpx))['Contents/section0.xml']);
  const placements = [...section.matchAll(/<hp:tbl\b[^>]*>[\s\S]*?<\/hp:tbl>/g)].map(([table]) => ({
    inline: /<hp:pos\b[^>]*\btreatAsChar="1"/.test(table),
    margins: [...table.match(/<hp:outMargin\b[^>]*\/>/)![0].matchAll(/\b(left|top|bottom)="(\d+)"/g)]
      .map(([, side, size]) => [side, Number(size)]),
    text: [...table.matchAll(/<hp:t>([^<]+)<\/hp:t>/g)].map(([, text]) => text).join(''),
  }));
  expect(placements).toEqual([
    { inline: true, margins: [['left', 150], ['top', 60], ['bottom', 141]], text: 'Projected table' },
    { inline: true, margins: [['left', 283], ['top', 283], ['bottom', 283]], text: 'Ordinary table' },
    { inline: true, margins: [['left', 283], ['top', 283], ['bottom', 283]], text: 'Larger reserve' },
  ]);
  expect(section).not.toContain('WORDSTUDIO-TABLE-');
});

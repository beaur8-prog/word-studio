import { expect, test } from '@playwright/test';
import JSZip from 'jszip';

test('standalone inline textbox carriers do not become blank editable paragraphs after saving', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const frame = { width: 240, height: 92, x: 0, y: 0, vertical: 'paragraph', position: 'inline',
      shape: 'rect', align: 'center', padding: [7, 8, 7, 8], border: '#3483c8', borderWidth: 1.0 };
    const authored = '<p data-word-line-rule="exact" style="line-height:0.05pt;margin:0"></p>';
    const proof = { fingerprint: hwpWrapFingerprint('가나다'), ends: [3],
      inlineTextBox: { width: 240, height: 92, textWidth: 224, pitch: 18, hanging: 0, size: 14, family: 'Example' } };
    const witness = JSON.stringify(proof).replaceAll('"', '&quot;');
    const box = '<div data-document-textbox="' + JSON.stringify(frame).replaceAll('"', '&quot;') +
      '"><div data-textbox-content="true"><p data-hwp-source-wrap="' + witness +
      '">가나다</p><p></p><p>one editable story</p></div></div>';
    let file = new File(['<p>Before</p>' + authored + box + '<p>After</p>'], 'inline.html');
    const body = [], saves = [], witnesses = [];
    for (let pass = 0; pass < 3; pass++) {
      const imported = await importFile(file);
      const model = parseDocumentHtml(imported.html!);
      const prose = (node: typeof model): string => (node.text ?? '') + (node.content ?? []).map(prose).join('');
      body.push(model.content?.map(node => ({ type: node.type, text: prose(node) })));
      const stored = model.content?.find(node => node.type === 'documentTextBox')?.content?.[0]?.attrs?.hwpSourceWrap;
      witnesses.push(typeof stored === 'string' ? JSON.parse(stored) : null);
      if (pass < 2) {
        const bytes = new Uint8Array(await (await exportDocx({ ...imported.document, content: model })).arrayBuffer());
        saves.push([...bytes]);
        file = new File([bytes], 'inline.docx');
      }
    }
    return { body, saves, witnesses, proof };
  });
  const first = [
    { type: 'paragraph', text: 'Before' },
    { type: 'paragraph', text: '' },
    { type: 'documentTextBox', text: '가나다one editable story' },
    { type: 'paragraph', text: 'After' },
  ];
  expect(result.body).toEqual([first, first, first]);
  expect(result.witnesses).toEqual([result.proof, result.proof, result.proof]);
  for (const bytes of result.saves) {
    const zip = await JSZip.loadAsync(Buffer.from(bytes));
    const xml = await zip.file('word/document.xml')!.async('string');
    expect(xml.match(/<wp:inline\b/g)).toHaveLength(1);
    expect(xml).toMatch(/WordStudioTextBoxAnchor[^<]*\/><w:spacing[^>]*w:lineRule="atLeast"/);
    expect(xml.match(/<w:txbxContent>/g)).toHaveLength(1);
    expect(xml.match(/wsa:wrap="/g)).toHaveLength(1);
  }
});

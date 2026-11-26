import { expect, test } from '@playwright/test';

type Location = 'body' | 'header' | 'footer';

const cases = [
  { name: 'larger gap across two saves', gap: 20, offset: 4, location: 'body' as const },
  { name: 'offset already covers the preceding gap', gap: 8, offset: 12, location: 'body' as const },
  { name: 'near the body edge', gap: 20, offset: 4, location: 'body' as const, nearEnd: true },
  { name: 'externally edited carrier', gap: 20, offset: 4, location: 'body' as const, edit: 'text' as const },
  { name: 'externally styled carrier', gap: 20, offset: 4, location: 'body' as const, edit: 'style' as const },
  { name: 'header tables', gap: 20, offset: 4, location: 'header' as const },
  { name: 'footer tables', gap: 20, offset: 4, location: 'footer' as const },
];

for (const scenario of cases) test(`adjacent table offset preserves ${scenario.name}`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ gap, offset, location, edit, nearEnd }) => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const anchor = JSON.stringify({ x: 4, y: offset, bottom: 0, reference: 'paragraph', horizontal: 'margin' }).replaceAll('"', '&quot;');
    const table = (text: string, attrs: string, height: number) =>
      `<table ${attrs}><colgroup><col style="width:460px"></colgroup><tr style="height:${height}px"><td style="padding:0"><p>${text}</p></td></tr></table>`;
    const pair = table('INLINE TABLE', `style="width:460px;margin-bottom:${gap}px"`, 42) +
      table('ANCHORED TABLE', `data-table-anchor="${anchor}" style="width:460px"`, 64);
    const body = '<p>START</p>' + (nearEnd ? table('PRECEDING TABLE', 'style="width:460px;margin-bottom:0"', 405) : '') +
      (location === 'body' ? pair : '') + '<p>END</p>';
    const props = { ...defaultPage, size: 'Custom' as const, width: 150, height: 150, margin: 10, pageNumbers: false };
    if (location === 'header') props.headerHtml = pair;
    if (location === 'footer') props.footerHtml = pair;
    let model: import('../src/types').StudioDocument = { format: 'word-studio', version: 1, title: 'Table layout', content: parseDocumentHtml(body),
      comments: [], page: props, updatedAt: '2026-01-01T00:00:00Z' };
    const measure = async () => {
      const print = await createPrintLayout(generateHTML(model.content, editorExtensions), model.page);
      try {
        const rect = (s: string) => {
          const target = [...print.element.querySelectorAll('table')].find(t => t.textContent?.includes(s));
          const sheet = target?.closest<HTMLElement>('.print-page');
          if (!target || !sheet) return null;
          const r = target.getBoundingClientRect(), top = sheet.getBoundingClientRect().top;
          return { page: Number(sheet.dataset.page), top: r.top - top, bottom: r.bottom - top };
        };
        const inlineTable = rect('INLINE TABLE'), anchoredTable = rect('ANCHORED TABLE');
        return { pages: print.pages, inlineTable, anchoredTable,
          gap: inlineTable && anchoredTable && inlineTable.page === anchoredTable.page ? anchoredTable.top - inlineTable.bottom : null,
          text: [...print.element.querySelectorAll('.print-page')].map(p => p.textContent).join('').replace(/\s+/g, ''),
          exactAnchor: model.content.content?.filter(n => n.type === 'table').at(-1)?.attrs?.tableAnchor ?? null };
      } finally { print.dispose(); }
    };
    const reference = await measure(), rounds = [];
    for (let i = 0; i < 2; i++) {
      const file = await exportDocx(model), zip = await JSZip.loadAsync(await file.arrayBuffer());
      const story = location === 'body' ? 'word/document.xml' : `word/${location}1.xml`;
      const xml = new DOMParser().parseFromString(await zip.file(story)!.async('string'), 'application/xml');
      const carriers = [...xml.getElementsByTagNameNS(W, 'pStyle')].filter(s => s.getAttributeNS(W, 'val') === 'WordStudioTableOffset');
      const positions = carriers.map(s => {
        const spacing = s.parentElement!.getElementsByTagNameNS(W, 'spacing')[0];
        return { before: Number(spacing.getAttributeNS(W, 'before')), line: Number(spacing.getAttributeNS(W, 'line')),
          rule: spacing.getAttributeNS(W, 'lineRule') };
      });
      if (edit === 'text' && i === 0) {
        const run = xml.createElementNS(W, 'w:r'), text = xml.createElementNS(W, 'w:t');
        text.textContent = 'RETAIN MANUAL EDIT'; run.append(text); carriers[0].parentElement!.parentElement!.append(run);
        zip.file(story, new XMLSerializer().serializeToString(xml));
      }
      if (edit === 'style' && i === 0) {
        const styles = new DOMParser().parseFromString(await zip.file('word/styles.xml')!.async('string'), 'application/xml');
        const carrierStyle = [...styles.getElementsByTagNameNS(W, 'style')].find(s => s.getAttributeNS(W, 'styleId') === 'WordStudioTableOffset')!;
        carrierStyle.getElementsByTagNameNS(W, 'pPr')[0].append(styles.createElementNS(W, 'w:keepNext'));
        zip.file('word/styles.xml', new XMLSerializer().serializeToString(styles));
      }
      const next = await importFile(new File([edit && i === 0 ? await zip.generateAsync({ type: 'arraybuffer' }) : file], 'again.docx'));
      model = { ...next.document, content: parseDocumentHtml(next.html!) };
      rounds.push({ positions, ...(await measure()) });
    }
    return { reference, rounds };
  }, { ...scenario, location: scenario.location as Location });

  if (scenario.location !== 'body') {
    expect(result.reference.text).toContain('INLINETABLEANCHOREDTABLE');
    for (const round of result.rounds) {
      expect(round.positions.every(p => p.before === 0)).toBe(true);
      expect(round.text).toEqual(result.reference.text);
      expect(round.pages).toEqual(result.reference.pages);
    }
    return;
  }
  expect(result.rounds[0].positions).toEqual([{ before: Math.max(0, scenario.gap - scenario.offset) * 15, line: scenario.offset * 15, rule: 'exact' }]);
  if (scenario.edit) {
    for (const round of result.rounds) {
      expect(round.exactAnchor).toBeNull();
      if (scenario.edit === 'text') expect(round.text).toContain('RETAINMANUALEDIT');
      else expect(round.gap!).toBeCloseTo(result.reference.gap!, 0);
    }
    return;
  }
  for (const round of result.rounds) {
    expect(round.positions).toEqual(result.rounds[0].positions);
    expect(round.exactAnchor).toEqual(result.reference.exactAnchor);
    expect(round.pages).toBe(result.reference.pages);
    expect(round.text).toBe(result.reference.text);
    expect(round.anchoredTable?.page).toBe(result.reference.anchoredTable?.page);
    if (result.reference.gap !== null) expect(round.gap!).toBeCloseTo(result.reference.gap, 0);
  }
  if (scenario.nearEnd) expect(result.reference.pages).toBe(2);
});

import { test, expect } from '@playwright/test';

test('table bottom spacing keeps the following paragraph on the same page through two DOCX saves', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { defaultPage } = await import('/src/types.ts');
    const cases = [];
    for (const gap of ['0px', '8px', '2mm']) {
      let doc: any = { format: 'word-studio', version: 1, id: crypto.randomUUID(), title: 'Table gap', comments: [], updatedAt: new Date().toISOString(),
        page: { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10, margins: undefined },
        content: parseDocumentHtml(`<table style="margin-bottom:${gap}"><tr style="height:520px"><td style="padding:0;border:none"><p style="margin:0;line-height:20px">TABLE</p></td></tr></table><p style="margin:0;line-height:80px">AFTER</p>`) };
      const states = [];
      for (let save = 0; save < 3; save++) {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const bodies = [...layout.element.querySelectorAll('.print-content')];
          const after = bodies.at(-1)!.querySelector(':scope > p')!;
          states.push({ text: bodies.map(p => p.textContent), afterTop: after.getBoundingClientRect().top - bodies.at(-1)!.getBoundingClientRect().top });
        }
        finally { layout.dispose(); }
        if (save < 2) {
          const next = await importFile(new File([await exportDocx(doc)], 'gap.docx'));
          doc = { ...next.document, content: parseDocumentHtml(next.html!) };
        }
      }
      cases.push({ gap, states });
    }
    return cases;
  });
  expect(result[0].states.map(s => s.text)).toEqual([['TABLEAFTER'], ['TABLEAFTER'], ['TABLEAFTER']]);
  for (const item of result.slice(1)) {
    expect(item.states.map(s => s.text)).toEqual([['TABLE', 'AFTER'], ['TABLE', 'AFTER'], ['TABLE', 'AFTER']]);
    for (const state of item.states) expect(state.afterTop).toBeCloseTo(0, 1);
  }
});

test('unstyled text and fields inherit their paragraph font ahead of cell defaults on DOCX roundtrip', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { defaultPage } = await import('/src/types.ts');
    let doc: any = { format: 'word-studio', version: 1, id: crypto.randomUUID(), title: 'Inherited fonts', comments: [], updatedAt: new Date().toISOString(), page: { ...defaultPage, pageNumbers: false },
      content: parseDocumentHtml('<p style="font-family:Times New Roman;font-size:20pt">PLAIN<span style="font-size:9pt">SMALL</span>TAIL</p>' +
        '<table><tr><td style="font-family:Arial;font-size:9pt"><p style="font-family:Georgia;font-size:18pt">CELL<span data-page-number="PAGE">1</span></p></td></tr></table>') };
    const states = [];
    for (let save = 0; save < 3; save++) {
      const root = document.createElement('div'); root.className = 'tiptap'; root.style.cssText = 'position:fixed;left:-10000px';
      root.innerHTML = generateHTML(doc.content, editorExtensions); document.body.append(root);
      const runs: Record<string, { font: string; size: number }> = {};
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) { const n = walker.currentNode, style = getComputedStyle(n.parentElement!); runs[n.textContent!] = { font: style.fontFamily.replace(/["']/g, ''), size: parseFloat(style.fontSize) }; }
      states.push(runs); root.remove();
      if (save < 2) { const next = await importFile(new File([await exportDocx(doc)], 'fonts.docx')); doc = { ...next.document, content: parseDocumentHtml(next.html!) }; }
    }
    return states;
  });
  for (const state of result) {
    for (const text of ['PLAIN', 'TAIL']) { expect(state[text].font).toBe('Times New Roman'); expect(state[text].size).toBeCloseTo(80 / 3, 1); }
    expect(state.SMALL.size).toBe(12);
    for (const text of ['CELL', '1']) { expect(state[text].font).toBe('Georgia'); expect(state[text].size).toBe(24); }
  }
});

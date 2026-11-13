import { test, expect } from '@playwright/test';

test('cell line metrics run while earlier layout transactions continue to arrive', async ({ page }) => {
  await page.goto('/');
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true }); await expect(editor).toBeVisible();
  const doc = await editor.evaluate(async (root: any) => {
    const { Plugin, PluginKey } = await import('/node_modules/@tiptap/pm/state/index.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts'), { defaultPage } = await import('/src/types.ts');
    const key = new PluginKey('layoutActivityQA'); (window as any).__qaActivityKey = key;
    (window as any).__qaActivityObservation = null;
    root.editor.registerPlugin(new Plugin({ key, view(view) {
      let frame = 0, ticks = 0;
      const schedule = () => { cancelAnimationFrame(frame); frame = requestAnimationFrame(() => {
        if (view.state.doc.textContent.includes('STARVED CELL')) {
          ticks++;
          if (ticks === 15) {
            const metrics = view.state.plugins.find(plugin => plugin.key.startsWith('hwpCellLineMetrics'));
            (window as any).__qaActivityObservation = { ticks, decorations: metrics?.getState(view.state)?.find().length || 0,
              terminalParagraphs: view.dom.querySelectorAll('[data-hwp-cell-terminal]').length };
          }
        }
        if (ticks < 30) view.dispatch(view.state.tr.setMeta('qa-layout-activity', ticks).setMeta('addToHistory', false));
      }); };
      schedule(); return { update: schedule, destroy() { cancelAnimationFrame(frame); } };
    } }), (plugin: any, previous: any[]) => [plugin, ...previous]);
    return { format: 'word-studio', version: 1, id: 'cell-metrics-fairness', title: 'Cell metrics fairness', comments: [], updatedAt: new Date().toISOString(), page: { ...defaultPage, pageNumbers: false },
      content: parseDocumentHtml(`<table><tr><td colwidth="300" style="width:300px;padding:0"><p data-hwp-cell-line="1 Percent 16 32 16 32" style="font-family:Arial;font-size:12pt;line-height:32px;margin:0">STARVED CELL ${'ordinary cell text must keep its final line height. '.repeat(8)}</p></td></tr></table>`) };
  });
  try {
    await page.locator('input[type=file]').first().setInputFiles({ name: 'cell-metrics-fairness.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(doc)) });
    await expect(editor).toContainText('STARVED CELL');
    await expect.poll(() => page.evaluate(() => (window as any).__qaActivityObservation), { timeout: 10000 }).toBeTruthy();
    const observed = await page.evaluate(() => (window as any).__qaActivityObservation);
    expect(observed.ticks).toBe(15); expect(observed.decorations).toBeGreaterThan(0); expect(observed.terminalParagraphs).toBe(1);
  } finally { await editor.evaluate((root: any) => root.editor.unregisterPlugin((window as any).__qaActivityKey)); }
});


test('section paper geometry is available while earlier view updates continue', async ({ page }) => {
  await page.goto('/');
  await expect(page.getByRole('textbox', { name: '문서 편집 영역', exact: true })).toBeVisible();
  const observed = await page.evaluate(async () => {
    const { SectionNodeView } = await import('/src/section-node-view.ts');
    const { defaultPage } = await import('/src/types.ts');
    const editor = (document.querySelector('.tiptap[contenteditable=true]') as any).editor;
    const page = { ...defaultPage, margins: { top: 30, right: 20, bottom: 15, left: 20 }, pageNumbers: false, header: 'SOURCE HEADER' };
    const node = editor.schema.nodes.sectionContent.create({ page: JSON.stringify(page) }, editor.schema.nodes.paragraph.create());
    const host = document.createElement('div'); host.className = 'document-stack'; host.dataset.pages = '3';
    host.style.cssText = 'position:fixed;left:-100000px;top:0;width:800px;visibility:hidden'; document.body.append(host);
    let section: InstanceType<typeof SectionNodeView>, frame = 0, ticks = 0;
    try {
      const result = new Promise<any>(resolve => {
        const update = () => { frame = requestAnimationFrame(update); ticks++; section.update(node);
          if (ticks === 15) resolve({ ticks, height: section.dom.dataset.pageHeight, top: section.dom.dataset.margin,
            bottom: section.dom.dataset.bottomMargin, papers: section.dom.querySelectorAll('.section-paper').length,
            paddingTop: getComputedStyle(section.contentDOM).paddingTop });
        };
        frame = requestAnimationFrame(update);
      });
      section = new SectionNodeView(node); section.dom.dataset.pages = '3'; section.contentDOM.innerHTML = '<p>SECTION CONTENT</p>'; host.append(section.dom);
      return await result;
    } finally { cancelAnimationFrame(frame); section!.destroy(); host.remove(); }
  });
  expect(observed.ticks).toBe(15); expect(Number(observed.height)).toBeCloseTo(297 * 96 / 25.4, 1);
  expect(Number(observed.top)).toBeCloseTo(30 * 96 / 25.4, 1); expect(Number(observed.bottom)).toBeCloseTo(15 * 96 / 25.4, 1);
  expect(observed.papers).toBe(3); expect(parseFloat(observed.paddingTop)).toBeCloseTo(30 * 96 / 25.4, 1);
});

'use strict';
// Compare the real editor with print using complete, uniquely matching
// paragraphs. Split/duplicate paragraphs remain explicit exclusions; this
// does not replace independent source/PDF correspondence or fixed anchors.
const fs = require('node:fs/promises'), path = require('node:path'), os = require('node:os'), z = require('node:zlib');
const { chromium } = require('@playwright/test');
(async () => {
  const [models = 'analysis/hwp-12h-2026-10-28/variant-row-boundary/converted', evidence = 'analysis/hwp-12h-2026-10-28', label = 'editor-print', ...requested] = process.argv.slice(2);
  const ids = requested.length ? requested : ['02-work-rules', '04-language-plan', '06-mixed-document', '07-english-exam'];
  await fs.mkdir(evidence, { recursive: true });
  const temporary = await fs.mkdtemp(path.join(os.tmpdir(), 'word-editor-print-'));
  const browser = await chromium.launch({ channel: process.env.PLAYWRIGHT_CHANNEL || 'chrome', headless: true });
  try {
    const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } }); page.setDefaultTimeout(180000);
    await page.goto(process.env.HWP_AUDIT_URL || 'http://127.0.0.1:5191');
    const summaries = [];
    for (const id of ids) {
      const model = JSON.parse(z.gunzipSync(await fs.readFile(path.join(models, id + '.wordstudio.gz'))));
      // Playwright's in-memory upload stops at 50 MB; backups with embedded
      // image layers can be larger and must exercise the actual file path.
      const file = path.join(temporary, id + '.wordstudio'); await fs.writeFile(file, JSON.stringify(model));
      await page.locator('input[type=file]').first().setInputFiles(file);
      await page.getByRole('textbox', { name: '문서 제목' }).waitFor();
      await page.waitForFunction(title => document.querySelector('[aria-label="문서 제목"]')?.value === title, model.title);
      await page.evaluate(() => document.fonts.ready);
      let previous = '', stable = 0;
      for (let i = 0; i < 100 && stable < 8; i++) {
        const next = await page.getByRole('contentinfo').innerText(); stable = next === previous ? stable + 1 : 0; previous = next;
        await page.waitForTimeout(200);
      }
      const result = await page.evaluate(async model => {
        const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { editorExtensions } = await import('/src/extensions.ts');
        const { createPrintLayout } = await import('/src/print-layout.ts'), { documentScale } = await import('/src/document-scale.ts');
        const editor = document.querySelector('[role="textbox"][contenteditable]'), stack = editor.closest('.document-stack');
        const norm = text => text.normalize('NFKC').replace(/\s/g, '');
        const firstText = p => {
          const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
          while (walker.nextNode()) {
            const text = walker.currentNode, at = text.textContent.search(/\S/u);
            if (at < 0 || text.parentElement.closest('.ProseMirror-widget,[data-flow-widget]')) continue;
            const range = document.createRange(); range.setStart(text, at); range.setEnd(text, at + [...text.textContent.slice(at)][0].length);
            const rect = range.getBoundingClientRect(); if (rect.height > 0) return rect;
          }
          return p.getBoundingClientRect();
        };
        const contexts = [...editor.querySelectorAll('.section-editable')]; if (!contexts.length) contexts.push(editor);
        const editing = []; let offset = 0;
        for (const context of contexts) {
          const sheet = context.closest('[data-page-height]') || stack, origin = context.getBoundingClientRect().top;
          const scale = documentScale(context), stride = Number(sheet.dataset.pageHeight) + 24;
          for (const p of context.querySelectorAll('p,h1,h2,h3,h4,h5,h6')) {
            if (p.closest('.ProseMirror-widget,[data-document-note]')) continue;
            const y = (firstText(p).top - origin) / scale, local = Math.max(0, Math.floor((y + .5) / stride));
            editing.push({ text: norm(p.textContent), page: offset + local + 1, y: y - local * stride, blockY: (p.getBoundingClientRect().top - origin) / scale });
          }
          offset += Number(sheet.dataset.pages) || 0;
        }
        const layout = await createPrintLayout(generateHTML(model.content, editorExtensions), model.page);
        try {
          const printing = [...layout.element.querySelectorAll('.print-content p,.print-content h1,.print-content h2,.print-content h3,.print-content h4,.print-content h5,.print-content h6')].map(p => {
            const paper = p.closest('.print-page');
            return { text: norm(p.textContent), page: Number(paper.dataset.page), y: firstText(p).top - paper.getBoundingClientRect().top, blockY: p.getBoundingClientRect().top - paper.getBoundingClientRect().top };
          });
          const grouped = values => { const map = new Map(); for (const p of values) { const a = map.get(p.text) || []; a.push(p); map.set(p.text, a); } return map; };
          const a = grouped(editing), b = grouped(printing), checks = []; let excluded = 0;
          for (const [text, values] of a) {
            if (text.length < 10) continue;
            if (values.length !== 1 || b.get(text)?.length !== 1) { excluded += values.length; continue; }
            checks.push({ text, editor: values[0], print: b.get(text)[0] });
          }
          return { editorPages: Number(stack.dataset.pages), printPages: layout.pages, matched: checks.length, excluded,
            mismatches: checks.filter(p => p.editor.page !== p.print.page), samePageY: checks.filter(p => p.editor.page === p.print.page && Math.abs(p.editor.y - p.print.y) > 2), checks };
        } finally { layout.dispose(); }
      }, model);
      await fs.writeFile(path.join(evidence, `${label}-${id}.json`), JSON.stringify(result, null, 2));
      const summary = { id, editorPages: result.editorPages, printPages: result.printPages, matched: result.matched, excluded: result.excluded, mismatchedStartPages: result.mismatches.length, samePageYDifferencesAbove2px: result.samePageY.length };
      summaries.push(summary); console.log(JSON.stringify(summary));
    }
    await fs.writeFile(path.join(evidence, label + '-summary.json'), JSON.stringify(summaries, null, 2));
  } finally { await browser.close(); await fs.rm(temporary, { recursive: true, force: true }); }
})().catch(error => { console.error(error); process.exitCode = 1; });

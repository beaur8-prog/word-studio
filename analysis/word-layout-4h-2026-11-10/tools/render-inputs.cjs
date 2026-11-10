const fs = require('node:fs/promises');
const path = require('node:path');
const zlib = require('node:zlib');
const { createRequire } = require('node:module');
const { pathToFileURL } = require('node:url');
const { createHash } = require('node:crypto');
const root = path.resolve(__dirname, '../../..');
const evidence = path.dirname(__dirname);
const previous = path.join(root, 'analysis/hwp-cd-10h-2026-11-07');
const req = createRequire(root + '/package.json');
const variant = process.env.LAYOUT_VARIANT || 'current';
if (!['current','baseline','frozen-product'].includes(variant)) throw Error('Historical experiments are recorded in Git/progress. Use current or frozen-product with the final source.');
const stage = process.env.LAYOUT_STAGE || variant;
const port = Number(process.env.LAYOUT_PORT || 5441);
const selected = (process.env.LAYOUT_IDS || 'mss-funding-2026-0,mods-household-1,mois-2025-0').split(',');
const sha = bytes => createHash('sha256').update(bytes).digest('hex');
(async () => {
  const out = path.join(evidence, 'render-' + stage);
  await fs.mkdir(out, { recursive: true });
  const comparison = JSON.parse(await fs.readFile(previous + '/final-word-studio-comparison.json'));
  const corpus = process.env.LAYOUT_CORPUS || 'government';
  const wordInputs = corpus === 'word' ? JSON.parse(await fs.readFile(root + '/analysis/word-pagination-3h-2026-11-05/baseline.json')).rows.map(row => ({ id: row.id, inputFile: row.path, inputSHA256: row.sha256, split: 'existing-word-control', studioPages: row.wordReference.pages, wordPages: row.wordReference.pages })) : [];
  const suppliedInputs = process.env.LAYOUT_INPUT_MANIFEST ? JSON.parse(await fs.readFile(process.env.LAYOUT_INPUT_MANIFEST)).rows : null;
  const inputs = (suppliedInputs || (corpus === 'word' ? wordInputs : comparison.rows)).filter(row => selected.includes(row.id) || selected.includes('all'));
  const { createServer } = await import(pathToFileURL(req.resolve('vite')));
  const diagnostic = { name: 'diagnostic-full-line-capacity', enforce: 'pre', async load(id) {
    if(variant==='frozen-product'){const rel=path.relative(root,id.split('?')[0]);if(['src/document-io.ts','src/print-layout.ts','src/paragraph-fit.ts','src/pagination.ts','src/types.ts','src/section-node-view.ts','src/App.tsx','src/android/EditorCanvas.tsx','src/row-split-policy.ts','src/table-fragments.ts','src/editor-cell-pagination.ts'].includes(rel))return require('node:child_process').execFileSync('git',['show','5cb4246:'+rel],{cwd:root,encoding:'utf8'});}

  }};
  const server = await createServer({ root, plugins: [diagnostic], configFile: root + '/vite.config.ts',
    server: { host: '127.0.0.1', port, strictPort: true, hmr: false, watch: { ignored: ['**'] } }, logLevel: 'error' });
  await server.listen();
  const { chromium } = req('@playwright/test');
  const browser = await chromium.launch({ channel: 'chrome', headless: true });
  const code = {};
  for (const file of ['src/document-io.ts','src/paragraph-fit.ts','src/print-layout.ts','src/table-fragments.ts','src/row-split-policy.ts','src/pagination.ts','src/editor-cell-pagination.ts','src/table-row-resize.ts']) code[file] = sha(variant === 'frozen-product' ? require('node:child_process').execFileSync('git',['show','5cb4246:'+file],{cwd:root}) : await fs.readFile(root + '/' + file));
  const rows = [], errors = [];
  try {
    for (const input of inputs) {
      const started = Date.now();
      const page = await browser.newPage({ viewport: { width: 1440, height: 1000 }, deviceScaleFactor: 1 });
      page.on('pageerror', error => errors.push({ id: input.id, error: String(error) }));
      try {
        const bytes = await fs.readFile(path.resolve(previous, input.inputFile));
        if (sha(bytes) !== input.inputSHA256) throw Error('Frozen Word input changed');
        await page.goto('http://127.0.0.1:' + port + '/');
        await page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).waitFor();
        const state = await page.evaluate(async bytes => {
          const { importFile } = await import('/src/document-io.ts');
          const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
          const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
          const { createPrintLayout } = await import('/src/print-layout.ts');
          const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
          const loaded = await importFile(new File([new Uint8Array(bytes)], 'comparison.docx'));
          const doc = { ...loaded.document, content: loaded.html === undefined ? loaded.document.content : parseDocumentHtml(loaded.html) };
          const logical = n => n.type === 'hardBreak' ? '\n' : n.type === 'documentTab' ? '\t' : (n.text || '') + (n.content || []).map(logical).join('');
          const text = logical(doc.content);
          const bodySHA256 = [...new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text)))].map(n => n.toString(16).padStart(2, '0')).join('');
          const source = document.createElement('div'); source.innerHTML = generateHTML(doc.content, editorExtensions);
          [...source.querySelectorAll('p,h1,h2,h3,h4,h5,h6')].forEach((p, key) => p.dataset.auditKey = String(key));
          const sourceRows = [...source.querySelectorAll('tr')].map(row=>({height:row.style.height,rule:row.dataset.wordRowRule,split:row.dataset.rowSplit,wordSplit:row.dataset.wordRowSplit,text:row.textContent.slice(0,90)}));
          window.auditLayout = await createPrintLayout(source.innerHTML, doc.page);
          const layout = window.auditLayout;
          layout.element.classList.replace('measuring', 'preview'); document.body.classList.add('studio-printing');
          const paragraphs = [...layout.element.querySelectorAll('[data-audit-key]')].map(p => {
            const box = p.getBoundingClientRect(), sheet = p.closest('.print-page').getBoundingClientRect(), content = p.closest('.print-content,.print-footnotes').getBoundingClientRect(), css = getComputedStyle(p);
            return { key: Number(p.dataset.auditKey), page: Number(p.closest('.print-page').dataset.page), text: p.textContent,
              x: box.left - sheet.left, y: box.top - sheet.top, height: box.height, bottomToBody: box.bottom - content.bottom,
              lineHeight: css.lineHeight, before: css.marginTop, after: css.marginBottom, wordLineRule: p.dataset.wordLineRule,
              cell: !!p.closest('td,th'), lineOffsets: paragraphLineOffsets(p), fragmentOffset: p.dataset.fragmentOffset,
              terminal: p.dataset.hwpCellTerminal, hwpCellLine: p.dataset.hwpCellLine };
          });
          return { pages: layout.pages, warnings: layout.warnings, textCharacters: text.length, bodySHA256, paragraphs, sourceRows, rowProbes: window.__wordRowProbes, settings:doc.page,
            pageTexts: [...layout.element.querySelectorAll('.print-content')].map(p => p.textContent),
            tableFragments: layout.element.querySelectorAll('.print-content table').length,
            imagePlacements: layout.element.querySelectorAll('.print-content img').length };
        }, [...bytes]);
        const pdf = out + '/' + input.id + '.pdf';
        await page.pdf({ path: pdf, preferCSSPageSize: true, printBackground: true });
        await fs.writeFile(out + '/' + input.id + '.json.gz', zlib.gzipSync(JSON.stringify(state)));
        const row = { id: input.id, split: input.split, inputSHA256: input.inputSHA256, priorStudioPages: input.studioPages,
          wordPages: input.wordPages, pages: state.pages, warnings: state.warnings, textCharacters: state.textCharacters,
          bodySHA256: state.bodySHA256, imagePlacements: state.imagePlacements, tableFragments: state.tableFragments,
          pdfSHA256: sha(await fs.readFile(pdf)), durationMs: Date.now() - started };
        rows.push(row); console.log(JSON.stringify(row));
      } catch (error) { errors.push({ id: input.id, error: String(error) }); console.error(input.id, String(error)); }
      finally { await page.close(); }
      await fs.writeFile(out + '/results.json', JSON.stringify({ utc: new Date().toISOString(), variant, stage, code,
        method: 'Exact frozen actual web Word input bytes; current Studio print; DPR1, 100% display; no HWP model substitution', rows, errors }, null, 2));
    }
  } finally { await browser.close(); await server.close(); }
  if (errors.length) process.exitCode = 1;
})().catch(error => { console.error(error); process.exitCode = 1; });

const fs = require('node:fs/promises');
const path = require('node:path');
const zlib = require('node:zlib');
const crypto = require('node:crypto');
const { createRequire } = require('node:module');
const { pathToFileURL } = require('node:url');
const root = path.resolve(__dirname, '../../..');
const evidence = path.dirname(__dirname);
const previous = path.join(root, 'analysis/word-layout-4h-2026-11-10');
const req = createRequire(root + '/package.json');
const id = process.env.PROBE_ID || 'mods-household-1';
const stage = process.env.PROBE_STAGE || 'baseline';
const port = Number(process.env.PROBE_PORT || 5461);
const sha = bytes => crypto.createHash('sha256').update(bytes).digest('hex');

(async () => {
  const manifest = JSON.parse(await fs.readFile(previous + '/government-final-comparison.json'));
  const row = manifest.rows.find(row => row.id === id);
  if (!row) throw Error('Unknown fixed development document');
  const bytes = await fs.readFile(path.resolve(previous, row.sourceDOCX));
  if (sha(bytes) !== row.inputSHA256) throw Error('Frozen input hash changed');
  const out = path.join(evidence, 'probe-' + stage);
  await fs.mkdir(out, { recursive: true });
  const { createServer } = await import(pathToFileURL(req.resolve('vite')));
  const candidateFiles={'src/pagination.ts':'pagination-merge-header-experiment.ts','src/editor-merged-fragments.ts':'editor-merged-fragments-experiment.ts'};
  const plugins=process.env.PROBE_MERGE_EXPERIMENT==='1'?[{name:'read-only-merge-header-candidate',enforce:'pre',async load(id){const rel=path.relative(root,id.split('?')[0]);if(candidateFiles[rel])return fs.readFile(evidence+'/'+candidateFiles[rel],'utf8')}}]:[];
  const server = await createServer({ root,plugins, configFile: root + '/vite.config.ts',
    server: { host: '127.0.0.1', port, strictPort: true, hmr: false, watch: { ignored: ['**'] } }, logLevel: 'error' });
  await server.listen();
  const { chromium } = req('@playwright/test');
  const browser = await chromium.launch({ channel: 'chrome', headless: true });
  const page = await browser.newPage({ viewport: { width: 1440, height: 1000 }, deviceScaleFactor: 1 });
  page.setDefaultTimeout(180000);
  const errors = [];
  page.on('pageerror', e => errors.push(String(e)));
  try {
    await page.goto('http://127.0.0.1:' + port);
    await page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).waitFor();
    const model = await page.evaluate(async bytes => {
      const { importFile } = await import('/src/document-io.ts');
      const { parseDocumentHtml } = await import('/src/extensions.ts');
      const loaded = await importFile(new File([new Uint8Array(bytes)], 'probe.docx'));
      return { ...loaded.document, content: loaded.html === undefined ? loaded.document.content : parseDocumentHtml(loaded.html) };
    }, [...bytes]);
    await page.locator('input[type=file]').first().setInputFiles({
      name: id + '.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(model))
    });
    await page.locator('.busy-indicator').waitFor({ state: 'hidden' });
    await page.evaluate(() => document.fonts.ready);
    let last = '', stable = 0;
    const convergence = [];
    for (let i = 0; i < 18 && stable < 4; i++) {
      await page.waitForTimeout(500);
      const value = await page.evaluate(() => {
        const editor = document.querySelector('.tiptap[contenteditable=true]');
        const box = editor.getBoundingClientRect();
        return { pages: editor.closest('.document-stack').dataset.pages, height: box.height,
          gaps: [...editor.querySelectorAll('.pagination-gap,.pagination-cell-gap')].map(g => g.style.height).join('|') };
      });
      const signature = JSON.stringify(value);
      stable = last === signature ? stable + 1 : 0; last = signature;
      convergence.push({ sample: i, pages: value.pages, height: value.height, stable });
    }
    const state = await page.evaluate(async model => {
      const editor = document.querySelector('.tiptap[contenteditable=true]');
      const view = editor.editor.view;
      const sheet = editor.closest('.document-stack');
      const origin = editor.getBoundingClientRect();
      const h = Number(sheet.dataset.pageHeight), stride = h + 24;
      const sections = [...editor.querySelectorAll('.document-section-content')];
      const rows = [...editor.querySelectorAll('tr:not(.ProseMirror-widget)')].filter(r => !r.closest('.section-papers,.pagination-headers'));
      const paragraphs = [...editor.querySelectorAll('p,h1,h2,h3,h4,h5,h6')].filter(p => !p.closest('.section-papers,.pagination-headers'));
      const coord = (el, root, page) => {
        const section = root === editor ? el.closest('.document-section-content') : null;
        const context = section?.querySelector('.section-editable') || root;
        const localStride = section ? Number(section.dataset.pageHeight) + 24 : stride;
        const before = section ? sections.slice(0, sections.indexOf(section)).reduce((n,s) => n + Number(s.dataset.pages || 1), 0) : 0;
        const r = el.getBoundingClientRect(), start = context.getBoundingClientRect();
        const y = r.top - start.top;
        const n = page ?? Math.max(1, Math.floor((y + .5) / localStride) + 1);
        return { page: page ?? before + n, x: r.left - start.left, y: page ? y : y - (n - 1) * localStride, height: r.height, width: r.width };
      };
      const describe = (el, root, page) => ({
        ...coord(el, root, page), tag: el.tagName, text: el.textContent,
        style: el.style.cssText, attrs: Object.fromEntries([...el.attributes].filter(a => a.name.startsWith('data-')).map(a => [a.name, a.value])),
        computed: Object.fromEntries(['height','minHeight','paddingTop','paddingBottom','borderTopWidth','borderBottomWidth','verticalAlign','marginTop','marginBottom','lineHeight'].map(k => [k, getComputedStyle(el)[k]]))
      });
      const ancestors = p => {
        const chain = [];
        for (let a = p.parentElement; a && a !== editor; a = a.parentElement)
          if (a.matches('tr,td,th,table,.tableWrapper')) chain.push({ ...describe(a, editor), rowKey: rows.indexOf(a), firstText: a.textContent.slice(0,100), text: undefined });
        return chain;
      };
      const edited = paragraphs.map((p, key) => ({ key, ...describe(p, editor), ancestors: ancestors(p) }));
      const displayedRows = rows.map((r, key) => ({ key, ...describe(r, editor), text: r.textContent.slice(0,120) }));
      const gaps = [...editor.querySelectorAll('.pagination-gap,.pagination-cell-gap')].map(g => ({
        ...describe(g, editor), classes: g.className,
        previous: g.previousElementSibling?.textContent?.slice(0,90),
        next: g.nextElementSibling?.textContent?.slice(0,90)
      }));
      editor.classList.add('pagination-measuring-cells');
      const naturalRows = rows.map((r, key) => ({ key, ...describe(r, editor), text: r.textContent.slice(0,120) }));
      editor.classList.remove('pagination-measuring-cells');
      const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
      const { editorExtensions } = await import('/src/extensions.ts');
      const { createPrintLayout } = await import('/src/print-layout.ts');
      const source = document.createElement('div');
      source.innerHTML = generateHTML(model.content, editorExtensions);
      [...source.querySelectorAll('p,h1,h2,h3,h4,h5,h6')].forEach((p,key) => p.dataset.auditKey = String(key));
      const layout = await createPrintLayout(source.innerHTML, model.page);
      const printed = [...layout.element.querySelectorAll('.print-content [data-audit-key]')].map(p => {
        const paper = p.closest('.print-page');
        return { key: Number(p.dataset.auditKey), ...describe(p, paper, Number(paper.dataset.page)) };
      });
      const printedRows = [...layout.element.querySelectorAll('.print-content tr')].map(r => {
        const paper = r.closest('.print-page');
        return { ...describe(r, paper, Number(paper.dataset.page)), text: r.textContent.slice(0,120) };
      });
      const summary = {
        pages: { editor: Number(sheet.dataset.pages), print: layout.pages },
        paragraphCounts: { editor: edited.length, print: printed.length },
        sections: sections.map(s => ({pages:Number(s.dataset.pages),height:Number(s.dataset.pageHeight),top:Number(s.dataset.margin),bottom:Number(s.dataset.bottomMargin)})),
        exactParagraphOrder: edited.map(p=>p.text).join('\u0000') === printed.map(p=>p.text).join('\u0000'),
        sheet: { height:h,top:Number(sheet.dataset.margin),bottom:Number(sheet.dataset.bottomMargin) },
        warnings:layout.warnings
      };
      window.probePrint = layout;
      return { summary, edited, displayedRows, naturalRows, gaps, printed, printedRows };
    }, model);
    state.inputSHA256 = row.inputSHA256; state.convergence = convergence; state.errors = errors;
    await fs.writeFile(out + '/' + id + '.json.gz', zlib.gzipSync(JSON.stringify(state)));
    await fs.writeFile(out + '/' + id + '-summary.json', JSON.stringify(state.summary, null, 2));
    console.log(JSON.stringify({ id, stage, ...state.summary, convergence, errors }));
  } finally { await browser.close(); await server.close(); }
})().catch(e => { console.error(e); process.exitCode = 1; });

'use strict';
// Investigate the production HWP -> HTML -> editor -> DOCX path. Public inputs
// and generated documents remain under ignored analysis/, never in a release.
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const crypto = require('node:crypto');
const { pathToFileURL } = require('node:url');
const { createRequire } = require('node:module');
const { execFileSync } = require('node:child_process');
const { gunzipSync } = require('node:zlib');
const ts = require('typescript');
const { chromium } = require('@playwright/test');
const root = path.resolve(__dirname, '..');
const output = path.resolve(process.argv[2] || path.join(root, 'analysis/hwp-conversion-study-2026-10-26'));
async function readJson(file) {
  const bytes = await fs.readFile(file).catch(async error => {
    if (error.code !== 'ENOENT') throw error;
    return gunzipSync(await fs.readFile(file + '.gz'));
  });
  return JSON.parse(bytes.toString('utf8'));
}

async function main() {
  await fs.mkdir(path.join(output, 'samples'), { recursive: true });
  const manifest = JSON.parse(await fs.readFile(path.join(root, 'tests/hwp-conversion-samples.json'), 'utf8'));
  await fs.writeFile(path.join(output, 'manifest.json'), JSON.stringify(manifest, null, 2));
  for (const sample of manifest.samples) for (const kind of ['hwp', 'pdf']) {
    const input = sample[kind]; if (!input) continue;
    const file = path.join(output, input.file);
    let bytes = await fs.readFile(file).catch(error => { if (error.code !== 'ENOENT') throw error; return null; });
    if (!bytes) {
      const response = await fetch(kind === 'hwp' ? sample.url : sample.referenceURL, { signal: AbortSignal.timeout(90000) });
      if (!response.ok) throw Error(`Download failed (${response.status}): ${sample.id} ${kind}`);
      bytes = Buffer.from(await response.arrayBuffer());
      if (crypto.createHash('sha256').update(bytes).digest('hex') !== input.sha256) throw Error('Downloaded digest mismatch: ' + sample.id);
      await fs.writeFile(file, bytes);
    }
    if (crypto.createHash('sha256').update(bytes).digest('hex') !== input.sha256) throw Error('Cached digest mismatch: ' + sample.id);
  }
  const directory = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(directory, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(directory, 'rhwp_bg.wasm')) });
  const sourceFile = path.join(root, 'src/hwp/conversion.ts');
  const compiled = ts.transpileModule(await fs.readFile(sourceFile, 'utf8'), {
    compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.CommonJS, esModuleInterop: true },
  }).outputText;
  const bridge = { exports: {} };
  new Function('require', 'module', 'exports', compiled)(createRequire(sourceFile), bridge, bridge.exports);
  const metadataFile = path.join(root, 'src/hwp/layout-metadata.ts'), metadataBridge = { exports: {} };
  new Function('require', 'module', 'exports', ts.transpileModule(await fs.readFile(metadataFile, 'utf8'), {
    compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.CommonJS, esModuleInterop: true },
  }).outputText)(createRequire(metadataFile), metadataBridge, metadataBridge.exports);
  await fs.mkdir(path.join(output, 'converted'), { recursive: true });
  await fs.mkdir(path.join(output, 'evidence'), { recursive: true });
  const codeHash = crypto.createHash('sha256');
  const codeFiles = execFileSync('git', ['ls-files', '-z', '--cached', '--others', '--exclude-standard'], { cwd: root, encoding: 'utf8' })
    .split('\0').filter(name => /^(?:src\/|electron\/|vendor\/|package(?:-lock)?\.json$)/.test(name)).sort();
  for (const file of codeFiles) codeHash.update(file + '\0').update(await fs.readFile(path.join(root, file)));
  const sourceTreeSHA256 = codeHash.digest('hex');
  const workingTreeDirty = !!execFileSync('git', ['status', '--porcelain'], { cwd: root, encoding: 'utf8' }).trim();
  const results = [];
  for (const sample of manifest.samples) {
    const bytes = await fs.readFile(path.join(output, sample.hwp.file));
    if (crypto.createHash('sha256').update(bytes).digest('hex') !== sample.hwp.sha256) throw Error('Digest mismatch: ' + sample.id);
    const doc = new HwpDocument(bytes);
    try {
      const controls = JSON.parse(doc.getControls());
      const sections = Array.from({ length: doc.getSectionCount() }, (_, s) => ({
        page: JSON.parse(doc.getPageDef(s)), columns: JSON.parse(doc.getColumnDef(s)), paragraphs: doc.getParagraphCount(s),
      }));
      const sourcePages = doc.pageCount();
      const sourceText = Array.from({ length: sourcePages }, (_, p) =>
        (JSON.parse(doc.getPageTextLayout(p)).runs || []).map(run => run.text).join('\n'));
      const selectedPages = [...new Set([0, Math.min(1, sourcePages - 1), Math.floor(sourcePages / 2), sourcePages - 1])];
      for (const p of selectedPages) await fs.writeFile(path.join(output, 'evidence', `${sample.id}-source-${p + 1}.svg`), doc.renderPageSvg(p));
      const metadata = await metadataBridge.exports.readHwpLayoutMetadata(bytes);
      const editable = bridge.exports.editableHwp(doc, undefined, metadata);
      await fs.writeFile(path.join(output, 'converted', sample.id + '.editable.json'), JSON.stringify(editable));
      await fs.writeFile(path.join(output, 'converted', sample.id + '.source-text.json'), JSON.stringify(sourceText));
      await fs.writeFile(path.join(output, 'converted', sample.id + '.controls.json'), JSON.stringify(controls, null, 2));
      const paragraphExamples = [];
      for (let s = 0; s < sections.length; s++) for (let p = 0; p < Math.min(sections[s].paragraphs, 12); p++)
        paragraphExamples.push({ section: s, paragraph: p, properties: JSON.parse(doc.getParaPropertiesAt(s, p)) });
      results.push({ id: sample.id, title: sample.title, sourcePages, sections, selectedSourcePages: selectedPages.map(p => p + 1),
        documentInfo: JSON.parse(doc.getDocumentInfo()), controlCounts: controls.reduce((a, c) => (a[c.ctrlId] = (a[c.ctrlId] || 0) + 1, a), {}),
        topLevelControls: controls.filter(c => c.list === 0).length, paragraphExamples, warnings: editable.warnings,
        htmlBytes: Buffer.byteLength(editable.html) });
      console.log(sample.id + ': original ' + sourcePages + ' pages; controls ' + JSON.stringify(results.at(-1).controlCounts));
    } finally { doc.free(); }
  }
  await fs.writeFile(path.join(output, 'source-metrics.json'), JSON.stringify(results, null, 2));

  // A snapshot avoids File Provider changes reloading the document mid-audit.
  const snapshot = await fs.mkdtemp(path.join(os.homedir(), 'Library/Caches/word-hwp-audit-'));
  let server, browser;
  try {
    for (const entry of ['src', 'index.html', 'package.json', 'vite.config.ts', 'vendor', 'electron'])
      await fs.cp(path.join(root, entry), path.join(snapshot, entry), { recursive: true });
    await fs.mkdir(path.join(snapshot, 'node_modules'));
    for (const dependency of await fs.readdir(path.join(root, 'node_modules'))) if (!['.vite', '.cache'].includes(dependency))
      await fs.symlink(path.join(root, 'node_modules', dependency), path.join(snapshot, 'node_modules', dependency));
    await fs.writeFile(path.join(snapshot, 'src/hwp-conversion-audit.ts'), "export { generateHTML } from '@tiptap/core';\nexport { default as JSZip } from 'jszip';\n");
    const { createServer } = await import('vite');
    server = await createServer({ root: snapshot, server: { host: '127.0.0.1', port: 5193, strictPort: true, hmr: false, watch: { ignored: ['**'] }, fs: { allow: [snapshot, root] } } });
    await server.listen();
    browser = await chromium.launch({ channel: process.env.PLAYWRIGHT_CHANNEL || 'chrome', headless: true });
    const page = await browser.newPage({ viewport: { width: 1280, height: 1000 }, deviceScaleFactor: 1 });
    page.setDefaultTimeout(180000);
    const rendererErrors = [];
    page.on('pageerror', error => rendererErrors.push(error.message));
    await page.goto('http://127.0.0.1:5193');
    await page.locator('.tiptap').waitFor();
    await page.addStyleTag({ content: '#root{display:none!important} body{overflow:visible!important}' });
    if (process.env.HWP_AUDIT_EXPERIMENTS === '1') {
      const experiments = [];
      for (const id of ['03-health-table', '05-return-form']) {
        const doc = await readJson(path.join(output, 'converted', id + '.wordstudio'));
        const outcome = await page.evaluate(async ({ doc, id }) => {
          const { generateHTML } = await import('/src/hwp-conversion-audit.ts');
          const { editorExtensions } = await import('/src/extensions.ts');
          const { createPrintLayout } = await import('/src/print-layout.ts');
          const { exportDocx } = await import('/src/document-io.ts');
          const html = generateHTML(doc.content, editorExtensions), variants = [];
          for (const mode of ['baseline', 'serif-font', 'empty-paragraph-metrics', 'serif-and-empty']) {
            const container = document.createElement('div'); container.innerHTML = html;
            if (mode.includes('serif')) for (const e of container.querySelectorAll('[style*="font-family"]'))
              e.style.fontFamily = '"Nanum Myeongjo", serif';
            if (mode.includes('empty')) for (const p of container.querySelectorAll('p')) if (!p.textContent.trim() && !p.querySelector('img'))
              p.style.cssText += ';font-size:10pt;line-height:100%;min-height:0;';
            const layout = await createPrintLayout(container.innerHTML, doc.page, () => {});
            variants.push({ mode, studioPages: layout.pages });
            if (mode === 'serif-and-empty') {
              window.auditLayout?.dispose(); window.auditLayout = layout;
              layout.element.classList.remove('measuring'); layout.element.id = 'experiment-layout';
            } else layout.dispose();
          }
          const normalized = structuredClone(doc);
          let changed = 0;
          const walk = n => { if (typeof n.attrs?.lineSpacing === 'string' && n.attrs.lineSpacing.endsWith('%')) { n.attrs.lineSpacing = String(parseFloat(n.attrs.lineSpacing) / 100); changed++; } n.content?.forEach(walk); };
          walk(normalized.content);
          const blob = await exportDocx(normalized), bytes = new Uint8Array(await blob.arrayBuffer());
          let binary = ''; for (let i = 0; i < bytes.length; i += 8192) binary += String.fromCharCode(...bytes.subarray(i, i + 8192));
          return { id, variants, normalizedParagraphs: changed, normalizedDocx: btoa(binary) };
        }, { doc, id });
        await page.locator('#experiment-layout .print-page').first().screenshot({ path: path.join(output, 'evidence', id + '-experiment-serif-empty.png') });
        await fs.writeFile(path.join(output, 'converted', id + '-percent-normalized.docx'), Buffer.from(outcome.normalizedDocx, 'base64'));
        delete outcome.normalizedDocx; experiments.push(outcome);
      }
      for (const id of ['07-english-exam', '08-math-exam', '09-korean-exam']) {
        const base64 = (await fs.readFile(path.join(output, 'converted', id + '.docx'))).toString('base64');
        const sections = results.find(r => r.id === id).sections;
        const variants = await page.evaluate(async ({ base64, sections }) => {
          const { JSZip } = await import('/src/hwp-conversion-audit.ts');
          const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', output = [];
          for (const mode of ['paper', 'paper-columns']) {
            const zip = await JSZip.loadAsync(base64, { base64: true });
            const xml = new DOMParser().parseFromString(await zip.file('word/document.xml').async('string'), 'application/xml');
            if (xml.querySelector('parsererror')) throw Error('Invalid generated DOCX XML');
            [...xml.getElementsByTagNameNS(w, 'sectPr')].forEach((s, i) => {
              const ref = sections[Math.min(i, sections.length - 1)], size = s.getElementsByTagNameNS(w, 'pgSz')[0];
              size.setAttributeNS(w, 'w:w', String(Math.round(ref.page.width / 5)));
              size.setAttributeNS(w, 'w:h', String(Math.round(ref.page.height / 5)));
              if (mode === 'paper-columns') {
                let cols = s.getElementsByTagNameNS(w, 'cols')[0];
                if (!cols) { cols = xml.createElementNS(w, 'w:cols'); s.append(cols); }
                cols.setAttributeNS(w, 'w:num', String(ref.columns.columnCount));
                cols.setAttributeNS(w, 'w:space', String(Math.round(ref.columns.spacing / 5)));
              }
            });
            zip.file('word/document.xml', new XMLSerializer().serializeToString(xml));
            output.push({ mode, base64: await zip.generateAsync({ type: 'base64', compression: 'DEFLATE' }) });
          }
          return output;
        }, { base64, sections });
        for (const variant of variants) await fs.writeFile(path.join(output, 'converted', id + '-' + variant.mode + '.docx'), Buffer.from(variant.base64, 'base64'));
      }
      await fs.writeFile(path.join(output, 'experiments.json'), JSON.stringify(experiments, null, 2));
      console.log(JSON.stringify(experiments));
      return;
    }
    for (const result of results) {
      const source = JSON.parse(await fs.readFile(path.join(output, 'converted', result.id + '.editable.json'), 'utf8'));
      const converted = await page.evaluate(async ({ source, name }) => {
        const { importEditableHwp } = await import('/src/hwp/editing.ts');
        const { prepareEditableHwp } = await import('/src/hwp/svg-images.ts');
        const { exportDocx } = await import('/src/document-io.ts');
        const { generateHTML } = await import('/src/hwp-conversion-audit.ts');
        const { editorExtensions } = await import('/src/extensions.ts');
        const { createPrintLayout } = await import('/src/print-layout.ts');
        source = await prepareEditableHwp(source);
        const doc = importEditableHwp(source, name + '.hwp');
        const html = generateHTML(doc.content, editorExtensions);
        const layout = await createPrintLayout(html, doc.page, () => {});
        const stats = {};
        const walk = node => { stats[node.type] = (stats[node.type] || 0) + 1; node.content?.forEach(walk); };
        walk(doc.content);
        const originalHtml = document.createElement('div'); originalHtml.innerHTML = source.html;
        const selectors = ['p', 'table', 'tr', 'td', 'img', '[data-page-break]', '[data-section-page]'];
        const htmlCounts = Object.fromEntries(selectors.map(s => [s, originalHtml.querySelectorAll(s).length]));
        const pages = [...layout.element.querySelectorAll('.print-page')].map(p => ({
          text: p.querySelector('.print-content')?.textContent || '', tables: p.querySelectorAll('table').length,
          images: p.querySelectorAll('img').length,
        }));
        window.auditLayout?.dispose(); window.auditLayout = layout;
        layout.element.classList.remove('measuring');
        layout.element.style.cssText = 'position:relative;visibility:visible;z-index:99999;background:#bbb;';
        layout.element.id = 'audit-layout';
        const blob = await exportDocx(doc);
        const bytes = new Uint8Array(await blob.arrayBuffer());
        let binary = ''; for (let i = 0; i < bytes.length; i += 8192) binary += String.fromCharCode(...bytes.subarray(i, i + 8192));
        return { doc, html, pages, studioPages: layout.pages, layoutWarnings: layout.warnings, stats, htmlCounts, docxBase64: btoa(binary) };
      }, { source, name: result.id });
      await fs.writeFile(path.join(output, 'converted', result.id + '.docx'), Buffer.from(converted.docxBase64, 'base64'));
      await fs.writeFile(path.join(output, 'converted', result.id + '.wordstudio'), JSON.stringify(converted.doc));
      await fs.writeFile(path.join(output, 'converted', result.id + '.editor.html'), converted.html);
      await fs.writeFile(path.join(output, 'converted', result.id + '.studio-pages.json'), JSON.stringify(converted.pages));
      for (const p of [...new Set([0, converted.studioPages - 1])]) {
        await page.locator('#audit-layout .print-page').nth(p).screenshot({ path: path.join(output, 'evidence', result.id + '-studio-' + (p + 1) + '.png') });
      }
      Object.assign(result, { studioPages: converted.studioPages, editorNodes: converted.stats, htmlCounts: converted.htmlCounts, layoutWarnings: converted.layoutWarnings });
      console.log(result.id + ': Word Studio ' + result.studioPages + ' pages; ' + JSON.stringify(converted.stats));
      await fs.writeFile(path.join(output, 'results.json'), JSON.stringify({ corpusDate: manifest.date, runAtUTC: new Date().toISOString(),
        sourceCommit: execFileSync('git', ['rev-parse', 'HEAD'], { cwd: root, encoding: 'utf8' }).trim(), sourceTreeSHA256, workingTreeDirty, appVersion: require('../package.json').version,
        engine: require('../vendor/rhwp-core/package.json').version, browser: await browser.version(), rendererErrors, results }, null, 2));
    }
    if (rendererErrors.length) throw Error('Renderer errors: ' + rendererErrors.join('; '));
  } finally {
    await browser?.close(); await server?.close();
    await fs.rm(snapshot, { recursive: true, force: true });
  }
}
main().catch(error => { console.error(error); process.exitCode = 1; });

const { chromium } = require('@playwright/test');
const { spawn } = require('node:child_process');
const fs = require('node:fs/promises');
const { createHash } = require('node:crypto');
const path = require('node:path');
const JSZip = require('jszip');

const ROOT = path.resolve(__dirname, '..');
const INPUT_DIR = path.resolve(process.argv[2] || path.join(ROOT, 'analysis/official-sample-validation'));
const OUT_DIR = path.resolve(process.argv[3] || path.join(INPUT_DIR, 'after'));
const SAMPLES_DIR = path.join(OUT_DIR, 'samples');
const SCREENSHOTS_DIR = path.join(OUT_DIR, 'screenshots');
const PDF_DIR = path.join(OUT_DIR, 'pdfs');

function countMatches(str, regex) {
  if (!str) return 0;
  const m = str.match(regex);
  return m ? m.length : 0;
}

function decodeXmlEntities(s) {
  return s
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function extractXmlText(xml) {
  if (!xml) return '';
  const texts = [];
  const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
  let m;
  while ((m = re.exec(xml)) !== null) {
    texts.push(decodeXmlEntities(m[1]));
  }
  return texts.join('');
}

async function analyzeOoxml(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const docXml = await zip.file('word/document.xml')?.async('string') || '';
  const stylesXml = await zip.file('word/styles.xml')?.async('string') || '';
  const numberingXml = await zip.file('word/numbering.xml')?.async('string') || '';
  const commentsXml = await zip.file('word/comments.xml')?.async('string') || '';
  const footnotesXml = await zip.file('word/footnotes.xml')?.async('string') || '';
  const endnotesXml = await zip.file('word/endnotes.xml')?.async('string') || '';
  const settingsXml = await zip.file('word/settings.xml')?.async('string') || '';

  const headerFiles = Object.keys(zip.files).filter(f => /^word\/header\d*\.xml$/.test(f));
  const footerFiles = Object.keys(zip.files).filter(f => /^word\/footer\d*\.xml$/.test(f));
  let headersXml = '';
  for (const f of headerFiles) headersXml += (await zip.file(f).async('string')) + '\n';
  let footersXml = '';
  for (const f of footerFiles) footersXml += (await zip.file(f).async('string')) + '\n';

  // Extract multi-column info
  const colsMatches = [...docXml.matchAll(/<w:cols\b([^>]*)\/?>/g)].map(m => {
    const num = m[1].match(/w:num="(\d+)"/)?.[1] || '1';
    const equalWidth = m[1].match(/w:equalWidth="([^"]+)"/)?.[1] || '1';
    return { num: Number(num), equalWidth };
  });

  // Extract section break types
  const sectionTypes = [...docXml.matchAll(/<w:type\s+w:val="([^"]+)"/g)].map(m => m[1]);

  // Extract wrap types on anchor images
  const wrapTypes = [...docXml.matchAll(/<wp:(wrapSquare|wrapTight|wrapThrough|wrapTopAndBottom|wrapNone)\b/g)].map(m => m[1]);

  // Extract complex fields
  const instrTexts = [...docXml.matchAll(/<w:instrText(?:\s[^>]*)?>([\s\S]*?)<\/w:instrText>/g)].map(m => decodeXmlEntities(m[1]).trim()).filter(Boolean);
  const fldSimples = [...docXml.matchAll(/<w:fldSimple\s+[^>]*w:instr="([^"]+)"/g)].map(m => decodeXmlEntities(m[1]).trim());

  // Extract w:sym elements (Symbol font characters)
  const symElements = [...docXml.matchAll(/<w:sym\b([^>]*)\/>/g)].map(m => {
    const font = m[1].match(/w:font="([^"]+)"/)?.[1] || '';
    const char = m[1].match(/w:char="([^"]+)"/)?.[1] || '';
    return { font, char };
  });

  // Separator IDs differ between producers; inspect the note type.
  const realNotes = (xml, kind) => [...xml.matchAll(new RegExp(`<w:${kind}\\b([^>]*)>([\\s\\S]*?)<\\/w:${kind}>`, 'g'))]
    .filter(m => !/w:type="(?:separator|continuationSeparator|continuationNotice)"/.test(m[1]) && /w:id="[0-9]+"/.test(m[1]));
  const footnotes = realNotes(footnotesXml, 'footnote'), endnotes = realNotes(endnotesXml, 'endnote');
  const footnoteCount = footnotes.length, endnoteCount = endnotes.length;
  const footnoteText = footnotes.map(m=>extractXmlText(m[2])).join('');
  const endnoteText = endnotes.map(m=>extractXmlText(m[2])).join('');

  // Textbox contents
  const textboxCount = countMatches(docXml, /<w:txbxContent\b/g);

  // Header/Footer tables or images
  const headerFooterTables = countMatches(headersXml + footersXml, /<w:tbl\b/g);
  const headerFooterImages = countMatches(headersXml + footersXml, /<a:blip\b|<v:imagedata\b/g);

  // Numbering formats
  const lvlTexts = [...numberingXml.matchAll(/<w:lvlText\s+w:val="([^"]+)"/g)].map(m => m[1]);
  const numFmts = [...numberingXml.matchAll(/<w:numFmt\s+w:val="([^"]+)"/g)].map(m => m[1]);

  // Internal bookmark links
  const internalLinks = [...docXml.matchAll(/<w:hyperlink\b[^>]*w:anchor="([^"]+)"/g)].map(m => m[1]);
  const bookmarks = [...docXml.matchAll(/<w:bookmarkStart\b[^>]*w:name="([^"]+)"/g)].map(m => m[1]).filter(n => n !== '_GoBack');

  // Character styling properties
  const charProps = {
    smallCaps: countMatches(docXml, /<w:smallCaps\b/g),
    allCaps: countMatches(docXml, /<w:caps\b/g),
    doubleStrike: countMatches(docXml, /<w:dstrike\b/g),
    emboss: countMatches(docXml, /<w:emboss\b/g),
    imprint: countMatches(docXml, /<w:imprint\b/g),
    outline: countMatches(docXml, /<w:outline\b/g),
    shadow: countMatches(docXml, /<w:shadow\b/g),
    underlineStyles: [...new Set([...docXml.matchAll(/<w:u\s+w:val="([^"]+)"/g)].map(m => m[1]))],
  };

  let tableDepth = 0, nestedTables = 0;
  for (const match of docXml.matchAll(/<(\/?)w:tbl\b[^>]*>/g)) {
    if (match[1]) tableDepth--; else { if (tableDepth) nestedTables++; tableDepth++; }
  }
  return {
    bodyTextLength: extractXmlText(docXml).length,
    paragraphCount: countMatches(docXml, /<w:p\b/g),
    tableCount: countMatches(docXml, /<w:tbl\b/g),
    nestedTableCount: nestedTables,
    gridSpanCount: countMatches(docXml, /<w:gridSpan\b/g),
    vMergeCount: countMatches(docXml, /<w:vMerge\b/g),
    hMergeCount: countMatches(docXml, /<w:hMerge\b/g),
    gridBeforeAfterCount: countMatches(docXml, /<w:grid(?:Before|After)\b/g),
    inlineImageCount: countMatches(docXml, /<wp:inline\b/g),
    anchorImageCount: countMatches(docXml, /<wp:anchor\b/g),
    vmlImageCount: countMatches(docXml, /<v:imagedata\b/g),
    wrapTypes: [...new Set(wrapTypes)],
    sectionCount: countMatches(docXml, /<w:sectPr\b/g),
    sectionTypes: [...new Set(sectionTypes)],
    multiColumnSections: colsMatches.filter(c => c.num > 1),
    headerFileCount: headerFiles.length,
    footerFileCount: footerFiles.length,
    headerFooterTables,
    headerFooterImages,
    evenAndOddHeaders: /<w:evenAndOddHeaders\b/.test(settingsXml),
    titlePgCount: countMatches(docXml, /<w:titlePg\b/g),
    tabStopDefs: countMatches(docXml + stylesXml, /<w:tabs\b/g),
    tabChars: countMatches(docXml, /<w:tab\b/g),
    commentCount: countMatches(commentsXml, /<w:comment\b/g),
    footnoteCount,
    footnoteTextLength: footnoteText.length,
    endnoteCount,
    endnoteTextLength: endnoteText.length,
    textboxCount,
    instrTexts: instrTexts.slice(0, 20),
    fldSimples: fldSimples.slice(0, 20),
    ommlEquations: countMatches(docXml, /<m:oMath\b/g),
    symElements,
    internalLinksCount: internalLinks.length,
    bookmarksCount: bookmarks.length,
    lvlTexts: [...new Set(lvlTexts)],
    numFmts: [...new Set(numFmts)],
    charProps,
  };
}

async function main() {
  await fs.mkdir(SCREENSHOTS_DIR, { recursive: true });
  await fs.mkdir(PDF_DIR, { recursive: true });

  const manifest = JSON.parse(await fs.readFile(path.join(INPUT_DIR, 'manifest.json'), 'utf8'));

  // Start Vite dev server
  const vite = spawn('npx', ['vite', '--host', '127.0.0.1', '--port', '5199', '--strictPort'], {
    cwd: ROOT,
    stdio: ['ignore', 'pipe', 'pipe'],
  });

  await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('Vite timeout')), 15000);
    vite.stdout.on('data', data => {
      if (data.toString().includes('5199')) {
        clearTimeout(timer);
        resolve();
      }
    });
    vite.on('exit', code => reject(new Error(`Vite exited ${code}`)));
  });

  const browser = await chromium.launch({ channel: "chrome" });
  const results = [];

  try {
    for (const item of manifest) {
      const filePath = path.join(INPUT_DIR, item.file);
      const buffer = await fs.readFile(filePath);
      const ooxml = await analyzeOoxml(buffer);

      const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
      const page = await context.newPage();
      const pageErrors = [];
      const consoleErrors = [];
      page.on('pageerror', e => pageErrors.push(e.message));
      page.on('console', msg => {
        if (msg.type() === 'error') consoleErrors.push(msg.text());
      });

      await page.goto('http://127.0.0.1:5199/');
      await page.locator('.tiptap').waitFor();

      // Programmatic import + roundtrip inspection
      const programmatic = await page.evaluate(async bytes => {
        try {
          const io = await import('/src/document-io.ts');
          const { parseDocumentHtml } = await import('/src/extensions.ts');

          const first = await io.importFile(new File([new Uint8Array(bytes)], 'sample.docx'));
          const doc = { ...first.document, content: parseDocumentHtml(first.html) };
          // Compare formatted text and destinations, rather than counts of runs
          // that merge/split when OOXML is serialized.
          const semantic = (node, out = { text:'', bold:'', italic:'', underline:'', strike:'', subscript:'', superscript:'', links:[] }) => {
            if (node.type === 'text') { out.text += node.text; for (const mark of node.marks || []) {
              if (typeof out[mark.type] === 'string') out[mark.type] += node.text;
              if (mark.type === 'link') { const last = out.links.at(-1); if (last?.href === mark.attrs.href) last.text += node.text; else out.links.push({ href:mark.attrs.href, text:node.text }); }
            } }
            for (const child of node.content || []) semantic(child,out); return out;
          };
          let validationError = null;
          try { io.validateDocument(doc); } catch (e) { validationError = String(e); }

          let roundtripOk = false;
          let roundtripError = null;
          let roundtripDoc = null;
          try {
            const saved = await io.exportDocx(doc);
            const second = await io.importFile(new File([saved], 'roundtrip.docx'));
            roundtripDoc = { ...second.document, content: parseDocumentHtml(second.html) };
            io.validateDocument(roundtripDoc);
            roundtripOk = true;
          } catch (e) {
            roundtripError = String(e);
          }

          return {
            ok: true,
            warnings: first.warnings,
            validationError,
            roundtripOk,
            roundtripError,
            semanticEqual: !!roundtripDoc && JSON.stringify(semantic(doc.content)) === JSON.stringify(semantic(roundtripDoc.content)),
            pageSettings: { ...doc.page, ...Object.fromEntries(Object.entries(doc.page).filter(([k])=>k.endsWith('Html')).map(([k,v])=>[k,`[${v.length} chars]`])) },
            sectionCount: (doc.sections?.length || 0) + 1,
            sections: doc.sections || [],
            commentsCount: doc.comments?.length || 0,
            comments: doc.comments || [],
            htmlLength: first.html.length,
          };
        } catch (err) {
          return { ok: false, error: String(err) };
        }
      }, [...buffer]);

      // UI file upload
      await page.locator('input[type=file]').first().setInputFiles({
        name: `${item.id}.docx`,
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        buffer,
      });
      await page.waitForTimeout(600);
      await page.evaluate(async()=>{await document.fonts.ready;await Promise.all([...document.querySelectorAll('.document-stack img[src]')].map(img=>img.decode().catch(()=>{})));});

      // Measure Editor metrics
      const editorMetrics = await page.evaluate(() => {
        const tiptap = document.querySelector('.document-stack .tiptap.ProseMirror');
        if (!tiptap) return null;
        const tiptapRect = tiptap.getBoundingClientRect();
        const style = getComputedStyle(tiptap);
        const contentWidth = tiptapRect.width - parseFloat(style.paddingLeft) - parseFloat(style.paddingRight);

        const tables = [...tiptap.querySelectorAll('table')].map(t => {
          const r = t.getBoundingClientRect();
          const parent=t.parentElement.closest('td,th,.section-editable')||tiptap;const ps=getComputedStyle(parent);const available=parent.clientWidth-parseFloat(ps.paddingLeft)-parseFloat(ps.paddingRight);
          const rows = [...t.querySelectorAll(':scope > tbody > tr:not(.pagination-gap), :scope > tr:not(.pagination-gap)')].length;
          return { width: Math.round(r.width), height: Math.round(r.height), rows, available, overflows: r.width > available + 4 };
        });

        const images = [...tiptap.querySelectorAll('img[src]')].map(img => {
          const r = img.getBoundingClientRect();
          return {
            inline: img.hasAttribute('data-inline-image'),
            width: Math.round(r.width),
            height: Math.round(r.height),
            naturalWidth: img.naturalWidth,
            naturalHeight: img.naturalHeight,
          };
        });

        const links = [...tiptap.querySelectorAll('a')].map(a => ({
          href: a.getAttribute('href'),
          text: a.textContent,
        }));

        const tabs = tiptap.querySelectorAll('[data-docx-tab]').length;
        const pageBreaks = tiptap.querySelectorAll('.page-break').length;
        const orderedLists = tiptap.querySelectorAll('ol').length;
        const bulletLists = tiptap.querySelectorAll('ul').length;
        const headings = [...tiptap.querySelectorAll('h1,h2,h3,h4,h5,h6')].map(h => ({ tag: h.tagName, text: h.textContent?.trim() }));

        return {
          textLength: (tiptap.textContent || '').length,
          textSample: (tiptap.textContent || '').slice(0, 200),
          contentWidth: Math.round(contentWidth),
          tablesCount: tables.length,
          overflowingTablesCount: tables.filter(t => t.overflows).length,
          tables,
          imagesCount: images.length,
          images,
          linksCount: links.length,
          internalAnchorLinksCount: links.filter(l => l.href && l.href.startsWith('#')).length,
          tabsCount: tabs,
          pageBreaksCount: pageBreaks,
          orderedListsCount: orderedLists,
          bulletListsCount: bulletLists,
          headingsCount: headings.length,
        };
      });

      await page.screenshot({ path: path.join(SCREENSHOTS_DIR, `${item.id}-editor.png`) });

      // Switch to Print Preview
      await page.getByRole('button', { name: '보기', exact: true }).click();
      await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
      await page.locator('.print-page').first().waitFor({ timeout: 10000 });
      await page.waitForTimeout(400);

      const previewMetrics = await page.evaluate(() => {
        const pages = [...document.querySelectorAll('.print-page')];
        return {
          pageCount: pages.length,
          scaledItemsWarning: document.body.textContent?.includes('축소') || false,
          pages: pages.map((p, idx) => {
            const rect = p.getBoundingClientRect();
            const content = p.querySelector('.print-content');
            const header = p.querySelector('.print-header');
            const footer = p.querySelector('.print-footer');
            return {
              pageNumber: idx + 1,
              width: Math.round(rect.width),
              height: Math.round(rect.height),
              orientation: rect.width > rect.height ? 'landscape' : 'portrait',
              headerText: header?.textContent?.trim() || '',
              footerText: footer?.textContent?.trim() || '',
              textLength: (content?.textContent || '').length,
              footnotes: p.querySelectorAll('.print-footnotes [data-document-note]').length,
              columnOverflow: [...p.querySelectorAll('[data-column-count]')].some(el=>el.scrollWidth>el.clientWidth+2),
              overlapsHeader: !!(header?.textContent.trim() || header?.querySelector('img,table')) && header.getBoundingClientRect().bottom > content.getBoundingClientRect().top+1,
              empty: !content?.textContent.trim() && !content?.querySelector('img,table'),
              intentionalBlank: p.dataset.keepBlank === 'true',
              tablesCount: content ? content.querySelectorAll('table').length : 0,
              imagesCount: content ? content.querySelectorAll('img').length : 0,
            };
          }),
        };
      });

      await page.screenshot({ path: path.join(SCREENSHOTS_DIR, `${item.id}-preview.png`) });

      // Export PDF via Playwright print emulation (same print tree as Electron printToPDF)
      await page.evaluate(() => document.body.classList.add('studio-printing'));
      const pdfBuffer = await page.pdf({ printBackground: true, preferCSSPageSize: true });
      await page.evaluate(() => document.body.classList.remove('studio-printing'));
      await fs.writeFile(path.join(PDF_DIR, `${item.id}.pdf`), pdfBuffer);
      const pdfPageCount = (pdfBuffer.toString('latin1').match(/\/Type\s*\/Page\b/g) || []).length;

      await context.close();

      const record = {
        id: item.id,
        sha256: createHash('sha256').update(buffer).digest('hex'),
        sourceHashMatches: !item.sha256 || createHash('sha256').update(buffer).digest('hex') === item.sha256,
        source: item.source,
        desc: item.desc,
        ooxml,
        programmatic,
        editorMetrics,
        previewMetrics,
        pdfPageCount,
        pageErrors,
        consoleErrors,
      };
      record.passed = record.sourceHashMatches && programmatic.ok && !programmatic.validationError && programmatic.roundtripOk && programmatic.semanticEqual
        && !pageErrors.length && editorMetrics.overflowingTablesCount === 0 && previewMetrics.pageCount === pdfPageCount
        && previewMetrics.pages.every(p=>!p.overlapsHeader && !p.columnOverflow && (!p.empty || p.intentionalBlank));
      results.push(record);
      console.log(`Validated ${item.id}: open=${programmatic.ok}, previewPages=${previewMetrics.pageCount}, pdfPages=${pdfPageCount}, errors=${pageErrors.length}`);
    }
  } finally {
    await browser.close();
    vite.kill();
  }

  await fs.writeFile(path.join(OUT_DIR, 'validation-results.json'), JSON.stringify(results, null, 2));
  console.log('Saved validation-results.json');
  if (results.some(r=>!r.passed)) process.exitCode = 1;
}

module.exports = { analyzeOoxml };
if (require.main === module) main().catch(err => {
  console.error(err);
  process.exit(1);
});

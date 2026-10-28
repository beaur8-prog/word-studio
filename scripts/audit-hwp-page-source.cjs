'use strict';
// Read-only source geometry for page correspondence. It shares the HWP parser;
// independent PDF evidence is collected separately by the companion audit.
const fs = require('node:fs/promises'), path = require('node:path');
const { pathToFileURL } = require('node:url');
const { createHash } = require('node:crypto');
const normalize = text => text.replace(/[\s\u0000-\u001f\ufffc]/gu, '');
(async () => {
  const samples = process.argv[2] || 'analysis/hwp-row-flow-2026-10-26/samples';
  const output = process.argv[3] || 'analysis/hwp-12h-2026-10-28/source';
  const ids = process.argv.slice(4);
  const files = ids.length ? ids.map(id => id + '.hwp') : (await fs.readdir(samples)).filter(file => file.endsWith('.hwp')).sort();
  await fs.mkdir(output, { recursive: true });
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')));
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  for (const file of files) {
    const bytes = await fs.readFile(path.join(samples, file)), doc = new HwpDocument(bytes);
    const groups = new Map(), tables = new Map(), pages = [];
    try {
      for (let page = 0; page < doc.pageCount(); page++) {
        const runs = JSON.parse(doc.getPageTextLayout(page)).runs || [];
        pages.push({ page: page + 1, text: runs.filter(r => !r.headerFooter && r.paraIdx !== 0xffffffff).map(r => r.text).join('') });
        for (const run of runs) {
          if (run.headerFooter || run.paraIdx === 0xffffffff || !run.text?.trim()) continue;
          const paragraph = run.parentParaIdx ?? run.paraIdx;
          const cellPath = run.cellPath || [];
          const key = `${run.secIdx}:${paragraph}` + cellPath.map(c => `/t${c.controlIndex}/c${c.cellIndex}/p${c.cellParaIndex}`).join('');
          let group = groups.get(key);
          if (!group) {
            let text, properties;
            try {
              if (cellPath.length) {
                text = doc.getTextInCellByPath(run.secIdx, paragraph, JSON.stringify(cellPath), 0, 0xffffffff);
                if (cellPath.length === 1) { const c = cellPath[0]; properties = JSON.parse(doc.getCellParaPropertiesAt(run.secIdx, paragraph, c.controlIndex, c.cellIndex, c.cellParaIndex)); }
              } else { text = doc.getTextRange(run.secIdx, paragraph, 0, doc.getParagraphLength(run.secIdx, paragraph)); properties = JSON.parse(doc.getParaPropertiesAt(run.secIdx, paragraph)); }
            } catch { /* Keep positioned runs even if an optional query is unsupported. */ }
            group = { key, section: run.secIdx, paragraph, cellPath, text: text || '', normalized: normalize(text || ''), properties, lines: [] };
            groups.set(key, group);
          }
          group.lines.push({ page: page + 1, x: run.x, y: run.y, width: run.w, height: run.h, charStart: run.charStart, text: run.text, fontFamily: run.fontFamily, fontSize: run.fontSize, ratio: run.ratio, letterSpacing: run.letterSpacing });
        }
        for (const c of JSON.parse(doc.getPageControlLayout(page)).controls || []) {
          if (c.type !== 'table' || c.headerFooter || c.stableIndex?.length !== 3) continue;
          const key = c.stableIndex.join(':'); let table = tables.get(key);
          if (!table) {
            let properties, dimensions;
            try { properties = JSON.parse(doc.getTableProperties(...c.stableIndex)); dimensions = JSON.parse(doc.getTableDimensions(...c.stableIndex)); } catch {}
            table = { key, properties, dimensions, fragments: [] }; tables.set(key, table);
          }
          table.fragments.push({ page: page + 1, x: c.x, y: c.y, width: c.w, height: c.h, cells: c.cells });
        }
      }
      const result = { id: file.replace(/\.hwp$/, ''), inputSHA256: createHash('sha256').update(bytes).digest('hex'), sourcePages: doc.pageCount(), reference: 'native source reader, shared parser', pages, paragraphs: [...groups.values()], tables: [...tables.values()] };
      await fs.writeFile(path.join(output, result.id + '.json'), JSON.stringify(result, null, 2));
      console.log(JSON.stringify({ id: result.id, pages: result.sourcePages, paragraphs: groups.size, tables: tables.size }));
    } finally { doc.free(); }
  }
})().catch(error => { console.error(error); process.exitCode = 1; });

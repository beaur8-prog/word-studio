'use strict';

const { readFile } = require('node:fs/promises');
const { pathToFileURL } = require('node:url');
const { dirname, join } = require('node:path');

module.exports = async function captureTerminalWitness(page, sourcePath) {
  const engineRoot = dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(join(engineRoot, 'rhwp.js')).href);
  await init({ module_or_path: await readFile(join(engineRoot, 'rhwp_bg.wasm')) });
  const source = new HwpDocument(await readFile(sourcePath));
  const wanted = [119, 120, 121, 122, 123, 124, 163, 164, 165, 166, 167, 168];
  let keys;
  try {
    keys = wanted.map(index => [index, source.getTextRange(0, index, 0, 0xffffffff).replace(/\s/gu, '')]);
  } finally { source.free(); }

  const captured = await page.evaluate(keys => {
    const records = [];
    const paragraphs = [...document.querySelectorAll('.print-layout.preview .print-content p')];
    for (const [index, key] of keys) {
      const node = paragraphs.find(value => value.textContent.replace(/\s/gu, '') === key);
      if (!node) { records.push({ index, found: false }); continue; }
      node.dataset.qaHwpOwner = String(index);
      const paper = node.closest('.print-page');
      const frame = node.closest('.print-content');
      const column = node.closest('[data-column-count]');
      const paperBox = paper.getBoundingClientRect();
      const pBox = node.getBoundingClientRect();
      const frameBox = frame.getBoundingClientRect();
      const css = getComputedStyle(node);
      const relative = rect => [rect.left-paperBox.left, rect.top-paperBox.top, rect.width, rect.height].map(value=>Math.round(value*100)/100);
      const tail = node.querySelector('[data-hwp-column-tail]');
      const region = document.createRange();
      const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT);
      const glyphs = [];
      while (walker.nextNode()) {
        const leaf = walker.currentNode;
        if (!leaf.data.trim()) continue;
        const end = leaf.data.trimEnd().length;
        const start = leaf.data.length - leaf.data.trimStart().length;
        region.setStart(leaf, start);
        region.setEnd(leaf, end);
        const mark = leaf.parentElement.closest('span[style*="--word-character-scale"]');
        glyphs.push({ count: leaf.data.replace(/\s/gu, '').length,
          rects: [...region.getClientRects()].filter(rect => rect.width > 0).map(relative),
          scale: mark?.style.getPropertyValue('--word-character-scale') || null });
      }
      records.push({ index, found: true, page: Number(paper.dataset.page), p: relative(pBox), frame: relative(frameBox),
        fragmentRects: [...node.getClientRects()].map(relative), glyphs,
        tail: tail ? relative(tail.getBoundingClientRect()) : null,
        lineHeight: css.lineHeight, indent: css.textIndent,
        columnChildren: column?.childElementCount ?? null,
        flags: [...node.attributes].filter(value => value.name.startsWith('data-hwp-column')).map(value => value.name) });
    }
    return records;
  }, keys);
  const cdp = await page.context().newCDPSession(page);
  try {
    await cdp.send('DOM.enable');
    await cdp.send('CSS.enable');
    const { root } = await cdp.send('DOM.getDocument');
    for (const record of captured.filter(value => value.found)) {
      const { nodeId } = await cdp.send('DOM.querySelector', { nodeId: root.nodeId, selector: '[data-qa-hwp-owner="' + record.index + '"]' });
      const { fonts } = await cdp.send('CSS.getPlatformFontsForNode', { nodeId });
      record.physicalFonts = fonts.map(font=>[font.familyName,font.postScriptName,font.isCustomFont,font.glyphCount]);
    }
  } finally { await cdp.detach(); }
  console.log('WINDOWS_HWP_TERMINAL_DETAIL ' + JSON.stringify(captured));
};

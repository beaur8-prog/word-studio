'use strict';

module.exports = async function wordWrapWidthWitness(page) {
  const values = await page.evaluate(() => {
    const range = document.createRange();
    const segmenter = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
    const sheet = document.querySelector('.print-layout.preview .print-page[data-page="4"]');
    if (!sheet) throw Error('The pinned Word fixture was missing its fourth page');
    const records = [];
    for (const p of sheet.querySelectorAll('.print-content > p[data-word-line-rule="auto"]')) {
      const tab = p.querySelector('[data-docx-tab]');
      if (!tab || p.querySelectorAll('[data-docx-tab]').length !== 1 || !getComputedStyle(p).fontFamily.includes('Century Gothic')) continue;
      const text = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      const nodes = [];
      while (text.nextNode()) {
        const node = text.currentNode;
        if (node.parentElement.closest('[data-docx-tab]')) continue;
        for (const entry of segmenter.segment(node.data)) if (entry.segment.trim()) nodes.push({ node, ...entry });
      }
      if (nodes.length !== 212) continue;
      const style = getComputedStyle(p);
      const raw = p.getBoundingClientRect();
      p.setAttribute('data-qa-physical-wrap-witness', '');
      const physical = ({ zoom, first, last }) => {
        range.setStart(nodes[first].node, nodes[first].index);
        range.setEnd(nodes[last - 1].node, nodes[last - 1].index + nodes[last - 1].segment.length);
        const content = range.cloneContents();
        const host = document.createElement('div'), line = p.cloneNode(false), inline = document.createElement('span');
        host.className = 'tiptap print-content';
        host.style.cssText = 'position:fixed;visibility:hidden;left:-100000px;top:0;margin:0;padding:0;width:1400px';
        host.lang = p.closest('[lang]')?.getAttribute('lang') ?? 'en-US';
        line.removeAttribute('data-qa-physical-wrap-witness');
        Object.assign(line.style, {
          width: '1300px', margin: '0', padding: '0', textIndent: '0', textAlign: 'left',
          fontFamily: style.fontFamily, fontWeight: style.fontWeight, fontStyle: style.fontStyle,
          fontSize: style.fontSize, lineHeight: style.lineHeight, letterSpacing: style.letterSpacing,
          wordSpacing: style.wordSpacing, zoom: String(zoom),
        });
        inline.style.whiteSpace = 'pre';
        inline.style.display = 'inline-block';
        inline.append(content);
        line.append(inline);
        host.append(line);
        document.body.append(host);
        try { return inline.getBoundingClientRect().width / line.currentCSSZoom; }
        finally { host.remove(); }
      };
      const groups = [[65,134],[134,203]];
      const printed = groups.map(([first,last]) => ({
        first, last, natural: physical({ zoom:1,first,last }),
        referenceDensity: physical({ zoom:1/devicePixelRatio,first,last }),
      }));
      records.push({ width: raw.width, font: style.fontFamily, fontSize: style.fontSize, fontWeight: style.fontWeight,
        indent: style.textIndent, leftMargin: style.marginLeft, tab: tab.getBoundingClientRect().width, printed });
    }
    return records;
  });
  if (values.length !== 1) throw Error('The pinned Word line-width witness was missing or ambiguous');
  const cdp = await page.context().newCDPSession(page);
  try {
    await cdp.send('DOM.enable');
    await cdp.send('CSS.enable');
    const { root } = await cdp.send('DOM.getDocument');
    const { nodeIds } = await cdp.send('DOM.querySelectorAll', {
      nodeId: root.nodeId, selector: '[data-qa-physical-wrap-witness], [data-qa-physical-wrap-witness] span',
    });
    const faces = new Set();
    for (const nodeId of nodeIds.slice(0,20)) {
      const { fonts } = await cdp.send('CSS.getPlatformFontsForNode', { nodeId });
      for (const font of fonts) if (font.glyphCount) faces.add(font.familyName);
    }
    return { ...values[0], actualFaces: [...faces] };
  } finally {
    await cdp.detach();
  }
};

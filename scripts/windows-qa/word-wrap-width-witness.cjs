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
      const savedWordSpacing = p.style.wordSpacing;
      const windows = [], pitch = [];
      const base = parseFloat(style.wordSpacing) || 0;
      const grain = [0, 1, 2, 3, 4, 5, 6];
      try {
        for (const ticks of grain) {
          p.style.wordSpacing = (base - ticks / 64) + 'px';
          const lines = [];
          const bandTops = [];
          const edges = [];
          const area = p.getBoundingClientRect();
          for (const [index, node] of nodes.entries()) {
            range.setStart(node.node, node.index);
            range.setEnd(node.node, node.index + node.segment.length);
            const bounds = range.getBoundingClientRect();
            if (!bandTops.length || Math.abs(bandTops.at(-1) - bounds.top) > .5) {
              lines.push(index);
              bandTops.push(bounds.top);
            }
            edges.push(bounds.right - area.right);
          }
          windows.push({ ticks, starts: lines, height: area.height, rightOvershoot: Math.max(...edges) });
          pitch.push(bandTops.map((top, i) => i ? top - bandTops[i - 1] : 0));
        }
      } finally { p.style.wordSpacing = savedWordSpacing; }
      const compact = target => {
        const walker = document.createTreeWalker(target, NodeFilter.SHOW_TEXT);
        const breaks = [], bands = [], sheet = target.getBoundingClientRect();
        let right = -Infinity, index = 0;
        while (walker.nextNode()) {
          const node = walker.currentNode;
          if (node.parentElement.closest('[data-docx-tab]')) continue;
          for (const entry of segmenter.segment(node.data)) {
            if (!entry.segment.trim()) continue;
            range.setStart(node, entry.index);
            range.setEnd(node, entry.index + entry.segment.length);
            const rect = range.getBoundingClientRect();
            if (!bands.length || Math.abs(bands.at(-1)-rect.y) > .5) { bands.push(rect.y); breaks.push(index); }
            right = Math.max(right, rect.right-sheet.right);
            index++;
          }
        }
        return { breaks, glyphs:index, width:sheet.width / target.currentCSSZoom, right, height:sheet.height / target.currentCSSZoom };
      };
      const holder = document.createElement('div'), copy = p.cloneNode(true);
      holder.className = 'tiptap print-content';
      holder.lang = p.closest('[lang]')?.getAttribute('lang') ?? 'en-US';
      holder.style.cssText = 'position:fixed;left:-100000px;top:0;width:'+raw.width+'px;visibility:hidden;margin:0;padding:0;height:auto';
      Object.assign(copy.style, { width:raw.width+'px', margin:'0', fontFamily:style.fontFamily,
        fontSize:style.fontSize,fontStyle:style.fontStyle,fontWeight:style.fontWeight,
        lineHeight:style.lineHeight,letterSpacing:style.letterSpacing,zoom:String(1/devicePixelRatio) });
      holder.append(copy); document.body.append(holder);
      const clone = compact(copy);
      holder.remove();
      const visible = compact(p);
      const lineStyles = new Set(nodes.map(({node})=>{
        const run = getComputedStyle(node.parentElement);
        return [run.fontFamily,run.fontSize,run.fontStyle,run.fontWeight,run.letterSpacing,run.wordSpacing].join('|');
      }));
      const css = { align:style.textAlign, dir:style.direction, writingMode:style.writingMode, sizing:style.boxSizing,
        fit:p.style.width, zoom:p.currentCSSZoom, parent:p.parentElement.tagName, textLength:p.textContent.length,
        compatibleCharacters:/^[\x20-\x7e\t\u00a0\u2019]+$/u.test(p.textContent),
        fragment:p.hasAttribute('data-print-paragraph-fragment'), inlineFormats:p.querySelectorAll('br,img,svg,a,sup,sub,[data-note-reference],[data-document-field],[data-page-number],[data-symbol-font],[data-word-overlap]').length,
        lineStyles:[...lineStyles], paragraphStyles:[style.fontFamily,style.fontSize,style.fontStyle,style.fontWeight,style.letterSpacing,style.wordSpacing],
      };
      records.push({ width: raw.width, font: style.fontFamily, fontSize: style.fontSize, fontWeight: style.fontWeight,
        indent: style.textIndent, leftMargin: style.marginLeft, tab: tab.getBoundingClientRect().width, printed,
        spacingWitness: windows, pitch, clone, visible, css });
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

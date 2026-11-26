'use strict';

module.exports = async function wordDensityLineBoxes(page, id) {
  return page.evaluate(id => {
    const sheet = document.querySelector('.print-layout.preview .print-page[data-page="' + (id === '02-tables' ? 5 : 1) + '"]');
    if (!sheet) throw Error('Missing the pinned print page');
    const body = sheet.querySelector('.print-content'), frame = body.getBoundingClientRect();
    const fixed = value => Math.round(value * 1000) / 1000;
    const relative = box => [fixed(box.x-frame.x),fixed(box.y-frame.y),fixed(box.width),fixed(box.height)];
    const range = document.createRange();
    const description = paragraph => {
      const box = paragraph.getBoundingClientRect(), style = getComputedStyle(paragraph);
      const inline = [];
      const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
      while (walker.nextNode() && inline.length < 16) {
        const node = walker.currentNode, text = node.data, mark = node.parentElement;
        if (!text || !mark || mark.closest('[data-flow-widget]')) continue;
        range.selectNodeContents(node);
        const css = getComputedStyle(mark);
        inline.push({length:text.length, tab:!!mark.closest('[data-docx-tab]'), type:mark.tagName,
          font:css.fontFamily,size:css.fontSize,weight:css.fontWeight,line:css.lineHeight, rects:[...range.getClientRects()].map(relative)});
      }
      return {rect:relative(box), font:style.fontFamily, size:style.fontSize, height:style.lineHeight, inline};
    };
    let paragraphs = [...body.querySelectorAll('p,h1,h2,h3,h4,h5,h6')];
    if (id === '06-long-body') paragraphs = paragraphs.filter(p=>p.querySelector('[data-docx-tab]')).slice(0,8);
    else paragraphs = paragraphs.filter(p=>p.textContent.trim()).slice(0,10);
    return paragraphs.map(paragraph => {
      const baseline = description(paragraph);
      if (!paragraph.querySelector('[data-docx-tab]')) return {baseline};
      const frame = paragraph.getBoundingClientRect();
      // Measure a separate copy without changing the real printed line.
      const holder = document.createElement('div');
      holder.className = 'tiptap print-content';
      holder.style.cssText = 'position:fixed;left:-100000px;top:0;width:' + frame.width + 'px;visibility:hidden;margin:0;padding:0;height:auto';
      holder.lang = paragraph.closest('[lang]')?.getAttribute('lang') ?? 'en-US';
      const copy = paragraph.cloneNode(true), style = getComputedStyle(paragraph);
      Object.assign(copy.style, { width:frame.width+'px', margin:'0', fontSize:style.fontSize,
        fontFamily:style.fontFamily, fontWeight:style.fontWeight, fontStyle:style.fontStyle,
        lineHeight:style.lineHeight, color:style.color });
      holder.append(copy);
      document.body.append(holder);
      try {
        const heights = [fixed(copy.getBoundingClientRect().height)];
        for (const tab of copy.querySelectorAll('[data-docx-tab]')) tab.style.overflow = 'hidden';
        heights.push(fixed(copy.getBoundingClientRect().height));
        for (const tab of copy.querySelectorAll('[data-docx-tab]')) { tab.style.overflow = ''; tab.textContent = ''; }
        heights.push(fixed(copy.getBoundingClientRect().height));
        return {baseline, tabProbeHeights:heights};
      } finally { holder.remove(); }
    });
  }, id);
};

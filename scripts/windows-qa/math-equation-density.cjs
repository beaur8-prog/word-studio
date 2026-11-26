'use strict';

const { readFile, writeFile } = require('node:fs/promises');
const { join } = require('node:path');

async function record(page, destination) {
  const evidence = await page.evaluate(async () => {
    const normalized = text => text.replace(/\s/gu, '');
    const keys = new Map();
    function key(text) {
      if (!keys.has(text)) keys.set(text, crypto.subtle.digest('SHA-256', new TextEncoder().encode(text))
        .then(hash => Array.from(new Uint8Array(hash), n => n.toString(16).padStart(2, '0')).join('')));
      return keys.get(text);
    }
    const result = [];
    for (const pageNumber of [11,16]) {
      const sheet = document.querySelector('.print-layout.preview .print-page[data-page="' + pageNumber + '"]');
      const paper = sheet.getBoundingClientRect(), actualWidth = parseFloat(getComputedStyle(sheet).width);
      const scale = actualWidth / paper.width;
      const position = r => [r.x-paper.x,r.y-paper.y,r.width,r.height].map(n=>Math.round(n*scale*1000)/1000);
      for (const p of sheet.querySelectorAll('.print-content p')) {
        const imgs = [...p.querySelectorAll('img')].filter(image => image.alt.startsWith('수식:') && image.getBoundingClientRect().width > 0);
        if (!imgs.length) continue;
        const pBox = p.getBoundingClientRect();
        const style = getComputedStyle(p);
        const tabs = [...p.querySelectorAll('[data-docx-tab]')].map(tab => position(tab.getBoundingClientRect()));
        const fields = [...p.querySelectorAll('[data-print-paragraph-fragment],br')].length;
        const sameFlow = p.closest('[data-column-count]')?.getAttribute('data-column-count');
        const values = await Promise.all(imgs.map(async (img,index) => ({
          identity:await key(img.src), geometry:position(img.getBoundingClientRect()), index,
          display:getComputedStyle(img).display, verticalAlign:getComputedStyle(img).verticalAlign,
          inCell:!!img.closest('td'), floating:img.hasAttribute('data-image-floating'),
        })));
        result.push({page:pageNumber, key:await key(normalized(p.textContent) + ':' + values.map(v=>v.identity).join('')),
          length:normalized(p.textContent).length, p:position(pBox), font:style.fontFamily, pitch:style.lineHeight,
          whiteSpace:style.whiteSpace, indent:style.textIndent, leftMargin:style.marginLeft,
          tabs, fields, sameFlow, imageCount:imgs.length, values});
      }
    }
    return result;
  });
  await writeFile(destination, JSON.stringify(evidence));
}

async function compare(directory) {
  const read = async scale => JSON.parse(await readFile(join(directory, 'equations-' + scale + '.json'), 'utf8'));
  const before = await read('1');
  for (const density of ['1.25','1.5']) {
    const after = await read(density), changes = [];
    let checked = 0, missing = 0;
    for (const source of before) {
      const matched = after.filter(entry => entry.page === source.page && entry.key === source.key);
      if (matched.length !== 1) { missing++; continue; }
      const candidate = matched[0];
      for (const [index,image] of source.values.entries()) {
        const other = candidate.values[index];
        if (other?.identity !== image.identity) continue;
        checked++;
        const delta = Math.max(...image.geometry.map((n,i) => Math.abs(n-other.geometry[i])));
        if (delta < 4) continue;
        changes.push({ page:source.page, index, length:source.length, count:source.imageCount,
          delta:Math.round(delta*1000)/1000, p:[source.p,candidate.p], relativePicture:[image.geometry,other.geometry],
          whiteSpace:[source.whiteSpace,candidate.whiteSpace], tabs:[source.tabs,candidate.tabs],
          fields:[source.fields,candidate.fields], columns:source.sameFlow, pitch:[source.pitch,candidate.pitch],
          boxStyle:[image.display,image.verticalAlign,image.inCell,image.floating] });
      }
    }
    changes.sort((a,b) => b.delta - a.delta);
    console.log('WINDOWS_MATH_EQUATION_FLOW ' + JSON.stringify({ density, ownerCounts:[before.length,after.length],
      ownersWithoutExactMatch:missing, checked, changes:changes.slice(0,8) }));
  }
}

exports.record = record;

if (require.main === module) compare(process.argv[2]).catch(error => {
  console.error('WINDOWS_MATH_EQUATION_FLOW_FAILED ' + error.name);
  process.exitCode = 1;
});

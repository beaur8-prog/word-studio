'use strict';

async function wordTableWitness(page) {
  const sheets = await page.evaluate(async () => {
    const { createHash } = { createHash: (value) => {
      let hash = 2166136261;
      for (const char of value) hash = Math.imul(hash ^ char.codePointAt(0), 16777619);
      return (hash >>> 0).toString(16);
    } };
    const compact = value => value.normalize('NFKC').replace(/[\s\u00ad]+/g, '');
    const range = document.createRange();
    return [...document.querySelectorAll('.print-layout.preview .print-page')].map((paper, i) => {
      const frame = paper.getBoundingClientRect(), body = paper.querySelector('.print-content');
      const relative = element => {
        const b = element.getBoundingClientRect();
        return ['x','y','width','height'].map(key => +(b[key] - (key === 'x' || key === 'y' ? frame[key] : 0)).toFixed(2));
      };
      return { page: i + 1, body: relative(body), blocks: [...body.children].map(part => {
        const table = part.matches('table') ? part : part.querySelector(':scope > table');
        if (!table) return { type: part.tagName, rect: relative(part), chars: compact(part.textContent || '').length };
        return { type: 'table', rect: relative(table), sizing: getComputedStyle(table).tableLayout, rows: [...table.rows].map(row => {
          const cellText = [...row.cells].map(cell => compact(cell.textContent || ''));
          const lines = [...row.cells].map(cell => {
            range.selectNodeContents(cell);
            return new Set([...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0).map(rect => rect.top.toFixed(1))).size;
          });
          return { signature: createHash(cellText.join('|')), rect: relative(row), widths: [...row.cells].map(cell => relative(cell)[2]),
            lines, repeated: Boolean(row.closest('thead')), rowRule: row.dataset.wordRowRule || null };
        }) };
      }) };
    });
  });
  const session = await page.context().newCDPSession(page);
  try {
    await session.send('DOM.enable');
    await session.send('CSS.enable');
    const { root } = await session.send('DOM.getDocument');
    const { nodeIds } = await session.send('DOM.querySelectorAll', { nodeId: root.nodeId,
      selector: '.print-layout.preview .print-content table td:first-child > p:first-child' });
    const faces = new Map();
    for (const nodeId of nodeIds.slice(0, 20)) {
      const { fonts } = await session.send('CSS.getPlatformFontsForNode', { nodeId });
      for (const font of fonts) faces.set(font.postScriptName, { family: font.familyName, custom: font.isCustomFont });
    }
    return { sheets, fonts: [...faces] };
  } finally { await session.detach(); }
}

module.exports = wordTableWitness;

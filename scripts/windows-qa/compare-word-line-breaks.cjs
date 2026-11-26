'use strict';

const fs = require('node:fs');

function unique(index) {
  const result = new Map();
  for (const row of JSON.parse(fs.readFileSync(process.argv[index], 'utf8'))) {
    result.set(row.key, result.has(row.key) ? null : row);
  }
  return result;
}

const reference = unique(2);
const candidate = unique(3);
const pairs = [...reference.values()].filter(row => row && candidate.get(row.key))
  .map(row => [row, candidate.get(row.key)]);
const changed = [];
let sameSheet = 0, sameBreaks = 0, sameGlyphs = 0, onlyPosition = 0;
const compact = ({ sheet, starts, glyphs, lineY, y, width, height }) => ({
  sheet, starts, glyphs, lineY: lineY.map(v => Math.round(v * 100) / 100),
  y: Math.round(y * 100) / 100, width: Math.round(width * 100) / 100, height: Math.round(height * 100) / 100,
});
for (const [index, [a, b]] of pairs.entries()) {
  const startsMatch = String(a.starts) === String(b.starts);
  if (a.sheet === b.sheet) sameSheet++;
  if (a.glyphs === b.glyphs) sameGlyphs++;
  if (startsMatch) sameBreaks++;
  if (startsMatch && a.sheet === b.sheet && Math.abs(a.y - b.y) > .25) onlyPosition++;
  if (startsMatch && a.glyphs === b.glyphs) continue;
  changed.push({ ordinal: index + 1, font: a.font, size: a.size, rule: a.rule, tabs: a.tabs,
    reference: compact(a), candidate: compact(b) });
}
console.log('WINDOWS_WORD_LOGICAL_LINES ' + JSON.stringify({
  scale: process.argv[4], paired: pairs.length, sameSheet, sameBreaks, sameGlyphs, onlyPosition, changed: changed.slice(0, 20),
}));

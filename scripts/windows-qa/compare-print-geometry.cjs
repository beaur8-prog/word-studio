'use strict';

const fs = require('node:fs');
const tolerance = .25;
const before = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const after = JSON.parse(fs.readFileSync(process.argv[3], 'utf8'));
const keys = values => values.flatMap(sheet => [
  ...sheet.ink.map(([id]) => 'ink:' + id),
  ...sheet.objects.map(([shape]) => 'object:' + shape),
]);
const globalBefore = keys(before), globalAfter = keys(after);
const exactSequence = JSON.stringify(globalBefore) === JSON.stringify(globalAfter);
const result = { pages: [before.length, after.length], exactSequence, samePage: 0, maxDelta: 0,
  changedBoxes: 0, totalBoxes: 0, noCorrespondence: [] };
const compare = (a,b) => {
  const delta = Math.max(...a.map((value,i)=>Math.abs(value-b[i])));
  if (!Number.isFinite(delta)) throw Error('Invalid coordinate comparison');
  result.maxDelta = Math.max(result.maxDelta, delta);
  result.totalBoxes++;
  if (delta > tolerance) result.changedBoxes++;
  return delta <= tolerance;
};
for (const [i, original] of before.entries()) {
  const candidate = after[i];
  if (!candidate || JSON.stringify(keys([original])) !== JSON.stringify(keys([candidate])) ||
      original.ink.some(([_,rects], k) => rects.length !== candidate.ink[k][1].length)) {
    result.noCorrespondence.push(i + 1);
    continue;
  }
  let same = compare(original.paper,candidate.paper);
  for (const [j, entry] of original.ink.entries()) for (const [k,rect] of entry[1].entries())
    if (!compare(rect,candidate.ink[j][1][k])) same = false;
  for (const [j, entry] of original.objects.entries())
    if (!compare(entry[1],candidate.objects[j][1])) same = false;
  if (same) result.samePage++;
}
result.maxDelta=Math.round(result.maxDelta*1000)/1000;
console.log('WINDOWS_DENSITY_GEOMETRY ' + JSON.stringify({ scale:process.argv[4], tolerance, ...result }));

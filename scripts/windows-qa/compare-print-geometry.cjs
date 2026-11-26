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

const unique = sheets => {
  const paragraphs = new Map();
  for (const {index, paragraphs: entries} of sheets) for (const paragraph of entries ?? []) {
    paragraphs.set(paragraph.key, paragraphs.has(paragraph.key) ? null : {page:index, ...paragraph});
  }
  return paragraphs;
};
const beforeParagraphs = unique(before), afterParagraphs = unique(after);
const paired = [...beforeParagraphs.values()].filter(p=>p&&afterParagraphs.get(p.key)).map(p=>[p, afterParagraphs.get(p.key)]);
const sameStartPage = paired.filter(([a,b])=>a.page===b.page);
const changes = sameStartPage.filter(([a,b])=>a.lines!==b.lines||Math.max(...a.p.map((n,i)=>Math.abs(n-b.p[i])))>tolerance);
const compact = p => ({page:p.page,p:p.p.map(n=>Math.round(n*1000)/1000),y:p.y,lines:p.lines,font:p.font,size:p.size,lineHeight:p.lineHeight});
console.log('WINDOWS_DENSITY_PARAGRAPHS ' + JSON.stringify({scale:process.argv[4], paired:paired.length,
  sameStartPage:sameStartPage.length, sameLineLevel:sameStartPage.filter(([a,b])=>a.lines===b.lines).length,
  changed:changes.length, first:changes.slice(0,4).map(([a,b])=>({length:a.length,before:compact(a),after:compact(b)}))}));

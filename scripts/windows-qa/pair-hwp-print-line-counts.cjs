'use strict';

const { readFileSync } = require('node:fs');
const files = process.argv.slice(2);
if (files.length !== 2) throw Error('Expected the before and after line verdicts');
const [before, after] = files.map(file => new Map(JSON.parse(readFileSync(file, 'utf8')).map(row => [row.id, row])));
const common = [...before.keys()].filter(id => after.has(id));
if (!common.length) throw Error('No common visible paragraph');
const compare = field => [before, after].map(map => common.filter(id => map.get(id)[field]).length);
console.log('WINDOWS_HWP_COMMON_PRINT_LINES ' + JSON.stringify({
  eligible: common.length,
  individuallyEligible: [before.size, after.size],
  sameCount: compare('sameCount'),
  sameBreaks: compare('sameBreaks'),
  sameStartAndBreaks: compare('sameStartAndBreaks'),
}));

const { test } = require('node:test'), assert = require('node:assert/strict'), fixture = require('./hwp-empty-cell-terminal-fixture.cjs');
let samples; const ready = () => samples ??= fixture();
for (const format of ['hwp', 'hwpx']) for (const mode of ['empty', 'space', 'two']) test(`${format} ${mode}: nested unpainted terminal lines retain the ordinary cell policy`, async () => {
  const sample = await ready(), ordinary = sample[`${mode}-ordinary-${format}`].editable.html, nested = sample[`${mode}-nested-${format}`].editable.html;
  assert.equal((ordinary.match(/<table\b/g) || []).length, 1);
  assert.equal((nested.match(/<table\b/g) || []).length, 2, 'the native fixture really contains the nested table');
  const policies = html => [...html.matchAll(/--hwp-cell-line:([^;"']+)/g)].map(m => m[1]);
  assert.ok(policies(ordinary).length > 0, 'the ordinary source terminal line has a real end-height policy');
  const target = html => policies(html).filter(p => Math.abs(Number(p.split(/\s+/)[3]) - 40) < .1);
  assert.equal(target(ordinary).length, 1, 'the target 300% blank line policy is present');
  assert.deepEqual(target(nested), target(ordinary), 'same empty source line needs the same terminal height at every table depth');
  assert.equal((nested.match(/<br\b/g) || []).length, (ordinary.match(/<br\b/g) || []).length);
});

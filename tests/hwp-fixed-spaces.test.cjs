const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const ts = require('typescript');
const code = ts.transpileModule(fs.readFileSync(require('node:path').join(__dirname, '../src/hwp/fixed-spaces.ts'), 'utf8'), { compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.CommonJS } }).outputText;
const moduleValue = { exports: {} }; new Function('module', 'exports', code)(moduleValue, moduleValue.exports);
const { restoreHwpFixedSpaces } = moduleValue.exports;

test('only source fwSpace positions get quarter-em Unicode equivalents', () => {
  const html = '<p><span>가  나 다</span></p>';
  assert.equal(restoreHwpFixedSpaces(html, '가  나 다', [1, 2]), '<p><span>가  나 다</span></p>');
  assert.equal(restoreHwpFixedSpaces(html, '가  나 다', []), html);
});
test('UTF-16 positions survive supplementary characters, entities and inline styles', () => {
  const html = '<p>\n<span>😀&amp; </span><strong>&#8199;X</strong></p>';
  assert.equal(restoreHwpFixedSpaces(html, '😀&  X', [3, 4]), '<p>\n<span>😀&amp; </span><strong> X</strong></p>');
});
test('hard breaks and authored spaces keep their count and positions', () => {
  const html = '<p><span> A </span><br><span>  B </span></p>';
  const after = restoreHwpFixedSpaces(html, ' A \n  B ', [2, 5]);
  assert.equal(after, '<p><span> A </span><br><span>  B </span></p>');
});
test('ambiguous objects, missing text and invalid offsets remain unchanged', () => {
  for (const html of ['<p>A <img src="x"></p>', '<p>A B</p>', '<p>A <!-- rhwp: field --></p>'])
    assert.equal(restoreHwpFixedSpaces(html, 'A ', [1]), html);
  for (const position of [-1, 0, .5, 99]) assert.equal(restoreHwpFixedSpaces('<p>A </p>', 'A ', [position]), '<p>A </p>');
});

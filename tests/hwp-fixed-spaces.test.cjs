const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');

function load(file, resolve = name => name.startsWith('.')
  ? load(path.join(path.dirname(file), name + '.ts')) : require(name)) {
  const source = fs.readFileSync(path.join(__dirname, '../src', file), 'utf8');
  const code = ts.transpileModule(source, {
    compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.CommonJS },
  }).outputText;
  const module = { exports: {} };
  new Function('require', 'module', 'exports', code)(resolve, module, module.exports);
  return module.exports;
}

const advances = load('hwp-text-advances.ts');
const { restoreHwpFixedSpaces } = load('hwp/fixed-spaces.ts', name => {
  assert.equal(name, '../hwp-text-advances');
  return advances;
});

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

test('fixed-space conversion binds the existing source line evidence to the saved plain text', () => {
  const before = '가 나 다';
  const after = '가 나 다';
  const original = { fingerprint: advances.hwpWrapFingerprint(before), ends: [2, 5] };
  const tag = proof => '<p data-hwp-source-wrap="' + JSON.stringify(proof).replaceAll('"', '&quot;') + '">';
  const html = proof => tag(proof) + '<span>' + before + '</span></p>';
  assert.equal(restoreHwpFixedSpaces(html(original), before, [1]),
    tag({ ...original, fingerprint: advances.hwpWrapFingerprint(after) }) + '<span>' + after + '</span></p>');
  const stale = { ...original, fingerprint: advances.hwpWrapFingerprint('다른 텍스트') };
  assert.equal(restoreHwpFixedSpaces(html(stale), before, [1]), tag(stale) + '<span>' + after + '</span></p>');
});

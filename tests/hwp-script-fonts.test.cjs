const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const ts = require('typescript');
const moduleSource = ts.transpileModule(fs.readFileSync(require('node:path').join(__dirname, '../src/hwp/script-fonts.ts'), 'utf8'), {
  compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 },
}).outputText;
const loaded = { exports: {} };
new Function('module', 'exports', moduleSource)(loaded, loaded.exports);
const { restoreScriptFonts } = loaded.exports;
const fonts = ['Hangul Face', 'Latin Face', 'Hanja Face', 'Kana Face', 'Other Face', 'Symbol Face', 'User Face'];
const run = text => [{ start: 0, end: text.length, fonts }];

test('HWP script faces remain distinct while inherited size, tracking and bold stay intact', () => {
  const source = '한글 Abc 123, 漢字 カナ → ㆍ';
  const html = `<p><span style="font-family:Hangul Face;font-size:12pt;letter-spacing:-.5px"><b>${source}</b></span></p>`;
  const result = restoreScriptFonts(html, source, run(source));
  assert.match(result, /font-size:12pt;letter-spacing:-\.5px/);
  for (const face of ['Latin Face', 'Hanja Face', 'Kana Face', 'Symbol Face', 'User Face']) assert.ok(result.includes(face), face);
  assert.equal(result.replace(/<[^>]*>/g, ''), source);
  assert.equal((result.match(/<b>/g) || []).length, 1);
});

test('entities, explicit breaks and native HTML indentation preserve logical offsets', () => {
  const source = '한글 A&B\u00a0C\nNext';
  const html = '<p>\n<span>한글 A&amp;B&nbsp;C<br>Next</span></p>';
  const result = restoreScriptFonts(html, source, run(source));
  assert.match(result, /A&amp;B\u00a0C/);
  assert.match(result, /<br>/);
  assert.equal((result.match(/Latin Face/g) || []).length, 2);
});

test('source character shape boundaries and UTF-16 supplementary characters do not shift text', () => {
  const source = 'A𠀀B';
  const result = restoreScriptFonts(`<p><span>${source}</span></p>`, source, [
    { start: 0, end: 3, fonts }, { start: 3, end: 4, fonts: ['Hangul Face', 'Second Latin', ...fonts.slice(2)] },
  ]);
  assert.match(result, /Hanja Face/); assert.match(result, /Second Latin/);
  assert.equal(result.replace(/<[^>]*>/g, ''), source);
});

test('ambiguous control fragments and uniform fonts are left untouched', () => {
  const html = '<p><span>abc</span><img src="data:image/png;base64,AA=="></p>';
  assert.equal(restoreScriptFonts(html, 'abc\ufffc', run('abc\ufffc')), html);
  assert.equal(restoreScriptFonts(html, 'abc', [{ start: 0, end: 3, fonts: Array(7).fill('Arial') }]), html);
});

test('a font name cannot escape its CSS string or HTML attribute', () => {
  const result = restoreScriptFonts('<p><span>A</span></p>', 'A', [{ start: 0, end: 1, fonts: ['Arial', 'Bad";color:red;<img src=x>'] }]);
  assert.ok(!result.includes('<img'));
  assert.match(result, /&lt;img/);
  assert.match(result, /\\&quot;;color:red/);
});

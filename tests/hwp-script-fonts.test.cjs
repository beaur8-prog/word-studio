const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');

function load(source) {
  const code = ts.transpileModule(fs.readFileSync(source, 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 },
  }).outputText;
  const loaded = { exports: {} };
  const resolve = name => load(path.resolve(path.dirname(source), name + '.ts'));
  new Function('require', 'module', 'exports', code)(resolve, loaded, loaded.exports);
  return loaded.exports;
}

const { preserveHwpBodyLineEnds, preserveHwpSourceSpaces, restoreScriptFonts, restoreUniformLatinMetrics, sourceFontRuns } =
  load(path.join(__dirname, '../src/hwp/script-fonts.ts'));
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
  assert.equal((result.match(/font-family:&quot;Latin Face&quot;;/g) || []).length, 2);
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

test('text-only fragments after a control retain absolute source character-shape offsets', () => {
  const ranges = sourceFontRuns(JSON.stringify([
    { startOffset: 7, endOffset: 9, charShapeId: 4 }, { startOffset: 9, endOffset: 12, charShapeId: 5 },
  ]), 7, 'ABcde', offset => ({ fontFamilies: ['Hangul Face', offset === 7 ? 'First Latin' : 'Second Latin'] }));
  const result = restoreScriptFonts('<p><span>ABcde</span></p>', 'ABcde', ranges);
  assert.match(result, /First Latin[^>]*>AB<\/span>/);
  assert.match(result, /Second Latin[^>]*>cde<\/span>/);
});

test('native scalar offsets are converted to UTF-16 after supplementary Unicode', () => {
  const source = '😀A𠀀BC';
  const ranges = sourceFontRuns(JSON.stringify([
    { startOffset: 4, endOffset: 7, charShapeId: 1 }, { startOffset: 7, endOffset: 9, charShapeId: 2 },
  ]), 4, source, offset => ({ fontFamilies: ['Hangul Face', offset === 4 ? 'First Latin' : 'Second Latin', 'Hanja Face'] }));
  const result = restoreScriptFonts(`<p><span>${source}</span></p>`, source, ranges);
  assert.match(result, /First Latin[^>]*>A<\/span>/);
  assert.match(result, /Hanja Face[^>]*>𠀀<\/span>/);
  assert.match(result, /Second Latin[^>]*>BC<\/span>/);
  assert.equal(result.replace(/<[^>]*>/g, ''), source);
});

test('inline objects keep their original fragment even if visible text alone matches', () => {
  const html = '<p><span>abc</span><img src="data:image/png;base64,' + 'AAAA'.repeat(40000) + '"></p>';
  assert.equal(restoreScriptFonts(html, 'abc', run('abc')), html);
});


test('a known base script face does not inherit the native combined fallback list', () => {
  const source = '한글 Latin';
  const html = `<p><span style="font-family:Hangul Face,Latin Face;font-size:12pt;letter-spacing:-.5px">${source}</span></p>`;
  const result = restoreScriptFonts(html, source, run(source));
  assert.match(result, /font-family:&quot;Hangul Face&quot;;" data-hwp-latin-font="\[&quot;Hangul Face&quot;,&quot;Latin Face&quot;\]">한글 <\/span>/);
  assert.match(result, /font-family:&quot;Latin Face&quot;;">Latin<\/span>/);
  assert.match(result, /font-size:12pt;letter-spacing:-\.5px/);
  assert.equal(result.replace(/<[^>]*>/g, ''), source);
});

test('ordinary body runs recover the authored size even with one shared face', () => {
  const source = '한글 Latin', html = `<p><span style="font-size:11.5pt">${source}</span></p>`;
  const read = () => ({ fontFamilies: Array(7).fill('Arial'), fontSize: 1150, relativeSizes: [100, 106, 8, 100, 100, 100, 100] });
  const shape = JSON.stringify([{ startOffset: 0, endOffset: source.length, charShapeId: 4 }]);
  assert.equal(restoreScriptFonts(html, source, sourceFontRuns(shape, 0, source, read)), html);
  const result = restoreScriptFonts(html, source, sourceFontRuns(shape, 0, source, read, undefined, true));
  assert.match(result, /한글 <span data-hwp-script-size="12\.19pt">Latin<\/span>/);
  assert.equal(result.replace(/<[^>]*>/g, ''), source);
});

test('ordinary body runs recover language-specific tracking with the same face and size', () => {
  const source = '한글 Latin', html = `<p><span style="font-size:12pt;letter-spacing:-.6pt">${source}</span></p>`;
  const read = () => ({
    fontFamilies: Array(7).fill('Arial'), fontSize: 1200, relativeSizes: Array(7).fill(100),
    spacings: [-5, 3, -5, -5, -5, -5, -5],
  });
  const shapes = JSON.stringify([{ startOffset: 0, endOffset: source.length, charShapeId: 2 }]);
  const runs = sourceFontRuns(shapes, 0, source, read, undefined, true);
  const result = restoreScriptFonts(html, source, runs);
  assert.match(result, /letter-spacing:-\.6pt">한글 <span style="letter-spacing:0\.35pt;">Latin<\/span>/);
  assert.equal(result.replace(/<[^>]*>/g, ''), source);
  assert.equal(restoreScriptFonts('<p><sup>' + source + '</sup></p>', source, runs), '<p><sup>' + source + '</sup></p>');
});

test('a literal Latin cell corrects its rounded size and explicit zero tracking without changing other styles', () => {
  const source = 'First & last';
  const html = '<p><span style="font-family:Base;letter-spacing:-0.8px;font-size:11.5pt;color:red;">' +
    '<span style="font-family:Latin">First &amp; last</span></span></p>';
  const shape = { start: 0, end: source.length, fonts: ['Base', 'Latin'], base: 1150,
    sizes: [0, 12.19], tracking: [null, 0], widths: [95, 0] };
  const result = restoreUniformLatinMetrics(html, source, shape);
  assert.deepEqual(result, { size: 12, html: html.replace('letter-spacing:-0.8px;font-size:11.5pt', 'letter-spacing:0pt;font-size:12pt') });
  assert.equal(restoreUniformLatinMetrics(html, source, { ...shape, start: 1 }), undefined);
  const conflicting = html.replace('font-family:Latin', 'font-family:Latin;letter-spacing:.1px;');
  assert.equal(restoreUniformLatinMetrics(conflicting, source, shape), undefined);
  assert.equal(restoreUniformLatinMetrics(html, 'Other & last', shape), undefined);
  const quarterEm = html.replace('First ', 'First\u2005');
  assert.deepEqual(restoreUniformLatinMetrics(quarterEm, 'First\u2007& last', shape),
    { size: 12, html: result.html.replace('First ', 'First\u2005') });
  assert.equal(restoreUniformLatinMetrics(html, 'First\u2007& last', shape), undefined);
});

test('ordinary HWP spaces require a complete paragraph with only verified fixed-space shapes', () => {
  const text = '한글 사람', html = '<p><span>' + text + '</span></p>';
  const shape = [{ start: 0, end: 3, fonts, charShapeId: 0 }, { start: 3, end: text.length, fonts, charShapeId: 1 }];
  assert.equal(preserveHwpSourceSpaces(html, text, shape, [false, false]),
    '<p data-hwp-source-spaces="true"><span>' + text + '</span></p>');
  for (const flags of [undefined, [], [false], [false, true]]) assert.equal(preserveHwpSourceSpaces(html, text, shape, flags), html);
  assert.equal(preserveHwpSourceSpaces(html, text, shape.slice(1), [false, false]), html);
  const inline = html.replace('</p>', '<br></p>');
  assert.equal(preserveHwpSourceSpaces(inline, text, shape, [false, false]), inline);
});

test('stored HWP breaks validate natural wrapping without emitting break characters', () => {
  const text = '한글 문장', html = '<p><span>' + text + '</span></p>';
  const shape = [{ start: 0, end: text.length, fonts, charShapeId: 0 }];
  const stored = (textStarts, controls = 0) => () => JSON.stringify({ stored: true, controlCount: controls,
    lines: textStarts.map(textStart => ({ textStart })) });
  const result = preserveHwpSourceSpaces(html, text, shape, [false], stored([0, 3]));
  assert.match(result, /data-hwp-source-wrap=/);
  assert.match(result, /&quot;ends&quot;:\[2,4\]/);
  assert.equal(result.replace(/<[^>]*>/g, ''), text);
  const position = preserveHwpSourceSpaces(html, text, shape, [false], stored([0, 3]), { page: 2, x: 144.4, y: 80 });
  assert.match(position, /&quot;origin&quot;:{&quot;page&quot;:2,&quot;x&quot;:144\.4,&quot;y&quot;:80}/);
  assert.doesNotMatch(preserveHwpSourceSpaces(html, text, shape, [false], stored([0, 3], 1)), /data-hwp-source-wrap=/);
  assert.doesNotMatch(preserveHwpSourceSpaces(html, text, shape, [false], stored([0, 99])), /data-hwp-source-wrap=/);
});

test('a hanging body witness agrees with both stored and painted HWP starts, including leading spaces', () => {
  const text = '  첫 번째 줄과 다음 줄';
  const html = '<p><span>' + text + '</span></p>';
  const witness = { origin: { page: 2, x: 75.6, y: 240 }, hanging: 42,
    painted: [{ y: 240, charStart: 0 }, { y: 274, charStart: 10 }] };
  const stored = { stored: true, controlCount: 0, lines: [0, 10].map((textStart, index) => ({
    textStart, verticalPosition: index * 34, width: 642.5, height: 20, textHeight: 20, spacing: 14,
  })) };
  const read = () => JSON.stringify(stored);
  const witnessed = preserveHwpBodyLineEnds(html, text, read, witness);
  const metadata = markup => JSON.parse(markup.match(/data-hwp-source-wrap="([^"]+)"/)[1].replace(/&quot;/g, '"'));
  const { fingerprint, ends, authoredBodyLines, ...other } = metadata(witnessed);
  assert.deepEqual(authoredBodyLines, { origin:witness.origin, width:642.5, hanging:42, size:20, pitch:34 });
  assert.deepEqual(other, {});
  const old = { fingerprint, ends, bodyWidth:642.4, origin:{page:2,x:75.6,y:239.9} };
  const withProof = data => html.replace('<p>', '<p data-hwp-source-wrap="' + JSON.stringify(data).replace(/"/g, '&quot;') + '">');
  assert.deepEqual(metadata(preserveHwpBodyLineEnds(withProof(old), text, read, witness)), { ...old, authoredBodyLines });
  const conflicting = withProof({ ...old, ends:[ends[0]-1, ends[1]] });
  assert.equal(preserveHwpBodyLineEnds(conflicting, text, read, witness), conflicting);
  assert.equal(witnessed.replace(/<[^>]*>/g, ''), text);
  assert.equal(preserveHwpBodyLineEnds(html, text, read, {
    ...witness, painted: [{ y: 240, charStart: 2 }, witness.painted[1]],
  }), html);
  assert.equal(preserveHwpBodyLineEnds(html, text, read, {
    ...witness, painted: [witness.painted[0], { y: 276, charStart: 10 }],
  }), html);
});

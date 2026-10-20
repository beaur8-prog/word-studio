'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const source = fs.readFileSync(path.join(__dirname, '../src/hwp/viewport-layout.ts'), 'utf8');
const compiled = ts.transpileModule(source, { compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.CommonJS } }).outputText;
const bridge = { exports: {} }; new Function('module', 'exports', compiled)(bridge, bridge.exports);
const { hwpViewportLayout, hwpPageAtOffset, hwpVisiblePages } = bridge.exports;

test('mixed portrait/landscape paper has exact scroll positions and a stable current page across zoom levels', () => {
  const pages = [{ width: 794, height: 1122 }, { width: 1122, height: 794 }, { width: 816, height: 1056 }];
  for (const scale of [.5, .85, 1, 2]) {
    const layout = hwpViewportLayout(pages, scale);
    assert.equal(layout.width, 1122 * scale);
    for (let i = 0; i < pages.length; i++) {
      assert.equal(hwpPageAtOffset(layout, layout.offsets[i] + 10), i);
      assert.equal(hwpVisiblePages(layout, layout.offsets[i], Math.min(700, pages[i].height * scale)).current, i);
    }
    assert.equal(hwpPageAtOffset(layout, -10), 0);
    assert.equal(hwpPageAtOffset(layout, 100000), 2);
  }
});

test('a thousand-page scrollbar seeks directly to distant pages without treating every page as visible', () => {
  const layout = hwpViewportLayout(Array.from({ length: 1000 }, () => ({ width: 794, height: 1122 })), .85);
  for (const page of [0, 499, 998, 999]) {
    const range = hwpVisiblePages(layout, layout.offsets[page], 700);
    assert.equal(range.current, page); assert.ok(range.last - range.first <= 1);
  }
});

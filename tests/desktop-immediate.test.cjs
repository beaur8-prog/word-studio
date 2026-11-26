const { readFileSync } = require('node:fs');
const path = require('node:path');
const { runInNewContext } = require('node:vm');
const { MessageChannel } = require('node:worker_threads');
const assert = require('node:assert/strict');
const { test } = require('node:test');
const ts = require('typescript');

const source = readFileSync(path.join(__dirname, '../src/desktop-immediate.ts'), 'utf8');
const { outputText } = ts.transpileModule(source, { compilerOptions: { target: ts.ScriptTarget.ES2022 } });

function load(window, protocol) {
  const channels = [];
  class TrackedChannel extends MessageChannel {
    constructor() {
      super();
      channels.push(this);
    }
  }
  runInNewContext(outputText, { window, location: { protocol }, MessageChannel: TrackedChannel });
  return channels;
}

test('desktop file tasks are deferred, ordered, cancellable and receive their arguments', { timeout: 2000 }, async () => {
  const host = { wordDesktop: {} };
  const channels = load(host, 'file:');
  try {
    const calls = [];
    host.setImmediate((a, b) => calls.push([a, b]), 'zip part', 2);
    host.clearImmediate(host.setImmediate(() => calls.push('cancelled')));
    assert.deepEqual(calls, []);
    await new Promise(resolve => host.setImmediate(resolve));
    assert.deepEqual(calls, [['zip part', 2]]);
  } finally {
    for (const { port1, port2 } of channels) {
      port1.close();
      port2.close();
    }
  }
});

test('ordinary browser pages and an existing immediate implementation are left intact', () => {
  const current = { wordDesktop: {}, setImmediate, clearImmediate };
  for (const [host, protocol] of [[{}, 'file:'], [{ wordDesktop: {} }, 'http:'], [current, 'file:']]) {
    assert.deepEqual(load(host, protocol), []);
    assert.equal(host.setImmediate, host === current ? setImmediate : undefined);
    assert.equal(host.clearImmediate, host === current ? clearImmediate : undefined);
  }
});

'use strict';

const fs = require('node:fs');
const path = require('node:path');
const { TraceMap, originalPositionFor } = require('@jridgewell/trace-mapping');

module.exports = async function printProfile(page) {
  const session = await page.context().newCDPSession(page);
  await session.send('Profiler.enable');
  await session.send('Profiler.setSamplingInterval', { interval: 5000 });
  await session.send('Profiler.start');
  await page.evaluate(() => {
    const started = performance.now(), phases = [];
    window.qaPrintProfile = { started, phases };
    let state = '';
    const timer = setInterval(() => {
      const host = document.querySelector('.print-layout');
      const next = !host ? 'menu' : host.classList.contains('preview') ? 'preview' : 'measuring';
      if (next !== state) { phases.push({ name: next, at: Math.round(performance.now() - started) }); state = next; }
      if (next === 'preview') clearInterval(timer);
    }, 50);
  });
  return async () => {
    const { profile } = await session.send('Profiler.stop');
    const { phases, tasks } = await page.evaluate(() => ({
      phases: window.qaPrintProfile.phases,
      tasks: window.qaLongTasks.filter(t => t.at >= window.qaPrintProfile.started).map(t =>
        ({ at: Math.round(t.at - window.qaPrintProfile.started), ms: Math.round(t.duration) })),
    }));
    await session.detach();
    const maps = new Map();
    const sourceName = frame => {
      const name = frame.functionName || '(anonymous)';
      const bundle = frame.url?.match(/\/assets\/([^/]+\.js)(?:\?|$)/)?.[1];
      if (!bundle) return name;
      if (!maps.has(bundle)) {
        const source = path.join('dist', 'assets', bundle + '.map');
        maps.set(bundle, fs.existsSync(source) ? new TraceMap(fs.readFileSync(source, 'utf8')) : null);
      }
      const map = maps.get(bundle);
      if (!map) return name;
      const position = originalPositionFor(map, { line: frame.lineNumber + 1, column: frame.columnNumber });
      return position.source ? (position.name || name) + '@' + position.source.split('/').at(-1) + ':' + position.line : name;
    };
    const nodes = new Map(profile.nodes.map(node => [node.id, node]));
    const parents = new Map(profile.nodes.flatMap(node => (node.children || []).map(child => [child, node.id])));
    const intervals = [...tasks].sort((a,b) => b.ms - a.ms).slice(0,5);
    const measure = intervals.map(task => ({ ...task, stack: new Map() }));
    let elapsed = 0;
    for (let i = 0; i < (profile.samples || []).length; i++) {
      const micros = profile.timeDeltas[i]; elapsed += micros / 1000;
      const interval = measure.find(t => elapsed >= t.at && elapsed < t.at + t.ms);
      if (!interval) continue;
      const seen = new Set();
      for (let node = nodes.get(profile.samples[i]); node; node = nodes.get(parents.get(node.id))) {
        const label = sourceName(node.callFrame);
        if (!label.includes('@') || seen.has(label)) continue;
        seen.add(label);
        interval.stack.set(label, (interval.stack.get(label) || 0) + micros / 1000);
      }
    }
    const top = measure.map(({stack,...task}) => ({...task, frames:[...stack].sort((a,b)=>b[1]-a[1]).slice(0,12)
      .map(([name,millis]) => [name,Math.round(millis)])}));
    console.log('WINDOWS_NATIVE_PROFILE ' + JSON.stringify({phases, top}));
  };
};

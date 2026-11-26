type ImmediateTask = (...args: unknown[]) => void;

type ImmediateWindow = Window & {
  setImmediate?: (callback: ImmediateTask, ...args: unknown[]) => number;
  clearImmediate?: (id: number) => void;
};

// The ZIP dependencies queue separate tasks for chunks and entries. In the
// desktop file renderer their window-message polyfill can take minutes to
// reopen multi-section documents. Use a port owned by this window so those
// callbacks still run as separate tasks without relying on window messages.
const host: ImmediateWindow = window as Window;
if (window.wordDesktop && location.protocol === 'file:' && typeof host.setImmediate !== 'function') {
  const channel = new MessageChannel();
  const tasks = new Map<number, () => void>();
  let sequence = 0;
  channel.port1.onmessage = ({ data }: MessageEvent<number>) => {
    const task = tasks.get(data);
    tasks.delete(data);
    task?.();
  };
  host.setImmediate = (callback, ...args) => {
    if (typeof callback !== 'function') throw new TypeError('Immediate callback must be a function.');
    const id = ++sequence;
    tasks.set(id, () => callback(...args));
    channel.port2.postMessage(id);
    return id;
  };
  host.clearImmediate = id => { tasks.delete(id); };
}

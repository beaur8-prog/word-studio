import { androidServices, isAndroidRuntime } from './platform';

type Mode = 'native' | 'capacitor' | 'browser' | 'unavailable';
type Bars = { hide(): Promise<void>; show(): Promise<void> };
function capacitorBars(): Bars | undefined {
  const capacitor = (window as Window & { Capacitor?: { isNativePlatform?(): boolean; getPlatform?(): string; isPluginAvailable?(name: string): boolean; registerPlugin?<T>(name: string): T } }).Capacitor;
  if (capacitor?.isNativePlatform?.() && capacitor.getPlatform?.() === 'android' && capacitor.isPluginAvailable?.('SystemBars')) return capacitor.registerPlugin?.<Bars>('SystemBars');
}
/** A session owns only its fullscreen request and restores it on teardown. */
export function immersiveSession() {
  let mode: Mode = 'unavailable', requested = false;
  let pending: Promise<Mode> | undefined;
  const enter = (nativeOnly = false): Promise<Mode> => {
    if (!isAndroidRuntime()) return Promise.resolve('unavailable');
    if (mode === 'browser' && !document.fullscreenElement) { requested = false; pending = undefined; }
    if (requested && pending) return pending;
    const availableNative = androidServices()?.setImmersiveMode || capacitorBars();
    if (nativeOnly && !availableNative) return Promise.resolve('unavailable');
    requested = true;
    // Start the browser API synchronously in the trusted tap, before any
    // document parsing or awaiting would consume user activation.
    const native = androidServices();
    const bars = capacitorBars();
    pending = (async () => {
      try {
        if (native?.setImmersiveMode) { await native.setImmersiveMode(true); mode = 'native'; }
        else if (bars) { await bars.hide(); mode = 'capacitor'; }
        else if (!nativeOnly && document.documentElement.requestFullscreen) { await document.documentElement.requestFullscreen({ navigationUI: 'hide' }); mode = 'browser'; }
      } catch { mode = 'unavailable'; requested = false; }
      if (!requested) await restore();
      return mode;
    })();
    return pending;
  };
  const restore = async () => {
    const current = mode; mode = 'unavailable';
    try {
      if (current === 'native') await androidServices()?.setImmersiveMode?.(false);
      else if (current === 'capacitor') await capacitorBars()?.show();
      else if (current === 'browser' && document.fullscreenElement) await document.exitFullscreen();
    } catch { /* The OS may have already ended fullscreen. */ }
  };
  return { enter, async exit() { requested = false; await restore(); pending = undefined; } };
}

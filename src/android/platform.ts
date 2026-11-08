/** Android has its own entry point. Window width and touch support never opt a
 * desktop browser into this UI; an Electron bridge always wins over the UA. */
export function isAndroidRuntime() {
  if (window.wordDesktop) return false;
  const capacitor = (window as Window & { Capacitor?: { getPlatform?(): string; isNativePlatform?(): boolean } }).Capacitor;
  return !!(capacitor?.isNativePlatform?.() && capacitor.getPlatform?.() === 'android') || /\bAndroid\b/i.test(navigator.userAgent);
}

export type AndroidFile = { name: string; data: Uint8Array; type?: string };
export type AndroidResult = 'completed' | 'cancelled';
/** Optional native contract. No Electron API or unrestricted disk path crosses
 * this boundary. The host owns content:// grants, atomic writes and insets. */
export interface AndroidServices {
  openFile?(): Promise<AndroidFile | null>;
  saveFile?(file: AndroidFile): Promise<AndroidResult>;
  shareFile?(file: AndroidFile): Promise<AndroidResult>;
  onOpenFile?(receive: (file: AndroidFile) => void): () => void;
  onBack?(back: () => void): () => void;
  close?(): void;
  setImmersiveMode?(enabled: boolean): Promise<void>;
}
declare global { interface Window { wordAndroid?: AndroidServices } }
export const androidServices = () => isAndroidRuntime() ? window.wordAndroid : undefined;

export async function shareAndroidFile(file: File) {
  const services = androidServices();
  if (services?.shareFile) return services.shareFile({ name: file.name, type: file.type, data: new Uint8Array(await file.arrayBuffer()) });
  if (!navigator.canShare?.({ files: [file] }) || !navigator.share) throw new Error('이 환경에서는 파일 공유를 지원하지 않습니다. 사본 저장을 이용하세요.');
  try { await navigator.share({ files: [file], title: file.name }); return 'completed' as const; }
  catch (error) { if (error instanceof DOMException && error.name === 'AbortError') return 'cancelled' as const; throw error; }
}

export async function saveAndroidFile(file: File) {
  const services = androidServices();
  if (services?.saveFile) return services.saveFile({ name: file.name, type: file.type, data: new Uint8Array(await file.arrayBuffer()) });
  // A browser download has no write-completion acknowledgement. Do not label
  // it as a successful native save or clear the document's dirty state.
  const url = URL.createObjectURL(file), link = document.createElement('a');
  link.href = url; link.download = file.name; link.click();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
  return 'download-requested' as const;
}

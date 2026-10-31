export type NativeFileOperation = 'open' | 'save' | 'pdf' | 'print';
export type StudioUpdateAction = 'check' | 'download' | 'install' | 'cancel' | 'automatic' | 'release-page';
export interface StudioUpdateState {
  phase: 'idle' | 'checking' | 'available' | 'downloading' | 'verifying' | 'ready' | 'installing' | 'error';
  currentVersion: string;
  version: string | null;
  percent: number;
  transferred: number;
  total: number;
  checkedAt: number | null;
  result: 'latest' | 'available' | null;
  error: string | null;
  automatic: boolean;
  platform: string;
  packaged: boolean;
  supported: boolean;
}
declare global {
  interface Window {
    wordDesktop?: {
      paste?: () => Promise<void>;
      copyText?: (text: string) => Promise<void>;
      platform: string;
      getUpdateState?: () => Promise<StudioUpdateState>;
      updateAction?: (action: StudioUpdateAction, value?: boolean) => Promise<StudioUpdateState>;
      onUpdateState?: (callback: (state: StudioUpdateState) => void) => () => void;
      outlineFonts?: () => Promise<{ families: string[]; data: Uint8Array }[]>;
      readRecovery(): Promise<{ current: string; previous?: string; savedAt: string }[]>;
      saveRecovery(serialized: string): Promise<boolean>;
      saveFile(payload: {
        name: string;
        extension: "wordstudio" | "docx" | "html" | "txt" | "hwp" | "hwpx";
        data: Uint8Array;
      }): Promise<boolean>;
      openFile(): Promise<{ name: string; data: Uint8Array } | null>;
      onFileOperation?: (callback: (operation: NativeFileOperation | null) => void) => () => void;
      takeRequestedFile?: () => Promise<{ name: string; data: Uint8Array } | null>;
      onOpenRequested?: (callback: () => void) => () => void;
      exportPdf(payload: {
        name: string;
        header?: string;
        footer?: string;
        pageNumbers?: boolean;
        margin?: number;
        paginated?: boolean;
      }): Promise<boolean>;
      print(): Promise<boolean>;
      setTheme(theme: "dark" | "light"): void;
      onMenu(callback: (action: string) => void): () => void;
    };
  }
}

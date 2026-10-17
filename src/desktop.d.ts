export {};
declare global {
  interface Window {
    wordDesktop?: {
      paste?: () => Promise<void>;
      platform: string;
      readRecovery(): Promise<{ current: string; previous?: string; savedAt: string }[]>;
      saveRecovery(serialized: string): Promise<boolean>;
      saveFile(payload: {
        name: string;
        extension: "wordstudio" | "docx" | "html" | "txt";
        data: Uint8Array;
      }): Promise<boolean>;
      openFile(): Promise<{ name: string; data: Uint8Array } | null>;
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

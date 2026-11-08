import React from "react";
import { createRoot } from "react-dom/client";
import App from "./App";
import { isAndroidRuntime } from "./android/platform";
import "./styles.css";
import "./korean-font-fallbacks.css";
import { initializeRecovery } from './recovery';
import { initializeOutlineFonts } from './outline-fonts';
import { initializeWordFontFallbacks } from './word-font-fallbacks';
void Promise.all([initializeRecovery(), initializeOutlineFonts(), initializeWordFontFallbacks()]).then(async () => {
  const Root = isAndroidRuntime() ? (await import('./android/AndroidApp')).default : App;
  createRoot(document.getElementById("root")!).render(<React.StrictMode><Root /></React.StrictMode>);
});

import React from "react";
import { createRoot } from "react-dom/client";
import App from "./App";
import "./styles.css";
import { initializeRecovery } from './recovery';
import { initializeOutlineFonts } from './outline-fonts';
void Promise.all([initializeRecovery(), initializeOutlineFonts()]).then(() => createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
));

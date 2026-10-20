import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
export default defineConfig({
  plugins: [react()],
  base: "./",
  // The HWP module is first imported from a Worker. Discover it before serving
  // pages so Vite cannot reload the open document during the first HWP import.
  optimizeDeps: { include: ['@rhwp/core'] },
  server: { host: "127.0.0.1", port: 5191, strictPort: true },
});

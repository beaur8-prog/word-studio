import { defineConfig } from "@playwright/test";

export default defineConfig({
  testDir: "./tests",
  testMatch: '**/*.spec.ts',
  timeout: 45000,
  workers: 1,
  use: {
    baseURL: "http://127.0.0.1:5191",
    browserName: "chromium",
    channel: process.env.PLAYWRIGHT_CHANNEL || "chrome",
    viewport: { width: 1440, height: 1000 },
    screenshot: "only-on-failure",
    trace: "retain-on-failure",
  },
  webServer: {
    command: "npm run dev",
    url: "http://127.0.0.1:5191",
    reuseExistingServer: true,
    timeout: 30000,
  },
});

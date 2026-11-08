import { defineConfig } from '@playwright/test';
export default defineConfig({
  testDir: './tests', testMatch: 'android-ui.spec.ts', timeout: 60000, workers: 1,
  outputDir: 'test-results/android',
  use: { baseURL: 'http://127.0.0.1:5194', browserName: 'chromium', channel: process.env.PLAYWRIGHT_CHANNEL || 'chrome',
    viewport: { width: 412, height: 840 }, hasTouch: true,
    userAgent: 'Mozilla/5.0 (Linux; Android 14; Pixel 8) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Mobile Safari/537.36',
    screenshot: 'only-on-failure', trace: 'off' },
  webServer: { command: 'npm run dev -- --port 5194', url: 'http://127.0.0.1:5194', reuseExistingServer: false, timeout: 30000 },
});

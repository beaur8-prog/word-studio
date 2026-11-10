import { defineConfig } from '@playwright/test';
import base from '../../playwright.config';
export default defineConfig({ ...base, testDir: '../../tests', testMatch: ['oversized-word-rows.spec.ts','word-page-bottom-spacing.spec.ts'],
  use: { ...base.use, deviceScaleFactor: 2, trace: 'off' }, outputDir: './dpr2-test-results' });

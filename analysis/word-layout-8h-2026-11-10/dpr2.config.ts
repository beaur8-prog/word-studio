import {defineConfig} from '@playwright/test';
import base from '../../playwright.config';
export default defineConfig({...base,testDir:'../../tests',testMatch:['repeating-header-merged-flow.spec.ts','print-font-readiness.spec.ts','korean-gothic-fallbacks.spec.ts','section-geometry-layout.spec.ts','tab-prefix-pagination.spec.ts','oversized-word-rows.spec.ts','word-page-bottom-spacing.spec.ts'],
  use:{...base.use,baseURL:'http://127.0.0.1:5194',deviceScaleFactor:2,trace:'off'},
  webServer:{command:'npm run dev -- --port 5194',url:'http://127.0.0.1:5194',cwd:'/Users/swryu/Developer/word-studio',reuseExistingServer:true,timeout:30000},
  outputDir:'./dpr2-results'});

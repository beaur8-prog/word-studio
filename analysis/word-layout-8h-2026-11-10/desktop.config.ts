import {defineConfig} from '@playwright/test';
import base from '../../playwright.config';
const port=Number(process.env.LAYOUT_TEST_PORT||5193),url='http://127.0.0.1:'+port;
export default defineConfig({...base,testDir:'../../tests',testIgnore:'**/android-*.spec.ts',
  use:{...base.use,baseURL:url,trace:'off'},
  webServer:{command:'npm run dev -- --port '+port,url,cwd:'/Users/swryu/Developer/word-studio',reuseExistingServer:true,timeout:30000},
  outputDir:'./test-results'});

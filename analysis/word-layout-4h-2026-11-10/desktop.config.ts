import {defineConfig} from '@playwright/test';
import base from '../../playwright.config';
export default defineConfig({...base,testDir:'../../tests',testIgnore:'**/android-*.spec.ts',use:{...base.use,trace:'off'},outputDir:'./desktop-test-results'});

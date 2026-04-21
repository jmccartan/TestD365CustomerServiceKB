import { defineConfig } from '@playwright/test';
import * as path from 'path';
import * as fs from 'fs';
import 'dotenv/config';

// ---- Read saved settings written by global-setup ----------
const SETTINGS_FILE = path.resolve(__dirname, '.test-settings.json');
let d365Url = process.env.D365_URL || 'https://REPLACE_WITH_YOUR_ORG.crm.dynamics.com';

if (fs.existsSync(SETTINGS_FILE)) {
  try {
    const s = JSON.parse(fs.readFileSync(SETTINGS_FILE, 'utf-8'));
    d365Url = s.d365Url || d365Url;
  } catch { /* use default */ }
}

export default defineConfig({
  globalSetup: require.resolve('./global-setup'),
  testDir: './tests',
  timeout: 10 * 60 * 1000,
  expect: { timeout: 60_000 },
  fullyParallel: false,
  retries: 0,
  workers: 1,
  use: {
    baseURL: d365Url,
    storageState: path.resolve(__dirname, 'auth-state.json'),
    headless: false,
    viewport: { width: 1920, height: 1080 },
    actionTimeout: 30_000,
    navigationTimeout: 60_000,
    trace: 'on-first-retry',
    screenshot: 'only-on-failure',
  },
});

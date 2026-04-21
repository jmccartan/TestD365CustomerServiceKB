import { defineConfig } from '@playwright/test';
import * as path from 'path';
import * as fs from 'fs';
import 'dotenv/config';

const SETTINGS_FILE = path.resolve(__dirname, '.test-settings.json');
let d365Url = process.env.D365_URL || 'https://REPLACE_WITH_YOUR_ORG.crm.dynamics.com';

if (fs.existsSync(SETTINGS_FILE)) {
  try {
    const s = JSON.parse(fs.readFileSync(SETTINGS_FILE, 'utf-8'));
    d365Url = s.d365Url || d365Url;
  } catch { /* use default */ }
}

export default defineConfig({
  globalSetup: require.resolve('./global-setup-parallel'),
  globalTeardown: require.resolve('./global-teardown-parallel'),
  testDir: './tests',
  testMatch: 'd365-copilot-parallel.spec.ts',
  timeout: 5 * 60 * 1000,       // 5 min per individual prompt
  expect: { timeout: 60_000 },
  fullyParallel: true,
  retries: 1,                    // retry once on failure (e.g. transient load issues)
  workers: 3,                    // 3 parallel browser sessions
  use: {
    baseURL: d365Url,
    storageState: path.resolve(__dirname, 'auth-state.json'),
    headless: false,
    viewport: { width: 1920, height: 1080 },
    actionTimeout: 30_000,
    navigationTimeout: 120_000,
    trace: 'on-first-retry',
    screenshot: 'only-on-failure',
  },
});

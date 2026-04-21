import { defineConfig } from '@playwright/test';
import * as path from 'path';

export default defineConfig({
  testDir: './tests',
  timeout: 10 * 60 * 1000, // 10 minutes per test (handles all prompts)
  expect: { timeout: 60_000 },
  fullyParallel: false,
  retries: 0,
  workers: 1, // sequential — one browser session against D365
  use: {
    baseURL: process.env.D365_URL || 'https://REPLACE_WITH_YOUR_ORG.crm.dynamics.com',
    storageState: path.resolve(__dirname, 'auth-state.json'),
    headless: false,
    viewport: { width: 1920, height: 1080 },
    actionTimeout: 30_000,
    navigationTimeout: 60_000,
    trace: 'on-first-retry',
    screenshot: 'only-on-failure',
  },
});

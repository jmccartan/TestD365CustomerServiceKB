import { defineConfig } from '@playwright/test';
import * as path from 'path';
import * as fs from 'fs';
import 'dotenv/config';

// ---- Read saved settings written by global-setup ----------
const SETTINGS_FILE = path.resolve(__dirname, '.test-settings.json');

interface SavedSettings {
  d365Url: string;
  edgeProfile: string;
}

let settings: SavedSettings = {
  d365Url: process.env.D365_URL || 'https://REPLACE_WITH_YOUR_ORG.crm.dynamics.com',
  edgeProfile: '',
};

if (fs.existsSync(SETTINGS_FILE)) {
  try {
    settings = JSON.parse(fs.readFileSync(SETTINGS_FILE, 'utf-8'));
  } catch { /* use defaults */ }
}

const useEdgeProfile = settings.edgeProfile.length > 0;

export default defineConfig({
  globalSetup: require.resolve('./global-setup'),
  testDir: './tests',
  timeout: 10 * 60 * 1000, // 10 minutes per test (handles all prompts)
  expect: { timeout: 60_000 },
  fullyParallel: false,
  retries: 0,
  workers: 1,
  use: {
    baseURL: settings.d365Url,

    // When using an Edge profile the test launches a persistent
    // context itself (Playwright requires launchPersistentContext
    // for user-data-dir). Config just sets common defaults here.
    ...(useEdgeProfile
      ? { channel: 'msedge' }
      : { storageState: path.resolve(__dirname, 'auth-state.json') }),

    headless: false,
    viewport: { width: 1920, height: 1080 },
    actionTimeout: 30_000,
    navigationTimeout: 60_000,
    trace: 'on-first-retry',
    screenshot: 'only-on-failure',
  },
});

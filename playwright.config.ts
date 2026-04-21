import { defineConfig } from '@playwright/test';
import * as path from 'path';
import * as fs from 'fs';
import 'dotenv/config';

// ---- Read the profile choice written by global-setup ------
const CHOICE_FILE = path.resolve(__dirname, '.edge-profile-choice');
let edgeConfig: { profile: string; userDataDir: string } | null = null;

if (fs.existsSync(CHOICE_FILE)) {
  try {
    edgeConfig = JSON.parse(fs.readFileSync(CHOICE_FILE, 'utf-8'));
  } catch {
    // ignore — will fall back to storageState
  }
}

const useEdgeProfile = edgeConfig !== null;

export default defineConfig({
  globalSetup: require.resolve('./global-setup'),
  testDir: './tests',
  timeout: 10 * 60 * 1000, // 10 minutes per test (handles all prompts)
  expect: { timeout: 60_000 },
  fullyParallel: false,
  retries: 0,
  workers: 1, // sequential — one browser session against D365
  use: {
    baseURL: process.env.D365_URL || 'https://REPLACE_WITH_YOUR_ORG.crm.dynamics.com',

    // Auth: Edge profile takes priority; storageState is fallback
    ...(useEdgeProfile
      ? {
          channel: 'msedge',
          launchOptions: {
            args: [
              `--user-data-dir=${edgeConfig!.userDataDir}`,
              `--profile-directory=${edgeConfig!.profile}`,
            ],
          },
        }
      : {
          storageState: path.resolve(__dirname, 'auth-state.json'),
        }),

    headless: false,
    viewport: { width: 1920, height: 1080 },
    actionTimeout: 30_000,
    navigationTimeout: 60_000,
    trace: 'on-first-retry',
    screenshot: 'only-on-failure',
  },
});

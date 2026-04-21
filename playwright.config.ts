import { defineConfig } from '@playwright/test';
import * as path from 'path';
import * as fs from 'fs';
import * as os from 'os';

// ---- Edge profile vs storage-state auth -------------------
const EDGE_PROFILE = process.env.EDGE_PROFILE?.trim() || '';
const useEdgeProfile = EDGE_PROFILE.length > 0;

// Edge stores profiles under this root on Windows
const EDGE_USER_DATA = path.join(
  os.homedir(),
  'AppData', 'Local', 'Microsoft', 'Edge', 'User Data',
);

// Playwright can't share a live profile directory with an open
// Edge instance, so we copy it to a temp location at startup.
const EDGE_PROFILE_COPY = path.join(os.tmpdir(), 'pw-edge-profile');

function prepareEdgeProfile() {
  const src = path.join(EDGE_USER_DATA, EDGE_PROFILE);
  if (!fs.existsSync(src)) {
    throw new Error(
      `Edge profile "${EDGE_PROFILE}" not found at ${src}.\n` +
      'Open edge://version to find your profile directory name.',
    );
  }
  // Copy the full User Data dir (Playwright needs the root, not
  // just the profile subfolder) — we only copy the target profile
  // plus required root-level files to keep it fast.
  if (fs.existsSync(EDGE_PROFILE_COPY)) {
    fs.rmSync(EDGE_PROFILE_COPY, { recursive: true, force: true });
  }
  fs.mkdirSync(EDGE_PROFILE_COPY, { recursive: true });

  // Copy the profile subfolder
  fs.cpSync(src, path.join(EDGE_PROFILE_COPY, EDGE_PROFILE), { recursive: true });

  // Copy Local State (needed for cookie decryption)
  const localState = path.join(EDGE_USER_DATA, 'Local State');
  if (fs.existsSync(localState)) {
    fs.copyFileSync(localState, path.join(EDGE_PROFILE_COPY, 'Local State'));
  }
}

if (useEdgeProfile) {
  prepareEdgeProfile();
}

export default defineConfig({
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
              `--user-data-dir=${EDGE_PROFILE_COPY}`,
              `--profile-directory=${EDGE_PROFILE}`,
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

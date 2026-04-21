/**
 * Global setup — runs once before all tests.
 *
 * Interactive prompts (saved between runs in .test-settings.json):
 *   1. D365 URL — shown with last-used value; press Enter to keep it.
 *   2. Edge profile — numbered list; press Enter to reuse previous choice.
 *
 * If EDGE_PROFILE or D365_URL are set in .env they become the initial
 * defaults but can still be changed interactively.
 */

import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import * as readline from 'readline';
import 'dotenv/config';

const EDGE_USER_DATA = path.join(
  os.homedir(),
  'AppData', 'Local', 'Microsoft', 'Edge', 'User Data',
);
const SETTINGS_FILE = path.resolve(__dirname, '.test-settings.json');
const EDGE_PROFILE_COPY = path.join(os.tmpdir(), 'pw-edge-profile');

// ---------- types ----------
interface EdgeProfile {
  directory: string;
  displayName: string;
}

interface SavedSettings {
  d365Url: string;
  edgeProfile: string;  // directory name, or '' for fallback
}

// ---------- helpers ----------

function loadSettings(): SavedSettings {
  if (fs.existsSync(SETTINGS_FILE)) {
    try {
      return JSON.parse(fs.readFileSync(SETTINGS_FILE, 'utf-8'));
    } catch { /* ignore corrupt file */ }
  }
  return {
    d365Url: process.env.D365_URL || 'https://REPLACE_WITH_YOUR_ORG.crm.dynamics.com',
    edgeProfile: process.env.EDGE_PROFILE?.trim() || '',
  };
}

function saveSettings(settings: SavedSettings) {
  fs.writeFileSync(SETTINGS_FILE, JSON.stringify(settings, null, 2));
}

function ask(rl: readline.Interface, question: string): Promise<string> {
  return new Promise((resolve) => {
    rl.question(question, (answer) => resolve(answer));
  });
}

function discoverProfiles(): EdgeProfile[] {
  if (!fs.existsSync(EDGE_USER_DATA)) return [];
  const profiles: EdgeProfile[] = [];
  for (const entry of fs.readdirSync(EDGE_USER_DATA)) {
    const prefsPath = path.join(EDGE_USER_DATA, entry, 'Preferences');
    if (!fs.existsSync(prefsPath)) continue;
    try {
      const data = JSON.parse(fs.readFileSync(prefsPath, 'utf-8'));
      const displayName = data?.profile?.name || entry;
      profiles.push({ directory: entry, displayName });
    } catch { /* skip */ }
  }
  return profiles;
}

function prepareProfile(profileDir: string) {
  const src = path.join(EDGE_USER_DATA, profileDir);
  if (!fs.existsSync(src)) {
    throw new Error(`Edge profile "${profileDir}" not found at ${src}`);
  }

  // Copy the profile to a temp dir so we don't conflict with running Edge.
  // DPAPI cookie decryption still works because it's tied to the Windows
  // user account, not the file path.
  console.log('  Copying Edge profile (this may take a moment)...');

  if (fs.existsSync(EDGE_PROFILE_COPY)) {
    fs.rmSync(EDGE_PROFILE_COPY, { recursive: true, force: true });
  }
  fs.mkdirSync(EDGE_PROFILE_COPY, { recursive: true });

  // Copy the profile subfolder
  fs.cpSync(src, path.join(EDGE_PROFILE_COPY, profileDir), { recursive: true });

  // Copy Local State (contains the encryption key for cookies)
  const localState = path.join(EDGE_USER_DATA, 'Local State');
  if (fs.existsSync(localState)) {
    fs.copyFileSync(localState, path.join(EDGE_PROFILE_COPY, 'Local State'));
  }

  // Copy First Run sentinel so Edge doesn't show onboarding
  const firstRun = path.join(EDGE_USER_DATA, 'First Run');
  if (fs.existsSync(firstRun)) {
    fs.copyFileSync(firstRun, path.join(EDGE_PROFILE_COPY, 'First Run'));
  }
}

// ---------- main ----------

async function globalSetup() {
  const saved = loadSettings();
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

  // ── 1. D365 URL ──────────────────────────────────────────
  console.log('\n╔══════════════════════════════════════════╗');
  console.log('║         D365 Copilot Test Setup          ║');
  console.log('╚══════════════════════════════════════════╝');

  console.log(`\n  Current D365 URL: ${saved.d365Url}`);
  const urlInput = await ask(rl, '  Enter new URL or press Enter to keep: ');
  const d365Url = urlInput.trim() || saved.d365Url;

  // ── 2. Edge profile ──────────────────────────────────────
  const profiles = discoverProfiles();
  let edgeProfile = '';

  if (profiles.length > 0) {
    const savedIndex = profiles.findIndex((p) => p.directory === saved.edgeProfile);
    const defaultLabel = savedIndex >= 0
      ? `${profiles[savedIndex].displayName} (${profiles[savedIndex].directory})`
      : 'none';

    console.log(`\n  ℹ  The selected profile will be copied so Edge can remain open.`);
    console.log(`     If prompted to sign in, your first login will be remembered.\n`);
    console.log(`  Edge profiles found:`);
    profiles.forEach((p, i) => {
      const marker = p.directory === saved.edgeProfile ? ' ◄ current' : '';
      console.log(`  [${i + 1}]  ${p.displayName}  (${p.directory})${marker}`);
    });
    console.log(`  [0]  Skip — use auth-state.json fallback`);
    console.log(`\n  Default: ${defaultLabel}`);

    const profileInput = await ask(rl, '  Enter number or press Enter to keep default: ');
    const num = profileInput.trim() === '' ? -1 : parseInt(profileInput.trim(), 10);

    if (num === -1) {
      // keep previous
      edgeProfile = saved.edgeProfile;
    } else if (num > 0 && num <= profiles.length) {
      edgeProfile = profiles[num - 1].directory;
    } else {
      edgeProfile = '';
    }
  }

  rl.close();

  // ── Save settings for next run ───────────────────────────
  const settings: SavedSettings = { d365Url, edgeProfile };
  saveSettings(settings);

  // ── Apply Edge profile ───────────────────────────────────
  if (edgeProfile) {
    console.log(`\n→ D365 URL : ${d365Url}`);
    console.log(`→ Profile  : ${edgeProfile}`);
    prepareProfile(edgeProfile);
  } else {
    console.log(`\n→ D365 URL : ${d365Url}`);
    console.log('→ Profile  : none (using auth-state.json fallback)');
  }

  console.log('\n  ⏳ Starting tests — Edge will open shortly. This may take');
  console.log('     30–60 seconds to spin up, please be patient...\n');
}

export default globalSetup;

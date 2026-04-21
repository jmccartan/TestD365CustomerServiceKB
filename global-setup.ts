/**
 * Global setup — runs once before all tests.
 *
 * 1. Discovers Edge profiles on this machine.
 * 2. Prompts you to pick one (unless EDGE_PROFILE is already set in .env).
 * 3. Copies the chosen profile to a temp dir so it doesn't clash with a
 *    running Edge instance.
 * 4. Writes the selection to .edge-profile-choice so the config can read it.
 */

import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import * as readline from 'readline';

const EDGE_USER_DATA = path.join(
  os.homedir(),
  'AppData', 'Local', 'Microsoft', 'Edge', 'User Data',
);
const CHOICE_FILE = path.resolve(__dirname, '.edge-profile-choice');
const EDGE_PROFILE_COPY = path.join(os.tmpdir(), 'pw-edge-profile');

interface EdgeProfile {
  directory: string;   // e.g. "Default", "Profile 1"
  displayName: string; // e.g. "Work", "Personal"
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
    } catch {
      // skip unreadable profiles
    }
  }
  return profiles;
}

async function promptUser(profiles: EdgeProfile[]): Promise<string> {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

  console.log('\n╔══════════════════════════════════════════╗');
  console.log('║     Select an Edge profile for D365      ║');
  console.log('╚══════════════════════════════════════════╝\n');

  profiles.forEach((p, i) => {
    console.log(`  [${i + 1}]  ${p.displayName}  (${p.directory})`);
  });
  console.log(`  [0]  Skip — use auth-state.json fallback\n`);

  return new Promise((resolve) => {
    rl.question('Enter number: ', (answer) => {
      rl.close();
      const num = parseInt(answer.trim(), 10);
      if (num > 0 && num <= profiles.length) {
        resolve(profiles[num - 1].directory);
      } else {
        resolve('');
      }
    });
  });
}

function copyProfile(profileDir: string) {
  const src = path.join(EDGE_USER_DATA, profileDir);
  if (!fs.existsSync(src)) {
    throw new Error(`Edge profile "${profileDir}" not found at ${src}`);
  }

  if (fs.existsSync(EDGE_PROFILE_COPY)) {
    fs.rmSync(EDGE_PROFILE_COPY, { recursive: true, force: true });
  }
  fs.mkdirSync(EDGE_PROFILE_COPY, { recursive: true });

  // Copy the profile subfolder
  fs.cpSync(src, path.join(EDGE_PROFILE_COPY, profileDir), { recursive: true });

  // Copy Local State (needed for cookie decryption)
  const localState = path.join(EDGE_USER_DATA, 'Local State');
  if (fs.existsSync(localState)) {
    fs.copyFileSync(localState, path.join(EDGE_PROFILE_COPY, 'Local State'));
  }
}

async function globalSetup() {
  // If EDGE_PROFILE is hard-coded in .env, use it directly (no prompt)
  const envProfile = process.env.EDGE_PROFILE?.trim() || '';
  let chosenProfile = envProfile;

  if (!chosenProfile) {
    const profiles = discoverProfiles();
    if (profiles.length > 0) {
      chosenProfile = await promptUser(profiles);
    }
  }

  if (chosenProfile) {
    console.log(`\n→ Using Edge profile: ${chosenProfile}`);
    copyProfile(chosenProfile);
    // Persist the choice so playwright.config.ts can read it
    fs.writeFileSync(CHOICE_FILE, JSON.stringify({
      profile: chosenProfile,
      userDataDir: EDGE_PROFILE_COPY,
    }));
  } else {
    console.log('\n→ No Edge profile selected — falling back to auth-state.json');
    if (fs.existsSync(CHOICE_FILE)) fs.unlinkSync(CHOICE_FILE);
  }
}

export default globalSetup;

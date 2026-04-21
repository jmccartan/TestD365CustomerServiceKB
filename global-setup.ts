/**
 * Global setup — runs once before all tests.
 *
 * 1. Prompts for D365 URL (saved between runs).
 * 2. If no saved auth state, launches Chromium so you can sign in.
 *    Your session is saved to auth-state.json for future runs.
 */

import * as fs from 'fs';
import * as path from 'path';
import * as readline from 'readline';
import { chromium } from '@playwright/test';
import 'dotenv/config';

const SETTINGS_FILE = path.resolve(__dirname, '.test-settings.json');
const AUTH_STATE_FILE = path.resolve(__dirname, 'auth-state.json');

// ---------- types ----------
interface SavedSettings {
  d365Url: string;
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

function hasValidAuthState(): boolean {
  if (!fs.existsSync(AUTH_STATE_FILE)) return false;
  try {
    const data = JSON.parse(fs.readFileSync(AUTH_STATE_FILE, 'utf-8'));
    return data?.cookies?.length > 0 || data?.origins?.length > 0;
  } catch {
    return false;
  }
}

// ---------- main ----------

async function globalSetup() {
  const saved = loadSettings();

  // ── 0. Check Chromium is installed ───────────────────────
  const { chromium: pw } = require('@playwright/test');
  const chromiumPath: string = pw.executablePath();
  if (!fs.existsSync(chromiumPath)) {
    console.error('');
    console.error('  ┌──────────────────────────────────────────────────────┐');
    console.error('  │  Chromium is not installed!                          │');
    console.error('  │                                                      │');
    console.error('  │  Run this command first:                             │');
    console.error('  │    npx playwright install chromium                   │');
    console.error('  └──────────────────────────────────────────────────────┘');
    console.error('');
    process.exit(1);
  }

  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

  // ── 1. D365 URL ──────────────────────────────────────────
  console.log('\n╔══════════════════════════════════════════╗');
  console.log('║         D365 Copilot Test Setup          ║');
  console.log('╚══════════════════════════════════════════╝');

  console.log('\n  Tip: Use the full Customer Service workspace URL, e.g.:');
  console.log('  https://yourorg.crm.dynamics.com/main.aspx?appid=...');
  console.log(`\n  Current URL: ${saved.d365Url}`);
  const urlInput = await ask(rl, '  Enter new URL or press Enter to keep: ');
  const d365Url = urlInput.trim() || saved.d365Url;
  rl.close();

  // Save for next run
  saveSettings({ d365Url });

  console.log(`\n→ D365 URL: ${d365Url}`);

  // ── 2. Auth check ────────────────────────────────────────
  if (!hasValidAuthState()) {
    console.log('\n  No saved login found. Opening browser for you to sign in...');
    console.log('  Once you are fully logged into D365, press Enter in this terminal.\n');

    const browser = await chromium.launch({ headless: false });
    const context = await browser.newContext({ viewport: { width: 1920, height: 1080 } });
    const page = await context.newPage();

    await page.goto(d365Url, { waitUntil: 'domcontentloaded' });

    const rl2 = readline.createInterface({ input: process.stdin, output: process.stdout });
    await ask(rl2, '  ✅ Press Enter after you have signed in to D365... ');
    rl2.close();

    await context.storageState({ path: AUTH_STATE_FILE });
    console.log(`\n  Login saved to ${AUTH_STATE_FILE}`);
    console.log('  Future runs will skip this step.\n');

    await browser.close();
  } else {
    console.log('→ Auth    : using saved login (delete auth-state.json to re-login)');
  }

  console.log('\n  ⏳ Starting tests — browser will open shortly. This may take');
  console.log('     30–60 seconds to spin up, please be patient...\n');
}

export default globalSetup;

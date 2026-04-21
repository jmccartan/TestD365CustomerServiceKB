/**
 * AUTH SETUP — Run this once before tests to save your login session.
 *
 * Usage:
 *   npx playwright test auth-setup.ts --headed
 *
 * This opens a browser so you can manually sign in to D365.
 * Once signed in, press Enter in the terminal to save the session.
 * Subsequent test runs reuse the saved session (auth-state.json).
 */

import { chromium } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import * as readline from 'readline';

const D365_URL = process.env.D365_URL || 'https://REPLACE_WITH_YOUR_ORG.crm.dynamics.com';
const AUTH_STATE_PATH = path.resolve(__dirname, 'auth-state.json');

async function main() {
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext();
  const page = await context.newPage();

  console.log(`\nNavigating to: ${D365_URL}`);
  console.log('Please sign in to D365 in the browser window.\n');

  await page.goto(D365_URL, { waitUntil: 'domcontentloaded' });

  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  await new Promise<void>((resolve) => {
    rl.question('Press ENTER after you have signed in and the D365 page has fully loaded...', () => {
      rl.close();
      resolve();
    });
  });

  await context.storageState({ path: AUTH_STATE_PATH });
  console.log(`\nAuth state saved to: ${AUTH_STATE_PATH}`);
  console.log('You can now run: npx playwright test\n');

  await browser.close();
}

main().catch(console.error);

#!/usr/bin/env node
/**
 * Setup script for the Power Platform Playwright Toolkit.
 * Clones, installs, and builds the framework, then removes the
 * duplicate @playwright/test to avoid conflicts.
 *
 * Run: node setup-framework.js
 */
const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const FRAMEWORK_DIR = path.resolve(__dirname, '.pp-framework');
const TOOLKIT_DIR = path.join(FRAMEWORK_DIR, 'packages', 'power-platform-playwright-toolkit');
const REPO_URL = 'https://github.com/microsoft/power-platform-playwright-samples.git';

function run(cmd, cwd) {
  console.log(`> ${cmd}`);
  execSync(cmd, { cwd: cwd || __dirname, stdio: 'inherit' });
}

// 1. Clone if not present
if (!fs.existsSync(FRAMEWORK_DIR)) {
  console.log('\n📦 Cloning Power Platform Playwright Samples...');
  run(`git clone --depth 1 ${REPO_URL} .pp-framework`);
} else {
  console.log('\n✓ Framework already cloned');
}

// 2. Install toolkit dependencies
console.log('\n📦 Installing toolkit dependencies...');
run('npm install', TOOLKIT_DIR);

// 3. Build the toolkit
console.log('\n🔨 Building toolkit...');
run('npx tsc', TOOLKIT_DIR);

// 4. Remove duplicate @playwright/test (must use root project's copy)
console.log('\n🧹 Deduplicating @playwright/test...');
const dupes = [
  path.join(TOOLKIT_DIR, 'node_modules', 'playwright'),
  path.join(TOOLKIT_DIR, 'node_modules', '@playwright'),
];
for (const dir of dupes) {
  if (fs.existsSync(dir)) {
    fs.rmSync(dir, { recursive: true, force: true });
    console.log(`  Removed ${path.relative(__dirname, dir)}`);
  }
}

// 5. Install toolkit as local dependency in root project
console.log('\n📦 Linking toolkit to project...');
run(`npm install ${TOOLKIT_DIR}`);

console.log('\n✅ Framework setup complete!\n');

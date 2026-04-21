import { test, expect, Page } from '@playwright/test';
import * as path from 'path';
import ExcelJS from 'exceljs';

// ============================================================
// CONFIGURATION — reads from .test-settings.json (set during
// interactive setup) with .env / defaults as fallback.
// ============================================================
import * as fs from 'fs';

const SETTINGS_FILE = path.resolve(__dirname, '..', '.test-settings.json');
let savedUrl = '';
try {
  const s = JSON.parse(fs.readFileSync(SETTINGS_FILE, 'utf-8'));
  savedUrl = s.d365Url || '';
} catch { /* use fallback */ }

const D365_URL = savedUrl || process.env.D365_URL || 'https://REPLACE_WITH_YOUR_ORG.crm.dynamics.com';
const RESPONSE_TIMEOUT = parseInt(process.env.COPILOT_RESPONSE_TIMEOUT || '60', 10) * 1000;
const SIMILARITY_THRESHOLD = parseFloat(process.env.SIMILARITY_THRESHOLD || '0.6');

const INPUT_XLSX = path.resolve(__dirname, '..', 'Prompts and Responses.xlsx');
const OUTPUT_XLSX = path.resolve(__dirname, '..', `Test Results ${new Date().toISOString().slice(0, 10)}.xlsx`);

// ============================================================
// HELPERS
// ============================================================

interface PromptRow {
  prompt: string;
  expectedResponse: string;
  referencedDocs: string;
}

/** Read prompts from the source Excel file */
async function readPrompts(): Promise<PromptRow[]> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(INPUT_XLSX);
  const ws = wb.getWorksheet('Prompts & Responses') || wb.worksheets[0];
  const rows: PromptRow[] = [];

  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // skip header
    const prompt = row.getCell(1).text?.trim();
    const expectedResponse = row.getCell(2).text?.trim();
    const referencedDocs = row.getCell(3).text?.trim() || '';
    if (prompt) {
      rows.push({ prompt, expectedResponse, referencedDocs });
    }
  });

  return rows;
}

/** Simple cosine-ish similarity on word overlap (case-insensitive) */
function similarity(a: string, b: string): number {
  const tokenize = (s: string) => {
    const tokens = s.toLowerCase().replace(/[^a-z0-9\s]/g, '').split(/\s+/).filter(Boolean);
    return new Set(tokens);
  };
  const setA = tokenize(a);
  const setB = tokenize(b);
  if (setA.size === 0 || setB.size === 0) return 0;
  let intersection = 0;
  for (const word of setA) {
    if (setB.has(word)) intersection++;
  }
  return intersection / Math.max(setA.size, setB.size);
}

/** Write results to a new Excel file */
async function writeResults(results: TestResult[]) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Test Results');

  // Header row
  ws.columns = [
    { header: '#', key: 'index', width: 5 },
    { header: 'Prompt', key: 'prompt', width: 60 },
    { header: 'Expected Response', key: 'expected', width: 60 },
    { header: 'Actual Response', key: 'actual', width: 60 },
    { header: 'Similarity', key: 'similarity', width: 12 },
    { header: 'Result', key: 'result', width: 10 },
    { header: 'Referenced Docs', key: 'docs', width: 30 },
  ];

  // Style header
  ws.getRow(1).font = { bold: true };
  ws.getRow(1).alignment = { vertical: 'middle' };

  for (const r of results) {
    const row = ws.addRow({
      index: r.index,
      prompt: r.prompt,
      expected: r.expectedResponse,
      actual: r.actualResponse,
      similarity: `${(r.similarity * 100).toFixed(1)}%`,
      result: r.pass ? 'PASS' : 'FAIL',
      docs: r.referencedDocs,
    });

    // Color the result cell green/red
    const resultCell = row.getCell('result');
    resultCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: r.pass ? 'FF92D050' : 'FFFF4444' },
    };
    resultCell.font = { bold: true, color: { argb: r.pass ? 'FF006100' : 'FF9C0006' } };
  }

  // Summary row
  const passCount = results.filter((r) => r.pass).length;
  ws.addRow({});
  ws.addRow({
    index: '',
    prompt: `Total: ${results.length}  |  Pass: ${passCount}  |  Fail: ${results.length - passCount}`,
  });

  await wb.xlsx.writeFile(OUTPUT_XLSX);
}

interface TestResult {
  index: number;
  prompt: string;
  expectedResponse: string;
  actualResponse: string;
  similarity: number;
  pass: boolean;
  referencedDocs: string;
}

// ============================================================
// D365 COPILOT INTERACTION
//
// These selectors target the Copilot side-panel in D365
// Customer Service. Adjust if your environment differs.
// ============================================================

/** Try to locate the Copilot iframe (if the chat is embedded in one) */
async function getCopilotFrame(page: Page): Promise<Page | import('@playwright/test').Frame> {
  // D365 Copilot often lives inside an iframe
  const iframeSelectors = [
    'iframe[src*="copilot"]',
    'iframe[src*="omnichannelchat"]',
    'iframe[src*="webchat"]',
    'iframe[title*="Copilot"]',
    'iframe[title*="copilot"]',
    'iframe[id*="copilot"]',
    'iframe[name*="copilot"]',
  ];

  for (const sel of iframeSelectors) {
    const frameEl = page.locator(sel).first();
    if (await frameEl.isVisible({ timeout: 2000 }).catch(() => false)) {
      const frame = page.frame({ url: /copilot|omnichannelchat|webchat/i })
        || (await frameEl.contentFrame());
      if (frame) {
        console.log(`  Found Copilot iframe: ${sel}`);
        return frame;
      }
    }
  }

  // Also check all frames by URL pattern
  for (const frame of page.frames()) {
    const url = frame.url().toLowerCase();
    if (url.includes('copilot') || url.includes('webchat') || url.includes('omnichannelchat')) {
      console.log(`  Found Copilot frame by URL: ${url.slice(0, 80)}`);
      return frame;
    }
  }

  // No iframe found — Copilot is on the main page
  return page;
}

async function openCopilotPanel(page: Page) {
  const copilotButtonSelectors = [
    'button[aria-label*="Copilot"]',
    'button[title*="Copilot"]',
    '[data-id="msdyn_copilot"]',
    'button:has-text("Copilot")',
    '[data-id*="copilot" i]',
  ];

  for (const selector of copilotButtonSelectors) {
    const btn = page.locator(selector).first();
    if (await btn.isVisible({ timeout: 5000 }).catch(() => false)) {
      await btn.click();
      await page.waitForTimeout(2000);
      return;
    }
  }

  console.warn('  Could not find Copilot button — panel may already be open.');
}

async function sendPromptAndGetResponse(page: Page, prompt: string): Promise<string> {
  // Find the right context (main page or iframe)
  const frame = await getCopilotFrame(page);

  const inputSelectors = [
    'textarea[data-id*="copilot"]',
    'textarea[aria-label*="Type your message"]',
    'textarea[aria-label*="Ask a question"]',
    'textarea[aria-label*="Ask Copilot"]',
    'textarea[placeholder*="Ask"]',
    'textarea[placeholder*="Type"]',
    '[data-id="webchat-sendbox-input"]',
    'textarea[data-id="webchat-sendbox-input"]',
    'input[data-id="webchat-sendbox-input"]',
    // Generic fallbacks
    'textarea',
    'input[type="text"]',
  ];

  let input: ReturnType<typeof frame.locator> | null = null;
  for (const sel of inputSelectors) {
    const loc = frame.locator(sel).first();
    if (await loc.isVisible({ timeout: 3000 }).catch(() => false)) {
      input = loc;
      console.log(`  Using input selector: ${sel}`);
      break;
    }
  }

  if (!input) {
    // Dump diagnostic info to help identify the correct selectors
    console.error('\n  ── SELECTOR DEBUG INFO ──');
    console.error(`  Page URL: ${page.url()}`);
    console.error(`  Frames (${page.frames().length}):`);
    for (const f of page.frames()) {
      console.error(`    - ${f.url().slice(0, 100)}`);
    }
    const textareas = await frame.locator('textarea').count();
    const inputs = await frame.locator('input[type="text"]').count();
    console.error(`  Textareas in target frame: ${textareas}`);
    console.error(`  Text inputs in target frame: ${inputs}`);
    // List all textareas and their attributes
    const taInfo = await frame.locator('textarea').evaluateAll((els) =>
      els.map((el) => ({
        id: el.id,
        name: el.getAttribute('name'),
        ariaLabel: el.getAttribute('aria-label'),
        placeholder: el.getAttribute('placeholder'),
        dataId: el.getAttribute('data-id'),
        classes: el.className.slice(0, 60),
      }))
    );
    console.error('  Textarea details:', JSON.stringify(taInfo, null, 2));
    console.error('  ── END DEBUG INFO ──\n');

    throw new Error(
      'Could not find the Copilot chat input. See debug info above.'
    );
  }

  // Count existing messages before sending
  const messageContainerSelectors = [
    '[data-content="message-body"]',
    '.webchat__bubble__content',
    '[class*="message-content"]',
    '.ac-textBlock',
    '[role="listitem"]',
    '[role="log"] [role="group"]',
  ];

  let messageSelector = messageContainerSelectors[0];
  for (const sel of messageContainerSelectors) {
    if (await frame.locator(sel).first().isVisible({ timeout: 2000 }).catch(() => false)) {
      messageSelector = sel;
      break;
    }
  }

  const existingCount = await frame.locator(messageSelector).count();

  // Type and send the prompt
  await input.click();
  await input.fill(prompt);
  await page.keyboard.press('Enter');

  // Wait for a new bot response to appear
  await frame.waitForFunction(
    ({ selector, prevCount }) => {
      const msgs = document.querySelectorAll(selector);
      return msgs.length > prevCount + 1;
    },
    { selector: messageSelector, prevCount: existingCount },
    { timeout: RESPONSE_TIMEOUT }
  );

  // Wait for the response to finish streaming — keep checking until the
  // last message text stabilizes (no changes for 3 seconds)
  let lastText = '';
  let stableCount = 0;
  while (stableCount < 3) {
    await page.waitForTimeout(1000);
    const currentText = await frame.locator(messageSelector).last().innerText().catch(() => '');
    if (currentText === lastText && currentText.length > 0) {
      stableCount++;
    } else {
      stableCount = 0;
      lastText = currentText;
    }
  }

  // Also check for any typing/loading indicators to disappear
  const typingSelectors = [
    '[class*="typing"]',
    '[class*="loading"]',
    '[class*="spinner"]',
    '[aria-label*="typing"]',
    '[aria-label*="Thinking"]',
  ];
  for (const sel of typingSelectors) {
    const indicator = frame.locator(sel).first();
    if (await indicator.isVisible({ timeout: 500 }).catch(() => false)) {
      await indicator.waitFor({ state: 'hidden', timeout: RESPONSE_TIMEOUT }).catch(() => {});
    }
  }

  // Grab the last bot message
  const messages = frame.locator(messageSelector);
  const lastMessage = messages.last();
  const responseText = await lastMessage.innerText();

  return responseText.trim();
}

// ============================================================
// TEST
// ============================================================

test('D365 Copilot prompt regression test', async ({ page }) => {
  const prompts = await readPrompts();
  console.log(`\nLoaded ${prompts.length} prompts from: ${INPUT_XLSX}\n`);

  // Navigate to D365
  console.log(`Navigating to: ${D365_URL}\n`);
  await page.goto(D365_URL, { waitUntil: 'load', timeout: 120_000 });
  await page.waitForTimeout(5000);

  // Pause — let the user navigate to Customer Service workspace and open Copilot
  console.log('');
  console.log('  ┌──────────────────────────────────────────────────────┐');
  console.log('  │  Browser is open.                                    │');
  console.log('  │                                                      │');
  console.log('  │  1. Navigate to the Customer Service workspace       │');
  console.log('  │  2. Make sure the Copilot side panel is open         │');
  console.log('  │  3. Click the START TESTS button in the browser      │');
  console.log('  │                                                      │');
  console.log('  │  The URL will be saved automatically for next run.   │');
  console.log('  └──────────────────────────────────────────────────────┘');
  console.log('');

  // Inject a floating "Start Tests" button and wait for the user to click it
  await page.evaluate(() => {
    const btn = document.createElement('button');
    btn.id = 'pw-start-tests';
    btn.textContent = '▶  START TESTS';
    btn.style.cssText = `
      position: fixed; top: 10px; right: 10px; z-index: 999999;
      padding: 16px 32px; font-size: 18px; font-weight: bold;
      background: #107c10; color: white; border: none; border-radius: 8px;
      cursor: pointer; box-shadow: 0 4px 12px rgba(0,0,0,0.3);
    `;
    btn.onmouseover = () => btn.style.background = '#0b5e0b';
    btn.onmouseout = () => btn.style.background = '#107c10';
    btn.onclick = () => btn.remove();
    document.body.appendChild(btn);
  });

  // Wait until the button is clicked (removed from DOM)
  await page.waitForSelector('#pw-start-tests', { state: 'detached', timeout: 600_000 });
  console.log('  ✅ Starting tests...\n');

  // Capture the current URL from the browser and save it for next run
  const currentUrl = page.url();
  if (currentUrl && currentUrl !== D365_URL && !currentUrl.startsWith('about:')) {
    try {
      const settingsPath = path.resolve(__dirname, '..', '.test-settings.json');
      const settings = JSON.parse(fs.readFileSync(settingsPath, 'utf-8'));
      settings.d365Url = currentUrl;
      fs.writeFileSync(settingsPath, JSON.stringify(settings, null, 2));
      console.log(`  → URL saved for next run: ${currentUrl}\n`);
    } catch { /* ignore */ }
  }

  // Try to open Copilot panel (may already be open)
  await openCopilotPanel(page);
  await page.waitForTimeout(3000);

  const results: TestResult[] = [];

  for (let i = 0; i < prompts.length; i++) {
    const { prompt, expectedResponse, referencedDocs } = prompts[i];
    console.log(`[${i + 1}/${prompts.length}] Sending: ${prompt.slice(0, 80)}...`);

    let actualResponse = '';
    let sim = 0;
    let pass = false;

    try {
      actualResponse = await sendPromptAndGetResponse(page, prompt);
      sim = similarity(expectedResponse, actualResponse);
      pass = sim >= SIMILARITY_THRESHOLD;
      console.log(`  → Similarity: ${(sim * 100).toFixed(1)}% — ${pass ? 'PASS' : 'FAIL'}`);
    } catch (err: any) {
      actualResponse = `ERROR: ${err.message}`;
      console.error(`  → Error: ${err.message}`);
    }

    results.push({
      index: i + 1,
      prompt,
      expectedResponse,
      actualResponse,
      similarity: sim,
      pass,
      referencedDocs,
    });
  }

  // Write results Excel
  await writeResults(results);
  console.log(`\nResults written to: ${OUTPUT_XLSX}`);

  // Summary
  const passCount = results.filter((r) => r.pass).length;
  console.log(`\nSummary: ${passCount}/${results.length} passed (threshold: ${SIMILARITY_THRESHOLD * 100}%)\n`);
});

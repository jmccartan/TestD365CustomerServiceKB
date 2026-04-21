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
const now = new Date();
const timestamp = `${now.toISOString().slice(0, 10)}_${now.toTimeString().slice(0, 8).replace(/:/g, '-')}`;
const OUTPUT_XLSX = path.resolve(__dirname, '..', `Test Results ${timestamp}.xlsx`);

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

/** Search all frames for the one containing the Copilot chat input */
async function findCopilotInput(page: Page): Promise<{
  frame: Page | import('@playwright/test').Frame;
  input: ReturnType<Page['locator']>;
}> {
  // ── Shadow DOM selectors (D365 Copilot uses a <copilot-panel> web component) ──
  const shadowSelectors = [
    'copilot-panel >> input[placeholder*="Ask"]',
    'copilot-panel >> textarea[placeholder*="Ask"]',
    'copilot-panel >> input[placeholder*="question"]',
    'copilot-panel >> textarea',
    'copilot-panel >> input[type="text"]',
    '[data-testid="copilot-panel"] >> input[placeholder*="Ask"]',
    '[data-testid="copilot-panel"] >> textarea',
    '#msdyn_copilot >> input',
    '#msdyn_copilot >> textarea',
  ];

  // Regular DOM selectors (ordered from most specific to least)
  const regularSelectors = [
    'div[class*="CopilotRoot"] textarea',
    'textarea[aria-label="Type your message"]',
    'textarea[aria-label*="Type your message"]',
    'textarea[aria-label*="Ask a question"]',
    'textarea[aria-label*="Ask Copilot"]',
    'textarea[data-id*="copilot"]',
    'input[placeholder*="Ask a question"]',
    'input[placeholder*="Ask Copilot"]',
    '[data-id="webchat-sendbox-input"]',
    'textarea[data-id="webchat-sendbox-input"]',
    'input[data-id="webchat-sendbox-input"]',
    'textarea[placeholder*="Ask"]',
    'textarea[placeholder*="Type"]',
    'textarea[placeholder*="message"]',
    'textarea[role="textbox"]',
    'div[contenteditable="true"][aria-label*="message"]',
    'div[contenteditable="true"][aria-label*="Message"]',
    'div[contenteditable="true"][data-id*="copilot"]',
  ];

  const allSelectors = [...shadowSelectors, ...regularSelectors];

  // Search the main page and every iframe
  const framesToCheck: Array<Page | import('@playwright/test').Frame> = [page, ...page.frames()];

  for (const frame of framesToCheck) {
    for (const sel of allSelectors) {
      try {
        const loc = frame.locator(sel).first();
        if (await loc.isVisible({ timeout: 1000 }).catch(() => false)) {
          const frameUrl = 'url' in frame ? (frame as any).url() : page.url();
          console.log(`  Found input (${sel}) in frame: ${String(frameUrl).slice(0, 80)}`);
          return { frame, input: loc };
        }
      } catch { /* skip */ }
    }
  }

  // Last resort: find any visible textarea or contenteditable in any frame
  for (const frame of framesToCheck) {
    try {
      const taCount = await frame.locator('textarea').count();
      for (let i = 0; i < taCount; i++) {
        const ta = frame.locator('textarea').nth(i);
        if (await ta.isVisible({ timeout: 500 }).catch(() => false)) {
          const frameUrl = 'url' in frame ? (frame as any).url() : page.url();
          console.log(`  Found generic textarea in frame: ${String(frameUrl).slice(0, 80)}`);
          return { frame, input: ta };
        }
      }
      const ceCount = await frame.locator('div[contenteditable="true"]').count();
      for (let i = 0; i < ceCount; i++) {
        const ce = frame.locator('div[contenteditable="true"]').nth(i);
        if (await ce.isVisible({ timeout: 500 }).catch(() => false)) {
          const frameUrl = 'url' in frame ? (frame as any).url() : page.url();
          console.log(`  Found contenteditable div in frame: ${String(frameUrl).slice(0, 80)}`);
          return { frame, input: ce };
        }
      }
    } catch { /* skip */ }
  }

  // Dump comprehensive debug info
  console.error('\n  ── SELECTOR DEBUG INFO ──');
  console.error(`  Page URL: ${page.url()}`);
  console.error(`  Total frames: ${page.frames().length}`);

  // Check for shadow DOM hosts on each frame
  for (const frame of framesToCheck) {
    const url = 'url' in frame ? (frame as any).url() : page.url();
    if (String(url).includes('blank.htm')) continue;
    const taCount = await frame.locator('textarea').count().catch(() => 0);
    const inputCount = await frame.locator('input').count().catch(() => 0);
    const ceCount = await frame.locator('div[contenteditable="true"]').count().catch(() => 0);

    // Check for shadow DOM hosts
    const shadowHosts = await frame.evaluate(() => {
      const hosts: string[] = [];
      document.querySelectorAll('*').forEach((el) => {
        if (el.shadowRoot) {
          hosts.push(`<${el.tagName.toLowerCase()} id="${el.id}" class="${el.className}">`);
        }
      });
      return hosts;
    }).catch(() => [] as string[]);

    console.error(`\n  Frame: ${String(url).slice(0, 100)}`);
    console.error(`    textareas: ${taCount}, inputs: ${inputCount}, contenteditable: ${ceCount}`);
    if (shadowHosts.length > 0) {
      console.error(`    Shadow DOM hosts: ${JSON.stringify(shadowHosts)}`);
    }
    if (taCount > 0) {
      const taInfo = await frame.locator('textarea').evaluateAll((els) =>
        els.map((el) => ({
          visible: el.offsetParent !== null,
          ariaLabel: el.getAttribute('aria-label') || '(none)',
          placeholder: el.getAttribute('placeholder') || '(none)',
          dataId: el.getAttribute('data-id') || '(none)',
        }))
      ).catch(() => []);
      console.error('    textareas:', JSON.stringify(taInfo));
    }
  }
  console.error('  ── END DEBUG INFO ──\n');

  throw new Error('Could not find the Copilot chat input in any frame.');
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
  // Find the chat input across all frames
  const { frame, input } = await findCopilotInput(page);

  // Count existing messages before sending
  const messageContainerSelectors = [
    'div[class*="CopilotResponse"]',
    '[data-content="message-body"]',
    '.webchat__bubble__content',
    '[class*="message-content"]',
    '.ac-textBlock',
    '[role="listitem"]',
    '[role="log"] [role="group"]',
    // Fallback: any div that looks like a chat message
    '[class*="chat-message"]',
    '[class*="bot-message"]',
  ];

  let messageSelector = '';
  for (const sel of messageContainerSelectors) {
    const count = await frame.locator(sel).count().catch(() => 0);
    if (count > 0) {
      messageSelector = sel;
      console.log(`  Using message selector: ${sel} (${count} existing)`);
      break;
    }
  }

  // If no message selector found, we'll just wait by time
  const existingCount = messageSelector
    ? await frame.locator(messageSelector).count()
    : 0;

  // Type and send the prompt
  await input.click();
  await input.fill(prompt);
  await page.keyboard.press('Enter');

  // Wait for a new bot response to appear
  if (messageSelector) {
    await frame.waitForFunction(
      ({ selector, prevCount }) => {
        const msgs = document.querySelectorAll(selector);
        return msgs.length > prevCount + 1;
      },
      { selector: messageSelector, prevCount: existingCount },
      { timeout: RESPONSE_TIMEOUT }
    );
  } else {
    // No message selector found — just wait a fixed time
    console.log('  No message selector found — waiting for response by time...');
    await page.waitForTimeout(RESPONSE_TIMEOUT / 2);
  }

  // Wait for the response to finish streaming — keep checking until the
  // last message text stabilizes (no changes for 3 seconds)
  if (messageSelector) {
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
  if (messageSelector) {
    const messages = frame.locator(messageSelector);
    const lastMessage = messages.last();
    const responseText = await lastMessage.innerText();
    return responseText.trim();
  }

  // Fallback: try to get the last visible text block in the Copilot area
  const fallbackSelectors = [
    'div[class*="CopilotResponse"]',
    'div[class*="CopilotRoot"] div[class*="message"]',
    '[role="log"]',
  ];
  for (const sel of fallbackSelectors) {
    const el = frame.locator(sel).last();
    if (await el.isVisible({ timeout: 1000 }).catch(() => false)) {
      return (await el.innerText()).trim();
    }
  }

  return 'ERROR: Could not capture response — message selector not found.';
}

// ============================================================
// POPUP DISMISSAL
// ============================================================

async function dismissPopups(page: Page) {
  // Handle browser-level dialogs (permissions, alerts)
  page.on('dialog', async (dialog) => {
    console.log(`  Dismissing dialog: ${dialog.message().slice(0, 60)}`);
    await dialog.dismiss();
  });

  // Dismiss known D365/Copilot popups by clicking close/dismiss buttons
  const dismissSelectors = [
    // Microphone permission popups
    'button[aria-label*="Block"]',
    'button[aria-label*="Deny"]',
    'button[aria-label*="Don\'t allow"]',
    // "A copilot for you" and similar onboarding popups
    'button[aria-label*="Close"]',
    'button[aria-label*="Dismiss"]',
    'button[aria-label*="Got it"]',
    'button[aria-label*="Skip"]',
    'button:has-text("Got it")',
    'button:has-text("Skip")',
    'button:has-text("Close")',
    'button:has-text("Dismiss")',
    'button:has-text("No thanks")',
    'button:has-text("Maybe later")',
    // Generic close buttons on overlays/modals
    '[class*="dismiss"] button',
    '[class*="modal"] button[class*="close"]',
    '[role="dialog"] button[aria-label*="Close"]',
    '[role="dialog"] button[aria-label*="Dismiss"]',
  ];

  for (const sel of dismissSelectors) {
    try {
      const btn = page.locator(sel).first();
      if (await btn.isVisible({ timeout: 500 }).catch(() => false)) {
        const text = await btn.innerText().catch(() => sel);
        console.log(`  Dismissing popup: ${text.trim().slice(0, 40) || sel}`);
        await btn.click();
        await page.waitForTimeout(500);
      }
    } catch { /* ignore */ }
  }

  // Also check inside iframes for popups
  for (const frame of page.frames()) {
    for (const sel of ['button:has-text("Got it")', 'button:has-text("Close")', 'button:has-text("Dismiss")']) {
      try {
        const btn = frame.locator(sel).first();
        if (await btn.isVisible({ timeout: 300 }).catch(() => false)) {
          console.log(`  Dismissing popup in frame: ${sel}`);
          await btn.click();
          await page.waitForTimeout(500);
        }
      } catch { /* ignore */ }
    }
  }
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

  // Wait for D365 to fully initialize (no network activity for 2s)
  console.log('  Waiting for page to fully load...\n');
  await page.waitForLoadState('networkidle', { timeout: 60_000 }).catch(() => {
    console.log('  Network did not fully idle — continuing anyway.');
  });

  // Try to open Copilot panel (may already be open)
  await openCopilotPanel(page);
  await page.waitForTimeout(3000);

  // Dismiss popups multiple times with pauses to catch late-loading ones
  for (let i = 0; i < 3; i++) {
    await dismissPopups(page);
    await page.waitForTimeout(2000);
  }

  // Pause — let the user verify the page is ready
  console.log('');
  console.log('  ┌──────────────────────────────────────────────────────┐');
  console.log('  │  Browser is open.                                    │');
  console.log('  │                                                      │');
  console.log('  │  1. Navigate to the Customer Service workspace       │');
  console.log('  │  2. Make sure the Copilot side panel is open         │');
  console.log('  │  3. Click the button in the browser when ready       │');
  console.log('  │                                                      │');
  console.log('  │  The URL will be saved automatically for next run.   │');
  console.log('  └──────────────────────────────────────────────────────┘');
  console.log('');

  // Inject a floating button and wait for the user to click it
  await page.evaluate(() => {
    const btn = document.createElement('button');
    btn.id = 'pw-start-tests';
    btn.textContent = '▶  Before starting tests, ensure the page is ready';
    btn.style.cssText = `
      position: fixed; top: 10px; right: 10px; z-index: 999999;
      padding: 16px 32px; font-size: 16px; font-weight: bold;
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

  // Dismiss any new popups that appeared while waiting
  await dismissPopups(page);

  const results: TestResult[] = [];

  for (let i = 0; i < prompts.length; i++) {
    // Dismiss popups that may appear between prompts
    await dismissPopups(page);

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

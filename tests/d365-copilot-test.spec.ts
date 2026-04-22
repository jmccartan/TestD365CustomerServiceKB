import { test, expect, Page, BrowserContext } from '@playwright/test';
import * as path from 'path';
import ExcelJS from 'exceljs';
import {
  AppProvider,
  AppType,
  AppLaunchMode,
  ModelDrivenAppPage,
  waitForSpinnerToDisappear,
  handleDialog,
} from 'power-platform-playwright-toolkit';

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
    { header: 'Cited Sources', key: 'citedSources', width: 40 },
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
      citedSources: r.citedSources,
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
  citedSources: string;
}

// ============================================================
// D365 COPILOT INTERACTION
//
// The Copilot side-panel in D365 Customer Service uses a Fluent
// AI EditorInput component: <span role="textbox" aria-label=
// "Describe what you need">. Messages appear inside the panel's
// container (not in an iframe).
// ============================================================

// The Copilot panel container that holds all Copilot UI
const COPILOT_CONTAINER_SEL =
  '[data-id="MscrmControls.CSIntelligence.AICopilotControl_container"]';

/**
 * Locate the Copilot chat input on the main page.
 * The D365 Copilot uses a <span role="textbox"> from Fluent AI —
 * not a <textarea> or <input>.
 */
async function findCopilotInput(page: Page): Promise<ReturnType<Page['locator']>> {
  // Selectors ordered from most specific to least — all on the main page
  const selectors = [
    // Fluent AI EditorInput (the actual D365 Copilot input)
    `${COPILOT_CONTAINER_SEL} [role="textbox"][aria-label="Describe what you need"]`,
    `${COPILOT_CONTAINER_SEL} [role="textbox"]`,
    `${COPILOT_CONTAINER_SEL} [class*="EditorInput"] [role="textbox"]`,
    // Fallback: any textbox-like element inside the Copilot container
    `${COPILOT_CONTAINER_SEL} textarea`,
    `${COPILOT_CONTAINER_SEL} input[type="text"]`,
    `${COPILOT_CONTAINER_SEL} [contenteditable="true"]`,
  ];

  for (const sel of selectors) {
    const loc = page.locator(sel).first();
    if (await loc.isVisible({ timeout: 2000 }).catch(() => false)) {
      console.log(`  Found Copilot input: ${sel}`);
      return loc;
    }
  }

  // Debug: dump what's in the container
  console.error('\n  ── COPILOT INPUT DEBUG ──');
  const container = page.locator(COPILOT_CONTAINER_SEL);
  const exists = await container.count().catch(() => 0);
  console.error(`  Container exists: ${exists > 0}`);
  if (exists > 0) {
    const els = await container.evaluate((el) => {
      const results: string[] = [];
      el.querySelectorAll('[role="textbox"], textarea, input, [contenteditable]').forEach((e) => {
        const r = e.getBoundingClientRect();
        results.push(
          `<${e.tagName.toLowerCase()} role="${e.getAttribute('role')}" ` +
          `aria-label="${e.getAttribute('aria-label')}" ` +
          `visible=${r.width > 0 && r.height > 0}>`
        );
      });
      return results;
    }).catch(() => []);
    console.error(`  Input-like elements:`, els);
  }
  console.error('  ── END DEBUG ──\n');

  throw new Error('Could not find the Copilot chat input.');
}

/**
 * Send a prompt to the Copilot panel and capture the response.
 *
 * Strategy:
 *  1. Snapshot all text in the Copilot container before sending.
 *  2. Type the prompt into the Fluent EditorInput and click Send.
 *  3. Wait for new text to appear (the response).
 *  4. Wait for the response to stabilize (stop changing).
 *  5. Return only the NEW text that appeared after sending.
 */
async function sendPromptAndGetResponse(page: Page, prompt: string): Promise<string> {
  const container = page.locator(COPILOT_CONTAINER_SEL);

  // Snapshot text before sending
  const textBefore = await container.innerText().catch(() => '');

  // Find and populate the input
  const input = await findCopilotInput(page);
  await input.click();

  // The Fluent EditorInput is a <span role="textbox"> — fill() may not
  // work, so use pressSequentially() (types character by character) with
  // a fill() attempt first.
  try {
    await input.fill(prompt);
  } catch {
    // fill() not supported on this element — type character by character
    await input.pressSequentially(prompt, { delay: 20 });
  }

  // Click the Send button (more reliable than pressing Enter on a span)
  const sendBtn = page.locator(`${COPILOT_CONTAINER_SEL} button[aria-label="Send"]`).first();
  if (await sendBtn.isVisible({ timeout: 2000 }).catch(() => false)) {
    await sendBtn.click();
    console.log('  Clicked Send button');
  } else {
    // Fallback: press Enter
    await input.press('Enter');
    console.log('  Pressed Enter (Send button not found)');
  }

  // Wait for the container text to change (new response appeared)
  console.log('  Waiting for response...');
  const responseDeadline = Date.now() + RESPONSE_TIMEOUT;
  let currentText = textBefore;
  while (currentText === textBefore && Date.now() < responseDeadline) {
    await page.waitForTimeout(1000);
    currentText = await container.innerText().catch(() => '');
    // Break early if chat limit was hit
    if (currentText.includes("It's time to clear the chat")) {
      return 'CHAT_LIMIT_REACHED';
    }
  }

  if (currentText === textBefore) {
    return 'ERROR: No response appeared within timeout.';
  }

  // Wait for the response to finish streaming — text must stabilize
  // (no changes for 3 consecutive seconds), capped at 90s.
  let lastText = currentText;
  let stableCount = 0;
  const stabilityDeadline = Date.now() + 90_000;
  while (stableCount < 3 && Date.now() < stabilityDeadline) {
    await page.waitForTimeout(1000);
    currentText = await container.innerText().catch(() => '');
    if (currentText === lastText) {
      stableCount++;
    } else {
      stableCount = 0;
      lastText = currentText;
    }
  }

  // Extract only the new text (response) by removing the pre-send content.
  // The container text is structured: existing messages + new prompt echo + response.
  // We grab everything after the last occurrence of the prompt.
  const fullText = currentText;
  const promptIdx = fullText.lastIndexOf(prompt);
  if (promptIdx !== -1) {
    const afterPrompt = fullText.slice(promptIdx + prompt.length).trim();
    if (afterPrompt.length > 0) return afterPrompt;
  }

  // Fallback: return the delta between before and after
  if (fullText.length > textBefore.length) {
    return fullText.slice(textBefore.length).trim();
  }

  return fullText.trim();
}

/**
 * Expand the last "Check sources" accordion in the Copilot panel and
 * return the cited sources. If KB sources are cited, returns their names
 * comma-separated. If no KB sources, returns the disclaimer text
 * (e.g. "This response didn't come from your knowledge sources...").
 */
async function extractCitedSources(page: Page): Promise<string> {
  const container = page.locator(COPILOT_CONTAINER_SEL);

  // Find the last "Check sources" accordion button and expand it
  const accordionBtns = container.locator(
    'button[aria-expanded]:has-text("Check sources"), .fui-AccordionHeader button:has-text("Check sources")'
  );
  const count = await accordionBtns.count().catch(() => 0);
  if (count === 0) return '';

  const lastBtn = accordionBtns.last();
  const wasExpanded = (await lastBtn.getAttribute('aria-expanded').catch(() => 'false')) === 'true';
  if (!wasExpanded) {
    await lastBtn.click();
    await page.waitForTimeout(1500);
  }

  let result = '';

  // Check for citation links first (aria-label="Citation N Title")
  const citations = container.locator('a[aria-label^="Citation"]');
  const citCount = await citations.count().catch(() => 0);

  if (citCount > 0) {
    const sources: string[] = [];
    for (let i = 0; i < citCount; i++) {
      const text = await citations.nth(i).innerText().catch(() => '');
      if (text.trim()) sources.push(text.trim());
    }
    result = sources.join(', ');
  } else {
    // No citation links — grab the text from the expanded accordion panel.
    // The AccordionItem panel sits next to the header; grab the panel content.
    const accordionItem = container.locator('.fui-AccordionItem').last();
    const panelText = await accordionItem.innerText().catch(() => '');
    // Remove the "Check sources" header text itself
    const cleaned = panelText.replace(/Check sources/i, '').trim();
    if (cleaned) result = cleaned;
  }

  // Collapse it again to keep the panel tidy for the next prompt
  if (!wasExpanded) {
    await lastBtn.click().catch(() => {});
    await page.waitForTimeout(500);
  }

  return result;
}

/**
 * Detect and handle the D365 Copilot "15 of 15" conversation limit.
 * When the limit is hit, a "Clear chat" link appears. Click it to
 * reset the conversation so testing can continue.
 * Returns true if the chat was cleared.
 */
async function clearChatIfNeeded(page: Page): Promise<boolean> {
  const container = page.locator(COPILOT_CONTAINER_SEL);
  const containerText = await container.innerText().catch(() => '');

  // Only trigger on the specific limit message, not a generic "Clear chat" button
  if (!containerText.includes("It's time to clear the chat")) {
    return false;
  }

  console.log('  ⚠ Chat limit reached — clearing conversation...');

  // Click the "Clear chat" link (it's an <a> or <button> with that text)
  const clearLink = container.locator('a:has-text("Clear chat"), button:has-text("Clear chat")').first();
  if (await clearLink.isVisible({ timeout: 3000 }).catch(() => false)) {
    await clearLink.click();
    await page.waitForTimeout(3000);

    // Wait for the textbox to reappear (panel resets)
    await page.locator(`${COPILOT_CONTAINER_SEL} [role="textbox"]`)
      .waitFor({ state: 'visible', timeout: 15_000 })
      .catch(() => {});

    console.log('  ✓ Chat cleared — continuing tests');
    return true;
  }

  console.warn('  Could not find Clear chat link');
  return false;
}

// ============================================================
// D365 PAGE READINESS & POPUP DISMISSAL
//
// Uses the Power Platform Playwright Toolkit (AppProvider +
// ModelDrivenAppPage) for app launch and SPA readiness, with
// custom Copilot-specific waits layered on top.
// ============================================================

/**
 * Launch the D365 Customer Service app using the PP Playwright Toolkit.
 * Handles OAuth redirects, domain transitions, and SPA initialization.
 * Returns the ModelDrivenAppPage for further navigation if needed.
 */
async function launchD365App(page: Page, context: BrowserContext): Promise<ModelDrivenAppPage> {
  const app = new AppProvider(page, context);

  console.log('  Launching D365 via Power Platform Playwright Toolkit...');
  await app.launch({
    app: 'Customer Service workspace',
    type: AppType.ModelDriven,
    mode: AppLaunchMode.Play,
    skipMakerPortal: true,
    directUrl: D365_URL,
  });

  const mda = app.getModelDrivenAppPage();
  console.log('  ✓ D365 app launched');
  return mda;
}

/**
 * Wait for the Copilot panel to fully load inside the D365 app.
 * The framework handles D365 SPA readiness; this adds Copilot-specific waits.
 */
async function waitForCopilotReady(page: Page) {
  // Wait for the Copilot panel container to be present
  console.log('  Waiting for Copilot panel...');
  await page.locator(COPILOT_CONTAINER_SEL)
    .waitFor({ state: 'attached', timeout: 60_000 });

  // Wait for the Copilot panel to finish loading (textbox visible)
  console.log('  Waiting for Copilot panel content...');
  const textbox = page.locator(`${COPILOT_CONTAINER_SEL} [role="textbox"]`);
  await textbox.waitFor({ state: 'visible', timeout: 60_000 });
  console.log('  ✓ Copilot panel ready');
}

/**
 * Ensure the Copilot side panel is open by clicking the Copilot
 * tab button if the panel container is not yet visible.
 */
async function openCopilotPanel(page: Page) {
  const container = page.locator(COPILOT_CONTAINER_SEL);
  if (await container.isVisible({ timeout: 3000 }).catch(() => false)) {
    console.log('  Copilot panel already open');
    return;
  }

  const tabButton = page.locator(
    '[data-id*="sidepane-tab-button-AppSidePane_MscrmControls.CSIntelligence.AICopilotControl"] button'
  ).first();
  if (await tabButton.isVisible({ timeout: 5000 }).catch(() => false)) {
    await tabButton.click();
    console.log('  Clicked Copilot tab button');
    await page.waitForTimeout(2000);
    return;
  }

  const copilotBtn = page.locator('button[aria-label="Copilot"]').first();
  if (await copilotBtn.isVisible({ timeout: 3000 }).catch(() => false)) {
    await copilotBtn.click();
    console.log('  Clicked Copilot button');
    await page.waitForTimeout(2000);
    return;
  }

  console.warn('  Could not find Copilot button — panel may already be open.');
}

/**
 * Dismiss D365 popups/dialogs. Searches main page and all iframes.
 * Targets the specific "A Copilot for you!" onboarding dialog and
 * any other role="dialog" overlays.
 */
async function dismissPopups(page: Page) {
  const framesToCheck: Array<Page | import('@playwright/test').Frame> = [page, ...page.frames()];

  // Target dismiss/close buttons inside visible dialogs first (most specific)
  const dialogDismissSelectors = [
    '[role="dialog"] button[aria-label="Dismiss"]',
    '[role="dialog"] button[aria-label="Close"]',
    '[role="alertdialog"] button[aria-label="Dismiss"]',
    '[role="alertdialog"] button[aria-label="Close"]',
  ];

  // Then general dismiss-like buttons
  const generalDismissSelectors = [
    'button[aria-label="Dismiss"]',
    'button[aria-label="Close"]',
    'button[aria-label*="Got it"]',
    'button[aria-label*="Skip"]',
    'button:has-text("Got it")',
    'button:has-text("Skip")',
    'button:has-text("Dismiss")',
    'button:has-text("No thanks")',
    'button:has-text("Maybe later")',
  ];

  const allSelectors = [...dialogDismissSelectors, ...generalDismissSelectors];

  for (const frame of framesToCheck) {
    for (const sel of allSelectors) {
      try {
        const btns = frame.locator(sel);
        const count = await btns.count();
        for (let i = 0; i < count; i++) {
          const btn = btns.nth(i);
          if (await btn.isVisible({ timeout: 300 }).catch(() => false)) {
            const label = await btn.getAttribute('aria-label').catch(() => '');
            const text = await btn.innerText().catch(() => '');
            console.log(`  Dismissing: "${label || text.trim().slice(0, 40)}" (${sel})`);
            await btn.click();
            await page.waitForTimeout(500);
          }
        }
      } catch { /* ignore */ }
    }
  }
}

/**
 * Wait for any visible [role="dialog"] to disappear, with a timeout.
 * Useful after dismissing popups to ensure overlay is gone.
 */
async function waitForDialogsToClear(page: Page, timeout = 10_000) {
  const dialog = page.locator('[role="dialog"]:visible');
  await dialog.waitFor({ state: 'hidden', timeout }).catch(() => {
    console.log('  Some dialogs still visible after timeout — continuing.');
  });
}

// ============================================================
// TEST
// ============================================================

test('D365 Copilot prompt regression test', async ({ page, context }) => {
  // Register dialog handler once to avoid accumulating listeners
  page.on('dialog', async (dialog) => {
    console.log(`  Dismissing dialog: ${dialog.message().slice(0, 60)}`);
    await dialog.dismiss();
  });

  const prompts = await readPrompts();
  console.log(`\nLoaded ${prompts.length} prompts from: ${INPUT_XLSX}\n`);

  // Launch D365 via the Power Platform Playwright Toolkit
  console.log(`Navigating to: ${D365_URL}\n`);
  const mda = await launchD365App(page, context);

  // Open and wait for the Copilot side panel
  await openCopilotPanel(page);
  await waitForCopilotReady(page);

  // Dismiss the "A Copilot for you!" onboarding dialog and any others
  console.log('  Checking for popups...');
  await dismissPopups(page);
  await waitForDialogsToClear(page);
  // Second pass — some popups appear after the first is dismissed
  await page.waitForTimeout(2000);
  await dismissPopups(page);

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

    // Clear the conversation if the 15-message limit was reached
    await clearChatIfNeeded(page);

    const { prompt, expectedResponse, referencedDocs } = prompts[i];
    console.log(`[${i + 1}/${prompts.length}] Sending: ${prompt.slice(0, 80)}...`);

    let actualResponse = '';
    let sim = 0;
    let pass = false;
    let citedSources = '';

    try {
      actualResponse = await sendPromptAndGetResponse(page, prompt);

      // If we hit the chat limit, clear and retry this prompt
      if (actualResponse === 'CHAT_LIMIT_REACHED') {
        await clearChatIfNeeded(page);
        actualResponse = await sendPromptAndGetResponse(page, prompt);
      }

      // Extract the cited sources from the "Check sources" accordion
      citedSources = await extractCitedSources(page);
      if (citedSources) console.log(`  → Sources: ${citedSources}`);

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
      citedSources,
    });
  }

  // Write results Excel
  await writeResults(results);
  console.log(`\nResults written to: ${OUTPUT_XLSX}`);

  // Summary
  const passCount = results.filter((r) => r.pass).length;
  console.log(`\nSummary: ${passCount}/${results.length} passed (threshold: ${SIMILARITY_THRESHOLD * 100}%)\n`);
});

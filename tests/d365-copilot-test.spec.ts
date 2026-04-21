import { test, expect, Page } from '@playwright/test';
import * as path from 'path';
import ExcelJS from 'exceljs';

// ============================================================
// CONFIGURATION — edit .env or set environment variables
// ============================================================
const D365_URL = process.env.D365_URL || 'https://REPLACE_WITH_YOUR_ORG.crm.dynamics.com';
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

async function openCopilotPanel(page: Page) {
  // Look for the Copilot button in the D365 command bar / side pane
  // Common selectors — update if your UI differs:
  const copilotButtonSelectors = [
    'button[aria-label*="Copilot"]',
    'button[title*="Copilot"]',
    '[data-id="msdyn_copilot"]',
    'button:has-text("Copilot")',
  ];

  for (const selector of copilotButtonSelectors) {
    const btn = page.locator(selector).first();
    if (await btn.isVisible({ timeout: 5000 }).catch(() => false)) {
      await btn.click();
      await page.waitForTimeout(2000);
      return;
    }
  }

  console.warn('Could not find Copilot button — panel may already be open.');
}

async function sendPromptAndGetResponse(page: Page, prompt: string): Promise<string> {
  // -------------------------------------------------------
  // COPILOT INPUT — the chat text box in the Copilot pane
  // Update these selectors to match your D365 environment.
  // -------------------------------------------------------
  const inputSelectors = [
    'textarea[data-id*="copilot"]',
    'textarea[aria-label*="Type your message"]',
    'textarea[aria-label*="Ask a question"]',
    'textarea[placeholder*="Ask"]',
    '[data-id="webchat-sendbox-input"]',
    'textarea[data-id="webchat-sendbox-input"]',
  ];

  let input: ReturnType<Page['locator']> | null = null;
  for (const sel of inputSelectors) {
    const loc = page.locator(sel).first();
    if (await loc.isVisible({ timeout: 3000 }).catch(() => false)) {
      input = loc;
      break;
    }
  }

  if (!input) {
    throw new Error(
      'Could not find the Copilot chat input. Please update the selectors in sendPromptAndGetResponse().'
    );
  }

  // Count existing messages before sending
  const messageContainerSelectors = [
    '[data-content="message-body"]',
    '.webchat__bubble__content',
    '[class*="message-content"]',
    '[role="listitem"]',
  ];

  let messageSelector = messageContainerSelectors[0];
  for (const sel of messageContainerSelectors) {
    if (await page.locator(sel).first().isVisible({ timeout: 2000 }).catch(() => false)) {
      messageSelector = sel;
      break;
    }
  }

  const existingCount = await page.locator(messageSelector).count();

  // Type and send the prompt
  await input.click();
  await input.fill(prompt);
  await page.keyboard.press('Enter');

  // Wait for a new bot response to appear
  await page.waitForFunction(
    ({ selector, prevCount }) => {
      const msgs = document.querySelectorAll(selector);
      return msgs.length > prevCount + 1; // +1 for our sent message, +1 for bot reply
    },
    { selector: messageSelector, prevCount: existingCount },
    { timeout: RESPONSE_TIMEOUT }
  );

  // Small delay for streaming to finish
  await page.waitForTimeout(3000);

  // Grab the last bot message
  const messages = page.locator(messageSelector);
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
  await page.goto(D365_URL, { waitUntil: 'domcontentloaded', timeout: 60_000 });
  await page.waitForTimeout(5000); // let D365 finish loading

  // Open Copilot panel
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

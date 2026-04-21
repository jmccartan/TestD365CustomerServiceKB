/**
 * Global setup for parallel tests.
 * Reads the Excel prompts and caches them as JSON so each parallel
 * worker can read them synchronously at test generation time.
 */

import * as fs from 'fs';
import * as path from 'path';
import ExcelJS from 'exceljs';
import 'dotenv/config';

const INPUT_XLSX = path.resolve(__dirname, 'Prompts and Responses.xlsx');
const CACHE_FILE = path.resolve(__dirname, 'test-results', 'prompts-cache.json');
const RESULTS_DIR = path.resolve(__dirname, 'test-results', 'parallel');

async function globalSetup() {
  // 1. Read prompts from Excel and cache as JSON
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(INPUT_XLSX);
  const ws = wb.getWorksheet('Prompts & Responses') || wb.worksheets[0];
  const rows: Array<{ prompt: string; expectedResponse: string; referencedDocs: string }> = [];

  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const prompt = row.getCell(1).text?.trim();
    const expectedResponse = row.getCell(2).text?.trim();
    const referencedDocs = row.getCell(3).text?.trim() || '';
    if (prompt) {
      rows.push({ prompt, expectedResponse, referencedDocs });
    }
  });

  fs.mkdirSync(path.dirname(CACHE_FILE), { recursive: true });
  fs.writeFileSync(CACHE_FILE, JSON.stringify(rows, null, 2));
  console.log(`Cached ${rows.length} prompts to ${CACHE_FILE}`);

  // 2. Clean up previous parallel results
  if (fs.existsSync(RESULTS_DIR)) {
    for (const file of fs.readdirSync(RESULTS_DIR)) {
      if (file.endsWith('.json')) {
        fs.unlinkSync(path.join(RESULTS_DIR, file));
      }
    }
  }
}

export default globalSetup;

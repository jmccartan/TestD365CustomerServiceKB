// Helper: reads prompts from Excel and writes a JSON cache file.
// Called by the parallel test spec if the cache doesn't exist.
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const INPUT_XLSX = path.resolve(__dirname, 'Prompts and Responses.xlsx');
const CACHE_FILE = path.resolve(__dirname, 'test-results', 'prompts-cache.json');

(async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(INPUT_XLSX);
  const ws = wb.getWorksheet('Prompts & Responses') || wb.worksheets[0];
  const rows = [];

  ws.eachRow((row, n) => {
    if (n === 1) return;
    const prompt = row.getCell(1).text?.trim();
    const expectedResponse = row.getCell(2).text?.trim();
    const referencedDocs = row.getCell(3).text?.trim() || '';
    if (prompt) rows.push({ prompt, expectedResponse, referencedDocs });
  });

  fs.mkdirSync(path.dirname(CACHE_FILE), { recursive: true });
  fs.writeFileSync(CACHE_FILE, JSON.stringify(rows, null, 2));
  console.log(`Cached ${rows.length} prompts to ${CACHE_FILE}`);
})();

/**
 * Global teardown for parallel tests.
 * Merges per-test JSON result files into a single Excel report.
 */

import * as fs from 'fs';
import * as path from 'path';
import ExcelJS from 'exceljs';

const RESULTS_DIR = path.resolve(__dirname, 'test-results', 'parallel');

interface TestResult {
  index: number;
  prompt: string;
  expectedResponse: string;
  actualResponse: string;
  similarity: number;
  pass: boolean;
  referencedDocs: string;
}

async function globalTeardown() {
  if (!fs.existsSync(RESULTS_DIR)) {
    console.log('No parallel results to merge.');
    return;
  }

  const files = fs.readdirSync(RESULTS_DIR).filter((f) => f.endsWith('.json')).sort();
  if (files.length === 0) {
    console.log('No result files found.');
    return;
  }

  const results: TestResult[] = files.map((f) =>
    JSON.parse(fs.readFileSync(path.join(RESULTS_DIR, f), 'utf-8'))
  );

  // Sort by index
  results.sort((a, b) => a.index - b.index);

  // Build Excel
  const now = new Date();
  const timestamp = `${now.toISOString().slice(0, 10)}_${now.toTimeString().slice(0, 8).replace(/:/g, '-')}`;
  const outputFile = path.resolve(__dirname, `Test Results ${timestamp}.xlsx`);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Test Results');

  ws.columns = [
    { header: '#', key: 'index', width: 5 },
    { header: 'Prompt', key: 'prompt', width: 60 },
    { header: 'Expected Response', key: 'expected', width: 60 },
    { header: 'Actual Response', key: 'actual', width: 60 },
    { header: 'Similarity', key: 'similarity', width: 12 },
    { header: 'Result', key: 'result', width: 10 },
    { header: 'Referenced Docs', key: 'docs', width: 30 },
  ];

  ws.getRow(1).font = { bold: true };
  ws.getRow(1).alignment = { vertical: 'middle' };

  const SIMILARITY_THRESHOLD = parseFloat(process.env.SIMILARITY_THRESHOLD || '0.6');

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

    const resultCell = row.getCell('result');
    resultCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: r.pass ? 'FF92D050' : 'FFFF4444' },
    };
    resultCell.font = { bold: true, color: { argb: r.pass ? 'FF006100' : 'FF9C0006' } };
  }

  const passCount = results.filter((r) => r.pass).length;
  ws.addRow({});
  ws.addRow({
    index: '',
    prompt: `Total: ${results.length}  |  Pass: ${passCount}  |  Fail: ${results.length - passCount}`,
  });

  await wb.xlsx.writeFile(outputFile);
  console.log(`\n✅ Parallel results merged: ${outputFile}`);
  console.log(`   Summary: ${passCount}/${results.length} passed (threshold: ${SIMILARITY_THRESHOLD * 100}%)\n`);
}

export default globalTeardown;

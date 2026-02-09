/**
 * Excel Library Benchmark: SheetKit vs ExcelJS vs SheetJS (xlsx)
 *
 * Benchmarks:
 * -- READ --
 * 1. Read large file (50k rows)
 * 2. Read heavy-styles file
 * 3. Read multi-sheet file
 * 4. Read formulas file
 * 5. Read strings file
 * 6. Read data-validation file
 * 7. Read comments file
 * 8. Read merged-cells file
 * 9. Read mixed-workload (ERP document)
 *
 * -- SCALING (read) --
 * 10. Read 1k rows
 * 11. Read 10k rows
 * 12. Read 100k rows
 *
 * -- WRITE --
 * 13. Write large dataset (50k rows x 20 cols)
 * 14. Write with styles (5k rows, formatted)
 * 15. Write multi-sheet (10 sheets x 5k rows)
 * 16. Write formulas (10k rows)
 * 17. Write strings (20k rows text-heavy)
 * 18. Write data validation (5k rows, 8 rules)
 * 19. Write comments (2k rows)
 * 20. Write merged cells (500 merged regions)
 *
 * -- SCALING (write) --
 * 21. Write 1k rows x 10 cols
 * 22. Write 10k rows x 10 cols
 * 23. Write 50k rows x 10 cols
 * 24. Write 100k rows x 10 cols
 *
 * -- OTHER --
 * 25. Buffer round-trip (write to buffer, read back)
 * 26. Streaming write (50k rows) -- SheetKit and ExcelJS only
 * 27. Cell random-access read (1000 lookups on 50k-row file)
 * 28. Mixed workload write (ERP-style)
 */

import { Workbook as SheetKitWorkbook, JsStreamWriter } from '@sheetkit/node';
import ExcelJS from 'exceljs';
import XLSX from 'xlsx';
import { readFileSync, writeFileSync, existsSync, unlinkSync, statSync } from 'node:fs';
import { join } from 'node:path';

const FIXTURES_DIR = join(import.meta.dirname, 'fixtures');
const OUTPUT_DIR = join(import.meta.dirname, 'output');

// Match the Rust benchmark configuration
const WARMUP_RUNS = 1;
const BENCH_RUNS = 5;

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

interface BenchResult {
  library: string;
  scenario: string;
  category: string;
  timeMs: number;
  memoryMb: number;
  fileSizeKb?: number;
  timesMs: number[];
  memoryDeltas: number[];
}

const results: BenchResult[] = [];

function median(arr: number[]): number {
  if (arr.length === 0) return 0;
  const sorted = [...arr].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0
    ? (sorted[mid - 1] + sorted[mid]) / 2
    : sorted[mid];
}

function p95(arr: number[]): number {
  if (arr.length === 0) return 0;
  const sorted = [...arr].sort((a, b) => a - b);
  const idx = Math.ceil(0.95 * sorted.length) - 1;
  return sorted[Math.max(0, idx)];
}

function minVal(arr: number[]): number {
  if (arr.length === 0) return 0;
  return Math.min(...arr);
}

function maxVal(arr: number[]): number {
  if (arr.length === 0) return 0;
  return Math.max(...arr);
}

function formatMs(ms: number): string {
  if (ms < 1000) return `${ms.toFixed(0)}ms`;
  return `${(ms / 1000).toFixed(2)}s`;
}

function getMemoryMb(): number {
  if (global.gc) global.gc();
  return process.memoryUsage().heapUsed / 1024 / 1024;
}

function fileSizeKb(path: string): number | undefined {
  try {
    return statSync(path).size / 1024;
  } catch {
    return undefined;
  }
}

async function bench(
  library: string,
  scenario: string,
  category: string,
  fn: () => void | Promise<void>,
  outputPath?: string,
): Promise<BenchResult> {
  if (global.gc) global.gc();

  const memBefore = getMemoryMb();
  const start = performance.now();
  await fn();
  const elapsed = performance.now() - start;
  const memAfter = getMemoryMb();
  const memDelta = Math.max(0, memAfter - memBefore);
  const size = outputPath ? fileSizeKb(outputPath) : undefined;

  const result: BenchResult = {
    library,
    scenario,
    category,
    timeMs: elapsed,
    memoryMb: memDelta,
    fileSizeKb: size,
    timesMs: [elapsed],
    memoryDeltas: [memDelta],
  };
  results.push(result);

  const sizeStr = size != null ? ` | ${(size / 1024).toFixed(1)}MB` : '';
  console.log(
    `  [${library.padEnd(8)}] ${scenario.padEnd(46)} ${formatMs(elapsed).padStart(10)} | ${memDelta.toFixed(1).padStart(6)}MB${sizeStr}`,
  );
  return result;
}

async function benchMultiRun(
  library: string,
  scenario: string,
  category: string,
  fn: () => void | Promise<void>,
  outputPath?: string,
): Promise<BenchResult> {
  // Warmup runs (not measured)
  for (let i = 0; i < WARMUP_RUNS; i++) {
    if (global.gc) global.gc();
    await fn();
  }

  // Measured runs
  const timesMs: number[] = [];
  const memoryDeltas: number[] = [];
  let size: number | undefined;

  for (let i = 0; i < BENCH_RUNS; i++) {
    if (global.gc) global.gc();

    const memBefore = getMemoryMb();
    const start = performance.now();
    await fn();
    const elapsed = performance.now() - start;
    const memAfter = getMemoryMb();
    const memDelta = Math.max(0, memAfter - memBefore);

    timesMs.push(elapsed);
    memoryDeltas.push(memDelta);

    if (outputPath && i === BENCH_RUNS - 1) {
      size = fileSizeKb(outputPath);
    }
  }

  const med = median(timesMs);
  const min = minVal(timesMs);
  const max = maxVal(timesMs);
  const p95Val = p95(timesMs);
  const memMed = median(memoryDeltas);

  const result: BenchResult = {
    library,
    scenario,
    category,
    timeMs: med,
    memoryMb: memMed,
    fileSizeKb: size,
    timesMs,
    memoryDeltas,
  };
  results.push(result);

  const sizeStr = size != null ? ` | ${(size / 1024).toFixed(1)}MB` : '';
  console.log(
    `  [${library.padEnd(8)}] ${scenario.padEnd(46)} ` +
    `med=${formatMs(med).padStart(8)} ` +
    `min=${formatMs(min).padStart(8)} ` +
    `max=${formatMs(max).padStart(8)} ` +
    `p95=${formatMs(p95Val).padStart(8)} ` +
    `| ${memMed.toFixed(1).padStart(6)}MB${sizeStr}` +
    ` (${BENCH_RUNS} runs)`,
  );
  return result;
}

function cleanup(path: string) {
  try {
    if (existsSync(path)) unlinkSync(path);
  } catch { /* ignore */ }
}

function colLetter(n: number): string {
  let s = '';
  while (n >= 0) {
    s = String.fromCharCode((n % 26) + 65) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

// ---------------------------------------------------------------------------
// READ benchmarks
// ---------------------------------------------------------------------------

async function benchReadFile(filename: string, label: string, category: string) {
  const filepath = join(FIXTURES_DIR, filename);
  if (!existsSync(filepath)) {
    console.log(`  SKIP: ${filepath} not found. Run 'pnpm generate' first.`);
    return;
  }

  console.log(`\n--- ${label} ---`);

  await benchMultiRun('SheetKit', `Read ${label}`, category, () => {
    const wb = SheetKitWorkbook.openSync(filepath);
    for (const name of wb.sheetNames) {
      wb.getRows(name);
    }
  });

  await bench('ExcelJS', `Read ${label}`, category, async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filepath);
    wb.eachSheet((ws) => {
      ws.eachRow(() => { /* iterate */ });
    });
  });

  await bench('SheetJS', `Read ${label}`, category, () => {
    const buf = readFileSync(filepath);
    const wb = XLSX.read(buf, { type: 'buffer' });
    for (const name of wb.SheetNames) {
      XLSX.utils.sheet_to_json(wb.Sheets[name]);
    }
  });
}

// ---------------------------------------------------------------------------
// WRITE benchmarks
// ---------------------------------------------------------------------------

async function benchWriteLargeData() {
  const ROWS = 50_000;
  const COLS = 20;
  const label = `Write ${ROWS} rows x ${COLS} cols`;
  console.log(`\n--- ${label} ---`);

  const outSk = join(OUTPUT_DIR, 'write-large-sheetkit.xlsx');
  const outEj = join(OUTPUT_DIR, 'write-large-exceljs.xlsx');
  const outSj = join(OUTPUT_DIR, 'write-large-sheetjs.xlsx');

  await benchMultiRun('SheetKit', label, 'Write', () => {
    const wb = new SheetKitWorkbook();
    const sheet = 'Sheet1';
    const data: (string | number | boolean | null)[][] = [];
    for (let r = 1; r <= ROWS; r++) {
      const row: (string | number | boolean | null)[] = [];
      for (let c = 0; c < COLS; c++) {
        if (c % 3 === 0) row.push(r * (c + 1));
        else if (c % 3 === 1) row.push(`R${r}C${c}`);
        else row.push((r * c) / 100);
      }
      data.push(row);
    }
    wb.setSheetData(sheet, data);
    wb.saveSync(outSk);
  }, outSk);

  await bench('ExcelJS', label, 'Write', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    for (let r = 1; r <= ROWS; r++) {
      const row: (number | string)[] = [];
      for (let c = 0; c < COLS; c++) {
        if (c % 3 === 0) row.push(r * (c + 1));
        else if (c % 3 === 1) row.push(`R${r}C${c}`);
        else row.push((r * c) / 100);
      }
      ws.addRow(row);
    }
    await wb.xlsx.writeFile(outEj);
  }, outEj);

  await bench('SheetJS', label, 'Write', () => {
    const data: (number | string)[][] = [];
    for (let r = 0; r < ROWS; r++) {
      const row: (number | string)[] = [];
      for (let c = 0; c < COLS; c++) {
        if (c % 3 === 0) row.push((r + 1) * (c + 1));
        else if (c % 3 === 1) row.push(`R${r + 1}C${c}`);
        else row.push(((r + 1) * c) / 100);
      }
      data.push(row);
    }
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, outSj);
  }, outSj);

  cleanup(outSk); cleanup(outEj); cleanup(outSj);
}

async function benchWriteWithStyles() {
  const ROWS = 5_000;
  const label = `Write ${ROWS} styled rows`;
  console.log(`\n--- ${label} ---`);

  const outSk = join(OUTPUT_DIR, 'write-styles-sheetkit.xlsx');
  const outEj = join(OUTPUT_DIR, 'write-styles-exceljs.xlsx');
  const outSj = join(OUTPUT_DIR, 'write-styles-sheetjs.xlsx');

  await benchMultiRun('SheetKit', label, 'Write', () => {
    const wb = new SheetKitWorkbook();
    const sheet = 'Sheet1';
    const boldId = wb.addStyle({
      font: { bold: true, size: 12, name: 'Arial', color: 'FFFFFFFF' },
      fill: { pattern: 'solid', fgColor: 'FF4472C4' },
      border: {
        top: { style: 'thin', color: 'FF000000' },
        bottom: { style: 'thin', color: 'FF000000' },
        left: { style: 'thin', color: 'FF000000' },
        right: { style: 'thin', color: 'FF000000' },
      },
      alignment: { horizontal: 'center' },
    });
    const numId = wb.addStyle({ numFmtId: 4, font: { name: 'Calibri', size: 11 } });
    const pctId = wb.addStyle({ numFmtId: 10, font: { italic: true } });

    wb.setRowValues(sheet, 1, 'A', Array.from({ length: 10 }, (_, c) => `Header${c + 1}`));
    for (let c = 0; c < 10; c++) {
      wb.setCellStyle(sheet, `${colLetter(c)}1`, boldId);
    }

    const data: (string | number | boolean | null)[][] = [];
    for (let r = 2; r <= ROWS + 1; r++) {
      const row: (string | number | boolean | null)[] = [];
      for (let c = 0; c < 10; c++) {
        if (c % 3 === 0) row.push(r * c);
        else if (c % 3 === 1) row.push(`Data_${r}_${c}`);
        else row.push((r % 100) / 100);
      }
      data.push(row);
    }
    wb.setSheetData(sheet, data, 'A2');

    for (let r = 2; r <= ROWS + 1; r++) {
      for (let c = 0; c < 10; c++) {
        if (c % 3 === 0) wb.setCellStyle(sheet, `${colLetter(c)}${r}`, numId);
        else if (c % 3 === 2) wb.setCellStyle(sheet, `${colLetter(c)}${r}`, pctId);
      }
    }
    wb.saveSync(outSk);
  }, outSk);

  await bench('ExcelJS', label, 'Write', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    const headerRow = ws.addRow(Array.from({ length: 10 }, (_, c) => `Header${c + 1}`));
    headerRow.eachCell((cell) => {
      cell.font = { bold: true, size: 12, name: 'Arial', color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
      cell.border = {
        top: { style: 'thin' }, bottom: { style: 'thin' },
        left: { style: 'thin' }, right: { style: 'thin' },
      };
      cell.alignment = { horizontal: 'center' };
    });

    for (let r = 0; r < ROWS; r++) {
      const rowData: (number | string)[] = [];
      for (let c = 0; c < 10; c++) {
        if (c % 3 === 0) rowData.push((r + 2) * c);
        else if (c % 3 === 1) rowData.push(`Data_${r + 2}_${c}`);
        else rowData.push(((r + 2) % 100) / 100);
      }
      const row = ws.addRow(rowData);
      row.eachCell((cell, colNumber) => {
        const c = colNumber - 1;
        if (c % 3 === 0) {
          cell.numFmt = '#,##0.00';
          cell.font = { name: 'Calibri', size: 11 };
        } else if (c % 3 === 2) {
          cell.numFmt = '0.00%';
          cell.font = { italic: true };
        }
      });
    }
    await wb.xlsx.writeFile(outEj);
  }, outEj);

  await bench('SheetJS', label, 'Write', () => {
    const data: (number | string)[][] = [
      Array.from({ length: 10 }, (_, c) => `Header${c + 1}`),
    ];
    for (let r = 0; r < ROWS; r++) {
      const row: (number | string)[] = [];
      for (let c = 0; c < 10; c++) {
        if (c % 3 === 0) row.push((r + 2) * c);
        else if (c % 3 === 1) row.push(`Data_${r + 2}_${c}`);
        else row.push(((r + 2) % 100) / 100);
      }
      data.push(row);
    }
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, outSj);
  }, outSj);

  cleanup(outSk); cleanup(outEj); cleanup(outSj);
}

async function benchWriteMultiSheet() {
  const SHEETS = 10;
  const ROWS = 5_000;
  const COLS = 10;
  const label = `Write ${SHEETS} sheets x ${ROWS} rows`;
  console.log(`\n--- ${label} ---`);

  const outSk = join(OUTPUT_DIR, 'write-multi-sheetkit.xlsx');
  const outEj = join(OUTPUT_DIR, 'write-multi-exceljs.xlsx');
  const outSj = join(OUTPUT_DIR, 'write-multi-sheetjs.xlsx');

  await benchMultiRun('SheetKit', label, 'Write', () => {
    const wb = new SheetKitWorkbook();
    for (let s = 0; s < SHEETS; s++) {
      const name = s === 0 ? 'Sheet1' : `Sheet${s + 1}`;
      if (s > 0) wb.newSheet(name);
      const data: (string | number | boolean | null)[][] = [];
      for (let r = 1; r <= ROWS; r++) {
        const row: (string | number | boolean | null)[] = [];
        for (let c = 0; c < COLS; c++) {
          if (c % 2 === 0) row.push(r * (c + 1));
          else row.push(`S${s}R${r}C${c}`);
        }
        data.push(row);
      }
      wb.setSheetData(name, data);
    }
    wb.saveSync(outSk);
  }, outSk);

  await bench('ExcelJS', label, 'Write', async () => {
    const wb = new ExcelJS.Workbook();
    for (let s = 0; s < SHEETS; s++) {
      const ws = wb.addWorksheet(`Sheet${s + 1}`);
      for (let r = 0; r < ROWS; r++) {
        const row: (number | string)[] = [];
        for (let c = 0; c < COLS; c++) {
          if (c % 2 === 0) row.push((r + 1) * (c + 1));
          else row.push(`S${s}R${r + 1}C${c}`);
        }
        ws.addRow(row);
      }
    }
    await wb.xlsx.writeFile(outEj);
  }, outEj);

  await bench('SheetJS', label, 'Write', () => {
    const wb = XLSX.utils.book_new();
    for (let s = 0; s < SHEETS; s++) {
      const data: (number | string)[][] = [];
      for (let r = 0; r < ROWS; r++) {
        const row: (number | string)[] = [];
        for (let c = 0; c < COLS; c++) {
          if (c % 2 === 0) row.push((r + 1) * (c + 1));
          else row.push(`S${s}R${r + 1}C${c}`);
        }
        data.push(row);
      }
      const ws = XLSX.utils.aoa_to_sheet(data);
      XLSX.utils.book_append_sheet(wb, ws, `Sheet${s + 1}`);
    }
    XLSX.writeFile(wb, outSj);
  }, outSj);

  cleanup(outSk); cleanup(outEj); cleanup(outSj);
}

async function benchWriteFormulas() {
  const ROWS = 10_000;
  const label = `Write ${ROWS} rows with formulas`;
  console.log(`\n--- ${label} ---`);

  const outSk = join(OUTPUT_DIR, 'write-formulas-sheetkit.xlsx');
  const outEj = join(OUTPUT_DIR, 'write-formulas-exceljs.xlsx');
  const outSj = join(OUTPUT_DIR, 'write-formulas-sheetjs.xlsx');

  await benchMultiRun('SheetKit', label, 'Write', () => {
    const wb = new SheetKitWorkbook();
    const sheet = 'Sheet1';
    const data: (string | number | boolean | null)[][] = [];
    for (let r = 1; r <= ROWS; r++) {
      data.push([r * 1.5, (r % 100) + 0.5]);
    }
    wb.setSheetData(sheet, data);
    for (let r = 1; r <= ROWS; r++) {
      wb.setCellFormula(sheet, `C${r}`, `A${r}+B${r}`);
      wb.setCellFormula(sheet, `D${r}`, `A${r}*B${r}`);
      wb.setCellFormula(sheet, `E${r}`, `IF(A${r}>B${r},"A","B")`);
    }
    wb.saveSync(outSk);
  }, outSk);

  await bench('ExcelJS', label, 'Write', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    for (let r = 1; r <= ROWS; r++) {
      ws.getCell(`A${r}`).value = r * 1.5;
      ws.getCell(`B${r}`).value = (r % 100) + 0.5;
      ws.getCell(`C${r}`).value = { formula: `A${r}+B${r}` } as any;
      ws.getCell(`D${r}`).value = { formula: `A${r}*B${r}` } as any;
      ws.getCell(`E${r}`).value = { formula: `IF(A${r}>B${r},"A","B")` } as any;
    }
    await wb.xlsx.writeFile(outEj);
  }, outEj);

  await bench('SheetJS', label, 'Write', () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([]);
    for (let r = 0; r < ROWS; r++) {
      XLSX.utils.sheet_add_aoa(ws, [[r * 1.5 + 1.5, ((r + 1) % 100) + 0.5]], { origin: `A${r + 1}` });
      ws[`C${r + 1}`] = { t: 's', f: `A${r + 1}+B${r + 1}` };
      ws[`D${r + 1}`] = { t: 's', f: `A${r + 1}*B${r + 1}` };
      ws[`E${r + 1}`] = { t: 's', f: `IF(A${r + 1}>B${r + 1},"A","B")` };
    }
    ws['!ref'] = `A1:E${ROWS}`;
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, outSj);
  }, outSj);

  cleanup(outSk); cleanup(outEj); cleanup(outSj);
}

async function benchWriteStrings() {
  const ROWS = 20_000;
  const COLS = 10;
  const label = `Write ${ROWS} text-heavy rows`;
  console.log(`\n--- ${label} ---`);

  const outSk = join(OUTPUT_DIR, 'write-strings-sheetkit.xlsx');
  const outEj = join(OUTPUT_DIR, 'write-strings-exceljs.xlsx');
  const outSj = join(OUTPUT_DIR, 'write-strings-sheetjs.xlsx');

  const words = ['alpha', 'bravo', 'charlie', 'delta', 'echo',
    'foxtrot', 'golf', 'hotel', 'india', 'juliet'];

  function makeRow(r: number): string[] {
    const w1 = words[(r * 3) % words.length];
    const w2 = words[(r * 7) % words.length];
    return [
      `${w1} ${w2}`, `${w1}.${w2}@example.com`, `Dept_${r % 20}`,
      `${r} ${w1} Street`, `Description for ${r}: ${w1} ${w2}`,
      `City_${w1}`, `Country_${w2}`, `+1-555-${String(r % 10000).padStart(4, '0')}`,
      `${w1} Specialist`, `Experienced ${w2} professional in ${w1} domain.`,
    ];
  }

  await benchMultiRun('SheetKit', label, 'Write', () => {
    const wb = new SheetKitWorkbook();
    const sheet = 'Sheet1';
    const data: (string | number | boolean | null)[][] = [];
    for (let r = 1; r <= ROWS; r++) {
      data.push(makeRow(r));
    }
    wb.setSheetData(sheet, data);
    wb.saveSync(outSk);
  }, outSk);

  await bench('ExcelJS', label, 'Write', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    for (let r = 1; r <= ROWS; r++) {
      ws.addRow(makeRow(r));
    }
    await wb.xlsx.writeFile(outEj);
  }, outEj);

  await bench('SheetJS', label, 'Write', () => {
    const data: string[][] = [];
    for (let r = 1; r <= ROWS; r++) {
      data.push(makeRow(r));
    }
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, outSj);
  }, outSj);

  cleanup(outSk); cleanup(outEj); cleanup(outSj);
}

async function benchWriteDataValidation() {
  const ROWS = 5_000;
  const label = `Write ${ROWS} rows + 8 validation rules`;
  console.log(`\n--- ${label} ---`);

  const outSk = join(OUTPUT_DIR, 'write-dv-sheetkit.xlsx');
  const outEj = join(OUTPUT_DIR, 'write-dv-exceljs.xlsx');

  // SheetKit
  await benchMultiRun('SheetKit', label, 'Write (DV)', () => {
    const wb = new SheetKitWorkbook();
    const sheet = 'Sheet1';

    wb.addDataValidation(sheet, {
      sqref: `A2:A${ROWS + 1}`, validationType: 'list',
      formula1: '"Active,Inactive,Pending,Closed"',
    });
    wb.addDataValidation(sheet, {
      sqref: `B2:B${ROWS + 1}`, validationType: 'whole',
      operator: 'between', formula1: '0', formula2: '100',
    });
    wb.addDataValidation(sheet, {
      sqref: `C2:C${ROWS + 1}`, validationType: 'decimal',
      operator: 'between', formula1: '0', formula2: '1',
    });
    wb.addDataValidation(sheet, {
      sqref: `D2:D${ROWS + 1}`, validationType: 'textLength',
      operator: 'lessThanOrEqual', formula1: '50',
    });

    const statuses = ['Active', 'Inactive', 'Pending', 'Closed'];
    const data: (string | number | boolean | null)[][] = [];
    for (let r = 2; r <= ROWS + 1; r++) {
      data.push([statuses[r % 4], r % 101, (r % 100) / 100, `Item_${r}`]);
    }
    wb.setSheetData(sheet, data, 'A2');
    wb.saveSync(outSk);
  }, outSk);

  // ExcelJS
  await bench('ExcelJS', label, 'Write (DV)', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');

    const statuses = ['Active', 'Inactive', 'Pending', 'Closed'];
    for (let r = 2; r <= ROWS + 1; r++) {
      ws.getCell(`A${r}`).value = statuses[r % 4];
      ws.getCell(`A${r}`).dataValidation = {
        type: 'list', formulae: ['"Active,Inactive,Pending,Closed"'],
        allowBlank: true, showDropDown: true,
      };
      ws.getCell(`B${r}`).value = r % 101;
      ws.getCell(`B${r}`).dataValidation = {
        type: 'whole', operator: 'between', formulae: [0, 100],
      };
      ws.getCell(`C${r}`).value = (r % 100) / 100;
      ws.getCell(`C${r}`).dataValidation = {
        type: 'decimal', operator: 'between', formulae: [0, 1],
      };
      ws.getCell(`D${r}`).value = `Item_${r}`;
      ws.getCell(`D${r}`).dataValidation = {
        type: 'textLength', operator: 'lessThanOrEqual', formulae: [50],
      };
    }
    await wb.xlsx.writeFile(outEj);
  }, outEj);

  // SheetJS -- no data validation support
  console.log('  [SheetJS  ] N/A (no data validation in community edition)');

  cleanup(outSk); cleanup(outEj);
}

async function benchWriteComments() {
  const ROWS = 2_000;
  const label = `Write ${ROWS} rows with comments`;
  console.log(`\n--- ${label} ---`);

  const outSk = join(OUTPUT_DIR, 'write-comments-sheetkit.xlsx');
  const outEj = join(OUTPUT_DIR, 'write-comments-exceljs.xlsx');
  const outSj = join(OUTPUT_DIR, 'write-comments-sheetjs.xlsx');

  await benchMultiRun('SheetKit', label, 'Write (Comments)', () => {
    const wb = new SheetKitWorkbook();
    const sheet = 'Sheet1';
    const data: (string | number | boolean | null)[][] = [];
    for (let r = 1; r <= ROWS; r++) {
      data.push([r, `Name_${r}`]);
    }
    wb.setSheetData(sheet, data);
    for (let r = 1; r <= ROWS; r++) {
      wb.addComment(sheet, {
        cell: `A${r}`, author: 'Reviewer',
        text: `Comment for row ${r}: review completed.`,
      });
    }
    wb.saveSync(outSk);
  }, outSk);

  await bench('ExcelJS', label, 'Write (Comments)', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    for (let r = 1; r <= ROWS; r++) {
      ws.getCell(`A${r}`).value = r;
      ws.getCell(`B${r}`).value = `Name_${r}`;
      ws.getCell(`A${r}`).note = `Comment for row ${r}: review completed.`;
    }
    await wb.xlsx.writeFile(outEj);
  }, outEj);

  await bench('SheetJS', label, 'Write (Comments)', () => {
    const data: (number | string)[][] = [];
    for (let r = 0; r < ROWS; r++) {
      data.push([r + 1, `Name_${r + 1}`]);
    }
    const ws = XLSX.utils.aoa_to_sheet(data);
    for (let r = 0; r < ROWS; r++) {
      const cellAddr = `A${r + 1}`;
      if (!ws[cellAddr]) ws[cellAddr] = { t: 'n', v: r + 1 };
      (ws[cellAddr] as any).c = [{ a: 'Reviewer', t: `Comment for row ${r + 1}: review completed.` }];
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, outSj);
  }, outSj);

  cleanup(outSk); cleanup(outEj); cleanup(outSj);
}

async function benchWriteMergedCells() {
  const REGIONS = 500;
  const label = `Write ${REGIONS} merged regions`;
  console.log(`\n--- ${label} ---`);

  const outSk = join(OUTPUT_DIR, 'write-merge-sheetkit.xlsx');
  const outEj = join(OUTPUT_DIR, 'write-merge-exceljs.xlsx');
  const outSj = join(OUTPUT_DIR, 'write-merge-sheetjs.xlsx');

  await benchMultiRun('SheetKit', label, 'Write (Merge)', () => {
    const wb = new SheetKitWorkbook();
    const sheet = 'Sheet1';
    const cells: { cell: string; value: string | number | boolean | null }[] = [];
    for (let i = 0; i < REGIONS; i++) {
      const row = i * 3 + 1;
      cells.push({ cell: `A${row}`, value: `Section ${i + 1}` });
      cells.push({ cell: `A${row + 1}`, value: i * 100 });
      cells.push({ cell: `B${row + 1}`, value: `Data_${i}` });
    }
    wb.setCellValues(sheet, cells);
    for (let i = 0; i < REGIONS; i++) {
      const row = i * 3 + 1;
      wb.mergeCells(sheet, `A${row}`, `D${row}`);
    }
    wb.saveSync(outSk);
  }, outSk);

  await bench('ExcelJS', label, 'Write (Merge)', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    for (let i = 0; i < REGIONS; i++) {
      const row = i * 3 + 1;
      ws.mergeCells(row, 1, row, 4);
      ws.getCell(row, 1).value = `Section ${i + 1}`;
      ws.getCell(row + 1, 1).value = i * 100;
      ws.getCell(row + 1, 2).value = `Data_${i}`;
    }
    await wb.xlsx.writeFile(outEj);
  }, outEj);

  await bench('SheetJS', label, 'Write (Merge)', () => {
    const ws: XLSX.WorkSheet = {};
    const merges: XLSX.Range[] = [];
    for (let i = 0; i < REGIONS; i++) {
      const row = i * 3;
      ws[XLSX.utils.encode_cell({ r: row, c: 0 })] = { t: 's', v: `Section ${i + 1}` };
      merges.push({ s: { r: row, c: 0 }, e: { r: row, c: 3 } });
      ws[XLSX.utils.encode_cell({ r: row + 1, c: 0 })] = { t: 'n', v: i * 100 };
      ws[XLSX.utils.encode_cell({ r: row + 1, c: 1 })] = { t: 's', v: `Data_${i}` };
    }
    ws['!ref'] = `A1:D${REGIONS * 3}`;
    ws['!merges'] = merges;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, outSj);
  }, outSj);

  cleanup(outSk); cleanup(outEj); cleanup(outSj);
}

// ---------------------------------------------------------------------------
// Scaling write benchmarks
// ---------------------------------------------------------------------------

async function benchWriteScale(rows: number) {
  const COLS = 10;
  const label = `Write ${rows >= 1000 ? `${rows / 1000}k` : rows} rows x ${COLS} cols`;
  console.log(`\n--- ${label} ---`);

  const tag = rows >= 1000 ? `${rows / 1000}k` : `${rows}`;
  const outSk = join(OUTPUT_DIR, `write-scale-${tag}-sheetkit.xlsx`);
  const outEj = join(OUTPUT_DIR, `write-scale-${tag}-exceljs.xlsx`);
  const outSj = join(OUTPUT_DIR, `write-scale-${tag}-sheetjs.xlsx`);

  await benchMultiRun('SheetKit', label, 'Write (Scale)', () => {
    const wb = new SheetKitWorkbook();
    const sheet = 'Sheet1';
    const data: (string | number | boolean | null)[][] = [];
    for (let r = 1; r <= rows; r++) {
      const row: (string | number | boolean | null)[] = [];
      for (let c = 0; c < COLS; c++) {
        if (c % 3 === 0) row.push(r * (c + 1));
        else if (c % 3 === 1) row.push(`R${r}C${c}`);
        else row.push((r * c) / 100);
      }
      data.push(row);
    }
    wb.setSheetData(sheet, data);
    wb.saveSync(outSk);
  }, outSk);

  await bench('ExcelJS', label, 'Write (Scale)', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    for (let r = 0; r < rows; r++) {
      const row: (number | string)[] = [];
      for (let c = 0; c < COLS; c++) {
        if (c % 3 === 0) row.push((r + 1) * (c + 1));
        else if (c % 3 === 1) row.push(`R${r + 1}C${c}`);
        else row.push(((r + 1) * c) / 100);
      }
      ws.addRow(row);
    }
    await wb.xlsx.writeFile(outEj);
  }, outEj);

  await bench('SheetJS', label, 'Write (Scale)', () => {
    const data: (number | string)[][] = [];
    for (let r = 0; r < rows; r++) {
      const row: (number | string)[] = [];
      for (let c = 0; c < COLS; c++) {
        if (c % 3 === 0) row.push((r + 1) * (c + 1));
        else if (c % 3 === 1) row.push(`R${r + 1}C${c}`);
        else row.push(((r + 1) * c) / 100);
      }
      data.push(row);
    }
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, outSj);
  }, outSj);

  cleanup(outSk); cleanup(outEj); cleanup(outSj);
}

// ---------------------------------------------------------------------------
// Buffer round-trip
// ---------------------------------------------------------------------------

async function benchBufferRoundTrip() {
  const ROWS = 10_000;
  const COLS = 10;
  const label = `Buffer round-trip (${ROWS} rows)`;
  console.log(`\n--- ${label} ---`);

  await benchMultiRun('SheetKit', label, 'Round-Trip', () => {
    const wb = new SheetKitWorkbook();
    const sheet = 'Sheet1';
    const data: (string | number | boolean | null)[][] = [];
    for (let r = 1; r <= ROWS; r++) {
      const row: (string | number | boolean | null)[] = [];
      for (let c = 0; c < COLS; c++) {
        row.push(r * (c + 1));
      }
      data.push(row);
    }
    wb.setSheetData(sheet, data);
    const buf = wb.writeBufferSync();
    const wb2 = SheetKitWorkbook.openBufferSync(buf);
    wb2.getRows('Sheet1');
  });

  await bench('ExcelJS', label, 'Round-Trip', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    for (let r = 0; r < ROWS; r++) {
      const row: number[] = [];
      for (let c = 0; c < COLS; c++) {
        row.push((r + 1) * (c + 1));
      }
      ws.addRow(row);
    }
    const buf = await wb.xlsx.writeBuffer();
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buf as Buffer);
    wb2.getWorksheet('Sheet1')!.eachRow(() => { /* iterate */ });
  });

  await bench('SheetJS', label, 'Round-Trip', () => {
    const data: number[][] = [];
    for (let r = 0; r < ROWS; r++) {
      const row: number[] = [];
      for (let c = 0; c < COLS; c++) {
        row.push((r + 1) * (c + 1));
      }
      data.push(row);
    }
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    const wb2 = XLSX.read(buf, { type: 'buffer' });
    XLSX.utils.sheet_to_json(wb2.Sheets[wb2.SheetNames[0]]);
  });
}

// ---------------------------------------------------------------------------
// Streaming write
// ---------------------------------------------------------------------------

async function benchStreamingWrite() {
  const ROWS = 50_000;
  const COLS = 20;
  const label = `Streaming write (${ROWS} rows)`;
  console.log(`\n--- ${label} ---`);

  const outSk = join(OUTPUT_DIR, 'stream-sheetkit.xlsx');
  const outEj = join(OUTPUT_DIR, 'stream-exceljs.xlsx');

  await benchMultiRun('SheetKit', label, 'Streaming', () => {
    const wb = new SheetKitWorkbook();
    const sw = wb.newStreamWriter('StreamSheet');
    for (let r = 1; r <= ROWS; r++) {
      const vals: (string | number | boolean | null)[] = [];
      for (let c = 0; c < COLS; c++) {
        if (c % 3 === 0) vals.push(r * (c + 1));
        else if (c % 3 === 1) vals.push(`R${r}C${c}`);
        else vals.push((r * c) / 100);
      }
      sw.writeRow(r, vals);
    }
    wb.applyStreamWriter(sw);
    wb.saveSync(outSk);
  }, outSk);

  await bench('ExcelJS', label, 'Streaming', async () => {
    const options = {
      filename: outEj,
      useStyles: false,
      useSharedStrings: false,
    };
    const wb = new ExcelJS.stream.xlsx.WorkbookWriter(options);
    const ws = wb.addWorksheet('StreamSheet');
    for (let r = 1; r <= ROWS; r++) {
      const row: (number | string)[] = [];
      for (let c = 0; c < COLS; c++) {
        if (c % 3 === 0) row.push(r * (c + 1));
        else if (c % 3 === 1) row.push(`R${r}C${c}`);
        else row.push((r * c) / 100);
      }
      ws.addRow(row).commit();
    }
    await ws.commit();
    await wb.commit();
  }, outEj);

  console.log('  [SheetJS  ] N/A (no streaming API)');

  cleanup(outSk); cleanup(outEj);
}

// ---------------------------------------------------------------------------
// Cell random-access read
// ---------------------------------------------------------------------------

async function benchRandomAccessRead() {
  const filepath = join(FIXTURES_DIR, 'large-data.xlsx');
  if (!existsSync(filepath)) {
    console.log('  SKIP: large-data.xlsx not found. Run pnpm generate first.');
    return;
  }

  const LOOKUPS = 1_000;
  const label = `Random-access read (${LOOKUPS} cells from 50k-row file)`;
  console.log(`\n--- ${label} ---`);

  // Pre-generate random cell addresses for consistency
  const cells: string[] = [];
  for (let i = 0; i < LOOKUPS; i++) {
    const r = Math.floor(Math.random() * 50000) + 2;
    const c = Math.floor(Math.random() * 20);
    cells.push(`${colLetter(c)}${r}`);
  }

  await benchMultiRun('SheetKit', label, 'Random Access', () => {
    const wb = SheetKitWorkbook.openSync(filepath);
    for (const cell of cells) {
      wb.getCellValue('Sheet1', cell);
    }
  });

  await bench('ExcelJS', label, 'Random Access', async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filepath);
    const ws = wb.getWorksheet('Sheet1')!;
    for (const cell of cells) {
      ws.getCell(cell).value;
    }
  });

  await bench('SheetJS', label, 'Random Access', () => {
    const buf = readFileSync(filepath);
    const wb = XLSX.read(buf, { type: 'buffer' });
    const ws = wb.Sheets['Sheet1'];
    for (const cell of cells) {
      const val = ws[cell];
      if (val) val.v; // access value
    }
  });
}

// ---------------------------------------------------------------------------
// Mixed workload write
// ---------------------------------------------------------------------------

async function benchMixedWorkloadWrite() {
  const label = 'Mixed workload write (ERP-style)';
  console.log(`\n--- ${label} ---`);

  const outSk = join(OUTPUT_DIR, 'write-mixed-sheetkit.xlsx');
  const outEj = join(OUTPUT_DIR, 'write-mixed-exceljs.xlsx');

  await benchMultiRun('SheetKit', label, 'Mixed Write', () => {
    const wb = new SheetKitWorkbook();
    const sheet = 'Sheet1';

    const boldId = wb.addStyle({
      font: { bold: true, size: 11 },
      fill: { pattern: 'solid', fgColor: 'FFD9E2F3' },
    });
    const numId = wb.addStyle({ customNumFmt: '$#,##0.00' });

    // Title
    wb.setCellValue(sheet, 'A1', 'Sales Report Q4');
    wb.mergeCells(sheet, 'A1', 'F1');
    wb.setCellStyle(sheet, 'A1', boldId);

    // Headers
    const headers = ['Order_ID', 'Product', 'Quantity', 'Unit_Price', 'Total', 'Region'];
    wb.setRowValues(sheet, 2, 'A', headers);
    for (let c = 0; c < headers.length; c++) {
      wb.setCellStyle(sheet, `${colLetter(c)}2`, boldId);
    }

    // Validation
    wb.addDataValidation(sheet, {
      sqref: 'F3:F5002', validationType: 'list',
      formula1: '"North,South,East,West"',
    });
    wb.addDataValidation(sheet, {
      sqref: 'C3:C5002', validationType: 'whole',
      operator: 'greaterThan', formula1: '0',
    });

    const regions = ['North', 'South', 'East', 'West'];
    const products = ['Widget A', 'Widget B', 'Gadget X', 'Gadget Y', 'Service Z'];
    const data: (string | number | boolean | null)[][] = [];
    for (let i = 1; i <= 5000; i++) {
      data.push([
        `ORD-${String(i).padStart(5, '0')}`,
        products[i % products.length],
        (i % 50) + 1,
        ((i * 19) % 500) + 10,
        null,
        regions[i % regions.length],
      ]);
    }
    wb.setSheetData(sheet, data, 'A3');
    for (let r = 3; r <= 5002; r++) {
      const i = r - 2;
      wb.setCellStyle(sheet, `D${r}`, numId);
      wb.setCellFormula(sheet, `E${r}`, `C${r}*D${r}`);
      wb.setCellStyle(sheet, `E${r}`, numId);

      if (i % 50 === 0) {
        wb.addComment(sheet, {
          cell: `A${r}`, author: 'Sales',
          text: `Bulk order - special pricing applied for order ${i}.`,
        });
      }
    }

    wb.saveSync(outSk);
  }, outSk);

  await bench('ExcelJS', label, 'Mixed Write', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');

    ws.mergeCells('A1:F1');
    ws.getCell('A1').value = 'Sales Report Q4';
    ws.getCell('A1').font = { bold: true, size: 11 };
    ws.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E2F3' } };

    const headers = ['Order_ID', 'Product', 'Quantity', 'Unit_Price', 'Total', 'Region'];
    const headerRow = ws.addRow(headers);
    headerRow.eachCell((cell) => {
      cell.font = { bold: true, size: 11 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E2F3' } };
    });

    const regions = ['North', 'South', 'East', 'West'];
    const products = ['Widget A', 'Widget B', 'Gadget X', 'Gadget Y', 'Service Z'];
    for (let i = 1; i <= 5000; i++) {
      const r = i + 2;
      ws.getCell(`A${r}`).value = `ORD-${String(i).padStart(5, '0')}`;
      ws.getCell(`B${r}`).value = products[i % products.length];
      ws.getCell(`C${r}`).value = (i % 50) + 1;
      ws.getCell(`C${r}`).dataValidation = { type: 'whole', operator: 'greaterThan', formulae: [0] };
      ws.getCell(`D${r}`).value = ((i * 19) % 500) + 10;
      ws.getCell(`D${r}`).numFmt = '$#,##0.00';
      ws.getCell(`E${r}`).value = { formula: `C${r}*D${r}` } as any;
      ws.getCell(`E${r}`).numFmt = '$#,##0.00';
      ws.getCell(`F${r}`).value = regions[i % regions.length];
      ws.getCell(`F${r}`).dataValidation = {
        type: 'list', formulae: ['"North,South,East,West"'],
        allowBlank: true, showDropDown: true,
      };
      if (i % 50 === 0) {
        ws.getCell(`A${r}`).note = `Bulk order - special pricing applied for order ${i}.`;
      }
    }
    await wb.xlsx.writeFile(outEj);
  }, outEj);

  // SheetJS: limited -- no validation, styles, or comments in community edition
  console.log('  [SheetJS  ] N/A (no validation/styles/comments in community edition)');

  cleanup(outSk); cleanup(outEj);
}

// ---------------------------------------------------------------------------
// Results formatting
// ---------------------------------------------------------------------------

function formatStats(r: BenchResult): string {
  if (r.timesMs.length <= 1) {
    return formatMs(r.timeMs);
  }
  const med = median(r.timesMs);
  const min = minVal(r.timesMs);
  const max = maxVal(r.timesMs);
  return `${formatMs(med)} [${formatMs(min)}-${formatMs(max)}]`;
}

function printSummaryTable() {
  console.log('\n\n========================================');
  console.log(' BENCHMARK RESULTS SUMMARY');
  console.log('========================================');
  console.log(` SheetKit: ${WARMUP_RUNS} warmup + ${BENCH_RUNS} measured runs (median shown)`);
  console.log(' ExcelJS/SheetJS: single run\n');

  const scenarios = [...new Set(results.map((r) => r.scenario))];

  console.log(
    '| Scenario'.padEnd(51) +
    '| SheetKit (median)'.padEnd(24) +
    '| ExcelJS'.padEnd(14) +
    '| SheetJS'.padEnd(14) +
    '| Mem(SK)'.padEnd(11) +
    '| Winner |'
  );
  console.log(
    '|' + '-'.repeat(50) +
    '|' + '-'.repeat(23) +
    '|' + '-'.repeat(13) +
    '|' + '-'.repeat(13) +
    '|' + '-'.repeat(10) +
    '|--------|'
  );

  for (const scenario of scenarios) {
    const sk = results.find((r) => r.scenario === scenario && r.library === 'SheetKit');
    const ej = results.find((r) => r.scenario === scenario && r.library === 'ExcelJS');
    const sj = results.find((r) => r.scenario === scenario && r.library === 'SheetJS');

    const skTime = sk ? formatStats(sk) : 'N/A';
    const ejTime = ej ? formatMs(ej.timeMs) : 'N/A';
    const sjTime = sj ? formatMs(sj.timeMs) : 'N/A';
    const skMem = sk ? `${median(sk.memoryDeltas).toFixed(1)}MB` : 'N/A';

    const times = [
      { lib: 'SheetKit', ms: sk?.timeMs ?? Infinity },
      { lib: 'ExcelJS', ms: ej?.timeMs ?? Infinity },
      { lib: 'SheetJS', ms: sj?.timeMs ?? Infinity },
    ].filter((t) => t.ms < Infinity);

    const winner = times.length > 0
      ? times.reduce((a, b) => (a.ms < b.ms ? a : b)).lib
      : 'N/A';

    console.log(
      `| ${scenario.padEnd(49)}` +
      `| ${skTime.padEnd(22)}` +
      `| ${ejTime.padEnd(12)}` +
      `| ${sjTime.padEnd(12)}` +
      `| ${skMem.padEnd(9)}` +
      `| ${winner.padEnd(7)}|`
    );
  }

  // SheetKit detailed statistics
  const skResults = results.filter((r) => r.library === 'SheetKit' && r.timesMs.length > 1);
  if (skResults.length > 0) {
    console.log('\n--- SheetKit Detailed Statistics ---');
    console.log(
      '| Scenario'.padEnd(51) +
      '| Median'.padEnd(12) +
      '| Min'.padEnd(12) +
      '| Max'.padEnd(12) +
      '| P95'.padEnd(12) +
      '| Memory |'
    );
    console.log(
      '|' + '-'.repeat(50) +
      '|' + '-'.repeat(11) +
      '|' + '-'.repeat(11) +
      '|' + '-'.repeat(11) +
      '|' + '-'.repeat(11) +
      '|--------|'
    );
    for (const r of skResults) {
      const med = formatMs(median(r.timesMs));
      const min = formatMs(minVal(r.timesMs));
      const max = formatMs(maxVal(r.timesMs));
      const p95v = formatMs(p95(r.timesMs));
      const mem = `${median(r.memoryDeltas).toFixed(1)}MB`;
      console.log(
        `| ${r.scenario.padEnd(49)}` +
        `| ${med.padEnd(10)}` +
        `| ${min.padEnd(10)}` +
        `| ${max.padEnd(10)}` +
        `| ${p95v.padEnd(10)}` +
        `| ${mem.padEnd(7)}|`
      );
    }
  }

  // Wins summary
  const wins: Record<string, number> = { SheetKit: 0, ExcelJS: 0, SheetJS: 0 };
  for (const scenario of scenarios) {
    const times = results
      .filter((r) => r.scenario === scenario)
      .map((r) => ({ lib: r.library, ms: r.timeMs }));
    if (times.length > 0) {
      const winner = times.reduce((a, b) => (a.ms < b.ms ? a : b)).lib;
      wins[winner] = (wins[winner] || 0) + 1;
    }
  }
  console.log('\nWins by library (by median time):');
  for (const [lib, count] of Object.entries(wins)) {
    console.log(`  ${lib}: ${count}/${scenarios.length}`);
  }
}

function generateMarkdownReport(): string {
  const lines: string[] = [];
  lines.push('# Excel Library Benchmark: SheetKit vs ExcelJS vs SheetJS');
  lines.push('');
  lines.push(`Benchmark run: ${new Date().toISOString()}`);
  lines.push(`Platform: ${process.platform} ${process.arch}`);
  lines.push(`Node.js: ${process.version}`);
  lines.push('');
  lines.push('## Methodology');
  lines.push('');
  lines.push(`- **SheetKit**: ${WARMUP_RUNS} warmup run(s) + ${BENCH_RUNS} measured runs per scenario. Median time reported.`);
  lines.push('- **ExcelJS / SheetJS**: Single run per scenario (comparison context only).');
  lines.push('');

  lines.push('## Libraries');
  lines.push('');
  lines.push('| Library | Description |');
  lines.push('|---------|-------------|');
  lines.push('| **SheetKit** (`@sheetkit/node`) | Rust-based Excel library with Node.js bindings via napi-rs |');
  lines.push('| **ExcelJS** (`exceljs`) | Pure JavaScript Excel library with streaming support |');
  lines.push('| **SheetJS** (`xlsx`) | Pure JavaScript spreadsheet library (community edition) |');
  lines.push('');

  lines.push('## Test Fixtures');
  lines.push('');
  lines.push('| Fixture | Description |');
  lines.push('|---------|-------------|');
  lines.push('| `large-data.xlsx` | 50,000 rows x 20 columns, mixed types (numbers, strings, floats, booleans) |');
  lines.push('| `heavy-styles.xlsx` | 5,000 rows x 10 columns with rich formatting |');
  lines.push('| `multi-sheet.xlsx` | 10 sheets, each with 5,000 rows x 10 columns |');
  lines.push('| `formulas.xlsx` | 10,000 rows with 5 formula columns |');
  lines.push('| `strings.xlsx` | 20,000 rows x 10 columns of text data (SST stress test) |');
  lines.push('| `data-validation.xlsx` | 5,000 rows with 8 validation rules (list, whole, decimal, textLength, custom) |');
  lines.push('| `comments.xlsx` | 2,000 rows with cell comments (2,667 total comments) |');
  lines.push('| `merged-cells.xlsx` | 500 merged regions (section headers and sub-headers) |');
  lines.push('| `mixed-workload.xlsx` | Multi-sheet ERP document with styles, formulas, validation, comments |');
  lines.push('| `scale-{1k,10k,100k}.xlsx` | Scaling benchmarks at 1K, 10K, and 100K rows |');
  lines.push('');

  lines.push('## Results');
  lines.push('');

  const categories = [...new Set(results.map((r) => r.category))];

  function renderTable(title: string, subset: BenchResult[]) {
    if (subset.length === 0) return;
    lines.push(`### ${title}`);
    lines.push('');
    lines.push('| Scenario | SheetKit (median) | ExcelJS | SheetJS | Winner |');
    lines.push('|----------|-------------------|---------|---------|--------|');

    const scenarios = [...new Set(subset.map((r) => r.scenario))];
    for (const scenario of scenarios) {
      const sk = subset.find((r) => r.scenario === scenario && r.library === 'SheetKit');
      const ej = subset.find((r) => r.scenario === scenario && r.library === 'ExcelJS');
      const sj = subset.find((r) => r.scenario === scenario && r.library === 'SheetJS');

      const entries = [
        { lib: 'SheetKit', ms: sk?.timeMs },
        { lib: 'ExcelJS', ms: ej?.timeMs },
        { lib: 'SheetJS', ms: sj?.timeMs },
      ].filter((e) => e.ms != null) as { lib: string; ms: number }[];

      const winner = entries.length > 0
        ? entries.reduce((a, b) => (a.ms < b.ms ? a : b)).lib
        : 'N/A';

      const skStr = sk ? formatMs(sk.timeMs) : 'N/A';
      const ejStr = ej ? formatMs(ej.timeMs) : 'N/A';
      const sjStr = sj ? formatMs(sj.timeMs) : 'N/A';

      lines.push(`| ${scenario} | ${skStr} | ${ejStr} | ${sjStr} | ${winner} |`);
    }
    lines.push('');
  }

  for (const cat of categories) {
    const subset = results.filter((r) => r.category === cat);
    renderTable(cat, subset);
  }

  // SheetKit detailed statistics table
  const skResults = results.filter((r) => r.library === 'SheetKit' && r.timesMs.length > 1);
  if (skResults.length > 0) {
    lines.push('### SheetKit Detailed Statistics');
    lines.push('');
    lines.push('| Scenario | Median | Min | Max | P95 | Memory (median) |');
    lines.push('|----------|--------|-----|-----|-----|-----------------|');
    for (const r of skResults) {
      const med = formatMs(median(r.timesMs));
      const min = formatMs(minVal(r.timesMs));
      const max = formatMs(maxVal(r.timesMs));
      const p95v = formatMs(p95(r.timesMs));
      const mem = `${median(r.memoryDeltas).toFixed(1)}MB`;
      lines.push(`| ${r.scenario} | ${med} | ${min} | ${max} | ${p95v} | ${mem} |`);
    }
    lines.push('');
  }

  // Summary
  const allScenarios = [...new Set(results.map((r) => r.scenario))];
  const wins: Record<string, number> = { SheetKit: 0, ExcelJS: 0, SheetJS: 0 };
  for (const scenario of allScenarios) {
    const times = results
      .filter((r) => r.scenario === scenario)
      .map((r) => ({ lib: r.library, ms: r.timeMs }));
    if (times.length > 0) {
      const winner = times.reduce((a, b) => (a.ms < b.ms ? a : b)).lib;
      wins[winner] = (wins[winner] || 0) + 1;
    }
  }

  lines.push('## Summary');
  lines.push('');
  lines.push(`Total scenarios: ${allScenarios.length}`);
  lines.push('');
  lines.push('| Library | Wins |');
  lines.push('|---------|------|');
  for (const [lib, count] of Object.entries(wins).sort((a, b) => b[1] - a[1])) {
    lines.push(`| ${lib} | ${count}/${allScenarios.length} |`);
  }
  lines.push('');

  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
  console.log('Excel Library Benchmark');
  console.log(`Platform: ${process.platform} ${process.arch} | Node.js: ${process.version}`);
  console.log(`SheetKit: ${WARMUP_RUNS} warmup + ${BENCH_RUNS} measured runs per scenario`);
  console.log(`Run with --expose-gc for accurate memory measurements.\n`);

  const { mkdirSync } = await import('node:fs');
  mkdirSync(OUTPUT_DIR, { recursive: true });

  // READ benchmarks
  console.log('=== READ BENCHMARKS ===');
  await benchReadFile('large-data.xlsx', 'Large Data (50k rows x 20 cols)', 'Read');
  await benchReadFile('heavy-styles.xlsx', 'Heavy Styles (5k rows, formatted)', 'Read');
  await benchReadFile('multi-sheet.xlsx', 'Multi-Sheet (10 sheets x 5k rows)', 'Read');
  await benchReadFile('formulas.xlsx', 'Formulas (10k rows)', 'Read');
  await benchReadFile('strings.xlsx', 'Strings (20k rows text-heavy)', 'Read');
  await benchReadFile('data-validation.xlsx', 'Data Validation (5k rows, 8 rules)', 'Read');
  await benchReadFile('comments.xlsx', 'Comments (2k rows with comments)', 'Read');
  await benchReadFile('merged-cells.xlsx', 'Merged Cells (500 regions)', 'Read');
  await benchReadFile('mixed-workload.xlsx', 'Mixed Workload (ERP document)', 'Read');

  // READ scaling
  console.log('\n\n=== READ SCALING ===');
  await benchReadFile('scale-1k.xlsx', 'Scale 1k rows', 'Read (Scale)');
  await benchReadFile('scale-10k.xlsx', 'Scale 10k rows', 'Read (Scale)');
  await benchReadFile('scale-100k.xlsx', 'Scale 100k rows', 'Read (Scale)');

  // WRITE benchmarks
  console.log('\n\n=== WRITE BENCHMARKS ===');
  await benchWriteLargeData();
  await benchWriteWithStyles();
  await benchWriteMultiSheet();
  await benchWriteFormulas();
  await benchWriteStrings();
  await benchWriteDataValidation();
  await benchWriteComments();
  await benchWriteMergedCells();

  // WRITE scaling
  console.log('\n\n=== WRITE SCALING ===');
  await benchWriteScale(1_000);
  await benchWriteScale(10_000);
  await benchWriteScale(50_000);
  await benchWriteScale(100_000);

  // Buffer round-trip
  console.log('\n\n=== BUFFER ROUND-TRIP ===');
  await benchBufferRoundTrip();

  // Streaming
  console.log('\n\n=== STREAMING ===');
  await benchStreamingWrite();

  // Random access
  console.log('\n\n=== RANDOM ACCESS ===');
  await benchRandomAccessRead();

  // Mixed workload write
  console.log('\n\n=== MIXED WORKLOAD ===');
  await benchMixedWorkloadWrite();

  // Summary
  printSummaryTable();

  // Write markdown report
  const report = generateMarkdownReport();
  const reportPath = join(import.meta.dirname, 'RESULTS.md');
  writeFileSync(reportPath, report);
  console.log(`\nMarkdown report written to: ${reportPath}`);
}

main().catch((err) => {
  console.error('Benchmark failed:', err);
  process.exit(1);
});

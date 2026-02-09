/**
 * Excel Library Benchmark: SheetKit vs ExcelJS vs SheetJS (xlsx)
 *
 * Benchmarks:
 * 1. Read large file (50k rows)
 * 2. Read heavy-styles file
 * 3. Read multi-sheet file
 * 4. Read formulas file
 * 5. Read strings file
 * 6. Write large dataset (50k rows x 20 cols)
 * 7. Write with styles (5k rows, formatted)
 * 8. Write multi-sheet (10 sheets x 5k rows)
 * 9. Write formulas (10k rows)
 * 10. Write strings (20k rows text-heavy)
 * 11. Buffer round-trip (write to buffer, read back)
 * 12. Streaming write (50k rows) -- SheetKit and ExcelJS only
 */

import { Workbook as SheetKitWorkbook, JsStreamWriter } from '@sheetkit/node';
import ExcelJS from 'exceljs';
import XLSX from 'xlsx';
import { readFileSync, writeFileSync, existsSync, unlinkSync, statSync } from 'node:fs';
import { join } from 'node:path';

const FIXTURES_DIR = join(import.meta.dirname, 'fixtures');
const OUTPUT_DIR = join(import.meta.dirname, 'output');

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

interface BenchResult {
  library: string;
  scenario: string;
  timeMs: number;
  memoryMb: number;
  fileSizeKb?: number;
}

const results: BenchResult[] = [];

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
  fn: () => void | Promise<void>,
  outputPath?: string,
): Promise<BenchResult> {
  // Warmup GC
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
    timeMs: elapsed,
    memoryMb: memDelta,
    fileSizeKb: size,
  };
  results.push(result);

  const sizeStr = size != null ? ` | ${(size / 1024).toFixed(1)}MB` : '';
  console.log(
    `  [${library.padEnd(8)}] ${scenario.padEnd(40)} ${formatMs(elapsed).padStart(10)} | ${memDelta.toFixed(1).padStart(6)}MB${sizeStr}`,
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

async function benchReadFile(filename: string, label: string) {
  const filepath = join(FIXTURES_DIR, filename);
  if (!existsSync(filepath)) {
    console.log(`  SKIP: ${filepath} not found. Run 'pnpm generate' first.`);
    return;
  }

  console.log(`\n--- ${label} ---`);

  // SheetKit (sync)
  await bench('SheetKit', `Read ${label}`, () => {
    const wb = SheetKitWorkbook.openSync(filepath);
    for (const name of wb.sheetNames) {
      wb.getRows(name);
    }
  });

  // ExcelJS
  await bench('ExcelJS', `Read ${label}`, async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filepath);
    wb.eachSheet((ws) => {
      ws.eachRow(() => { /* iterate */ });
    });
  });

  // SheetJS
  await bench('SheetJS', `Read ${label}`, () => {
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

  // SheetKit
  await bench('SheetKit', label, () => {
    const wb = new SheetKitWorkbook();
    const sheet = 'Sheet1';
    for (let r = 1; r <= ROWS; r++) {
      for (let c = 0; c < COLS; c++) {
        const cell = `${colLetter(c)}${r}`;
        if (c % 3 === 0) wb.setCellValue(sheet, cell, r * (c + 1));
        else if (c % 3 === 1) wb.setCellValue(sheet, cell, `R${r}C${c}`);
        else wb.setCellValue(sheet, cell, (r * c) / 100);
      }
    }
    wb.saveSync(outSk);
  }, outSk);

  // ExcelJS
  await bench('ExcelJS', label, async () => {
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

  // SheetJS
  await bench('SheetJS', label, () => {
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

  cleanup(outSk);
  cleanup(outEj);
  cleanup(outSj);
}

async function benchWriteWithStyles() {
  const ROWS = 5_000;
  const label = `Write ${ROWS} styled rows`;
  console.log(`\n--- ${label} ---`);

  const outSk = join(OUTPUT_DIR, 'write-styles-sheetkit.xlsx');
  const outEj = join(OUTPUT_DIR, 'write-styles-exceljs.xlsx');
  const outSj = join(OUTPUT_DIR, 'write-styles-sheetjs.xlsx');

  // SheetKit
  await bench('SheetKit', label, () => {
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

    for (let c = 0; c < 10; c++) {
      wb.setCellValue(sheet, `${colLetter(c)}1`, `Header${c + 1}`);
      wb.setCellStyle(sheet, `${colLetter(c)}1`, boldId);
    }

    for (let r = 2; r <= ROWS + 1; r++) {
      for (let c = 0; c < 10; c++) {
        const cell = `${colLetter(c)}${r}`;
        if (c % 3 === 0) {
          wb.setCellValue(sheet, cell, r * c);
          wb.setCellStyle(sheet, cell, numId);
        } else if (c % 3 === 1) {
          wb.setCellValue(sheet, cell, `Data_${r}_${c}`);
        } else {
          wb.setCellValue(sheet, cell, (r % 100) / 100);
          wb.setCellStyle(sheet, cell, pctId);
        }
      }
    }
    wb.saveSync(outSk);
  }, outSk);

  // ExcelJS
  await bench('ExcelJS', label, async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');

    const headerRow = ws.addRow(Array.from({ length: 10 }, (_, c) => `Header${c + 1}`));
    headerRow.eachCell((cell) => {
      cell.font = { bold: true, size: 12, name: 'Arial', color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' },
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

  // SheetJS (limited style support in community edition)
  await bench('SheetJS', label, () => {
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

  cleanup(outSk);
  cleanup(outEj);
  cleanup(outSj);
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

  // SheetKit
  await bench('SheetKit', label, () => {
    const wb = new SheetKitWorkbook();
    for (let s = 0; s < SHEETS; s++) {
      const name = s === 0 ? 'Sheet1' : `Sheet${s + 1}`;
      if (s > 0) wb.newSheet(name);
      for (let r = 1; r <= ROWS; r++) {
        for (let c = 0; c < COLS; c++) {
          const cell = `${colLetter(c)}${r}`;
          if (c % 2 === 0) wb.setCellValue(name, cell, r * (c + 1));
          else wb.setCellValue(name, cell, `S${s}R${r}C${c}`);
        }
      }
    }
    wb.saveSync(outSk);
  }, outSk);

  // ExcelJS
  await bench('ExcelJS', label, async () => {
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

  // SheetJS
  await bench('SheetJS', label, () => {
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

  cleanup(outSk);
  cleanup(outEj);
  cleanup(outSj);
}

async function benchWriteFormulas() {
  const ROWS = 10_000;
  const label = `Write ${ROWS} rows with formulas`;
  console.log(`\n--- ${label} ---`);

  const outSk = join(OUTPUT_DIR, 'write-formulas-sheetkit.xlsx');
  const outEj = join(OUTPUT_DIR, 'write-formulas-exceljs.xlsx');
  const outSj = join(OUTPUT_DIR, 'write-formulas-sheetjs.xlsx');

  // SheetKit
  await bench('SheetKit', label, () => {
    const wb = new SheetKitWorkbook();
    const sheet = 'Sheet1';
    for (let r = 1; r <= ROWS; r++) {
      wb.setCellValue(sheet, `A${r}`, r * 1.5);
      wb.setCellValue(sheet, `B${r}`, (r % 100) + 0.5);
      wb.setCellFormula(sheet, `C${r}`, `A${r}+B${r}`);
      wb.setCellFormula(sheet, `D${r}`, `A${r}*B${r}`);
      wb.setCellFormula(sheet, `E${r}`, `IF(A${r}>B${r},"A","B")`);
    }
    wb.saveSync(outSk);
  }, outSk);

  // ExcelJS
  await bench('ExcelJS', label, async () => {
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

  // SheetJS
  await bench('SheetJS', label, () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([]);
    for (let r = 0; r < ROWS; r++) {
      const rowNum = r; // 0-based for SheetJS
      XLSX.utils.sheet_add_aoa(ws, [[r * 1.5 + 1.5, ((r + 1) % 100) + 0.5]], { origin: `A${r + 1}` });
      ws[`C${r + 1}`] = { t: 's', f: `A${r + 1}+B${r + 1}` };
      ws[`D${r + 1}`] = { t: 's', f: `A${r + 1}*B${r + 1}` };
      ws[`E${r + 1}`] = { t: 's', f: `IF(A${r + 1}>B${r + 1},"A","B")` };
    }
    ws['!ref'] = `A1:E${ROWS}`;
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, outSj);
  }, outSj);

  cleanup(outSk);
  cleanup(outEj);
  cleanup(outSj);
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

  // SheetKit
  await bench('SheetKit', label, () => {
    const wb = new SheetKitWorkbook();
    const sheet = 'Sheet1';
    for (let r = 1; r <= ROWS; r++) {
      const row = makeRow(r);
      for (let c = 0; c < COLS; c++) {
        wb.setCellValue(sheet, `${colLetter(c)}${r}`, row[c]);
      }
    }
    wb.saveSync(outSk);
  }, outSk);

  // ExcelJS
  await bench('ExcelJS', label, async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    for (let r = 1; r <= ROWS; r++) {
      ws.addRow(makeRow(r));
    }
    await wb.xlsx.writeFile(outEj);
  }, outEj);

  // SheetJS
  await bench('SheetJS', label, () => {
    const data: string[][] = [];
    for (let r = 1; r <= ROWS; r++) {
      data.push(makeRow(r));
    }
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, outSj);
  }, outSj);

  cleanup(outSk);
  cleanup(outEj);
  cleanup(outSj);
}

// ---------------------------------------------------------------------------
// Buffer round-trip
// ---------------------------------------------------------------------------

async function benchBufferRoundTrip() {
  const ROWS = 10_000;
  const COLS = 10;
  const label = `Buffer round-trip (${ROWS} rows)`;
  console.log(`\n--- ${label} ---`);

  // SheetKit
  await bench('SheetKit', label, () => {
    const wb = new SheetKitWorkbook();
    const sheet = 'Sheet1';
    for (let r = 1; r <= ROWS; r++) {
      for (let c = 0; c < COLS; c++) {
        wb.setCellValue(sheet, `${colLetter(c)}${r}`, r * (c + 1));
      }
    }
    const buf = wb.writeBufferSync();
    const wb2 = SheetKitWorkbook.openBufferSync(buf);
    wb2.getRows('Sheet1');
  });

  // ExcelJS
  await bench('ExcelJS', label, async () => {
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

  // SheetJS
  await bench('SheetJS', label, () => {
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

  // SheetKit stream writer
  await bench('SheetKit', label, () => {
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

  // ExcelJS streaming
  await bench('ExcelJS', label, async () => {
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

  // SheetJS has no streaming API -- skip
  console.log('  [SheetJS  ] Streaming write                              N/A (no streaming API)');

  cleanup(outSk);
  cleanup(outEj);
}

// ---------------------------------------------------------------------------
// Results formatting
// ---------------------------------------------------------------------------

function printSummaryTable() {
  console.log('\n\n========================================');
  console.log(' BENCHMARK RESULTS SUMMARY');
  console.log('========================================\n');

  // Group by scenario
  const scenarios = [...new Set(results.map((r) => r.scenario))];

  console.log(
    '| Scenario'.padEnd(45) +
    '| SheetKit'.padEnd(14) +
    '| ExcelJS'.padEnd(14) +
    '| SheetJS'.padEnd(14) +
    '| Winner |'
  );
  console.log(
    '|' + '-'.repeat(44) +
    '|' + '-'.repeat(13) +
    '|' + '-'.repeat(13) +
    '|' + '-'.repeat(13) +
    '|--------|'
  );

  for (const scenario of scenarios) {
    const sk = results.find((r) => r.scenario === scenario && r.library === 'SheetKit');
    const ej = results.find((r) => r.scenario === scenario && r.library === 'ExcelJS');
    const sj = results.find((r) => r.scenario === scenario && r.library === 'SheetJS');

    const skTime = sk ? formatMs(sk.timeMs) : 'N/A';
    const ejTime = ej ? formatMs(ej.timeMs) : 'N/A';
    const sjTime = sj ? formatMs(sj.timeMs) : 'N/A';

    const times = [
      { lib: 'SheetKit', ms: sk?.timeMs ?? Infinity },
      { lib: 'ExcelJS', ms: ej?.timeMs ?? Infinity },
      { lib: 'SheetJS', ms: sj?.timeMs ?? Infinity },
    ].filter((t) => t.ms < Infinity);

    const winner = times.length > 0
      ? times.reduce((a, b) => (a.ms < b.ms ? a : b)).lib
      : 'N/A';

    console.log(
      `| ${scenario.padEnd(43)}` +
      `| ${skTime.padEnd(12)}` +
      `| ${ejTime.padEnd(12)}` +
      `| ${sjTime.padEnd(12)}` +
      `| ${winner.padEnd(7)}|`
    );
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
  console.log('\nWins by library:');
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
  lines.push('| `heavy-styles.xlsx` | 5,000 rows x 10 columns with rich formatting (fonts, fills, borders, number formats) |');
  lines.push('| `multi-sheet.xlsx` | 10 sheets, each with 5,000 rows x 10 columns |');
  lines.push('| `formulas.xlsx` | 10,000 rows with 5 formula columns (SUM, PRODUCT, AVERAGE, MAX, IF) |');
  lines.push('| `strings.xlsx` | 20,000 rows x 10 columns of text data (SST stress test) |');
  lines.push('');

  lines.push('## Results');
  lines.push('');

  // Group into read/write
  const readResults = results.filter((r) => r.scenario.startsWith('Read'));
  const writeResults = results.filter((r) => !r.scenario.startsWith('Read') && !r.scenario.startsWith('Buffer') && !r.scenario.startsWith('Streaming'));
  const bufferResults = results.filter((r) => r.scenario.startsWith('Buffer'));
  const streamResults = results.filter((r) => r.scenario.startsWith('Streaming'));

  function renderTable(title: string, subset: BenchResult[]) {
    if (subset.length === 0) return;
    lines.push(`### ${title}`);
    lines.push('');
    lines.push('| Scenario | SheetKit | ExcelJS | SheetJS | Winner |');
    lines.push('|----------|----------|---------|---------|--------|');

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

      const winner = entries.reduce((a, b) => (a.ms < b.ms ? a : b)).lib;

      const skStr = sk ? formatMs(sk.timeMs) : 'N/A';
      const ejStr = ej ? formatMs(ej.timeMs) : 'N/A';
      const sjStr = sj ? formatMs(sj.timeMs) : 'N/A';

      lines.push(`| ${scenario} | ${skStr} | ${ejStr} | ${sjStr} | ${winner} |`);
    }
    lines.push('');
  }

  renderTable('Read Performance', readResults);
  renderTable('Write Performance', writeResults);
  renderTable('Buffer Round-Trip', bufferResults);
  renderTable('Streaming Write', streamResults);

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
  console.log(`Run with --expose-gc for accurate memory measurements.\n`);

  // Ensure output dir
  const { mkdirSync } = await import('node:fs');
  mkdirSync(OUTPUT_DIR, { recursive: true });

  // READ benchmarks
  console.log('=== READ BENCHMARKS ===');
  await benchReadFile('large-data.xlsx', 'Large Data (50k rows x 20 cols)');
  await benchReadFile('heavy-styles.xlsx', 'Heavy Styles (5k rows, formatted)');
  await benchReadFile('multi-sheet.xlsx', 'Multi-Sheet (10 sheets x 5k rows)');
  await benchReadFile('formulas.xlsx', 'Formulas (10k rows)');
  await benchReadFile('strings.xlsx', 'Strings (20k rows text-heavy)');

  // WRITE benchmarks
  console.log('\n\n=== WRITE BENCHMARKS ===');
  await benchWriteLargeData();
  await benchWriteWithStyles();
  await benchWriteMultiSheet();
  await benchWriteFormulas();
  await benchWriteStrings();

  // Buffer round-trip
  console.log('\n\n=== BUFFER ROUND-TRIP ===');
  await benchBufferRoundTrip();

  // Streaming
  console.log('\n\n=== STREAMING ===');
  await benchStreamingWrite();

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

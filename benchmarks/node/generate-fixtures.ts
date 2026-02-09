/**
 * Generates test Excel files for benchmarking.
 *
 * Fixture types:
 * 1. large-data: 50,000 rows x 20 columns of mixed types
 * 2. heavy-styles: 5,000 rows with various cell formatting
 * 3. multi-sheet: 10 sheets, each with 5,000 rows
 * 4. formulas: 10,000 rows with computed columns
 * 5. strings: 20,000 rows of text-heavy data (SST stress test)
 */

import { Workbook } from '@sheetkit/node';
import { mkdirSync, existsSync } from 'node:fs';
import { join } from 'node:path';

const FIXTURES_DIR = join(import.meta.dirname, 'fixtures');

function colLetter(n: number): string {
  let s = '';
  while (n >= 0) {
    s = String.fromCharCode((n % 26) + 65) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

function cellRef(row: number, col: number): string {
  return `${colLetter(col)}${row}`;
}

function generateLargeData() {
  console.log('Generating large-data.xlsx (50000 rows x 20 cols)...');
  const wb = new Workbook();
  const sheet = 'Sheet1';

  // Header row
  for (let c = 0; c < 20; c++) {
    wb.setCellValue(sheet, cellRef(1, c), `Column_${c + 1}`);
  }

  // Data rows
  for (let r = 2; r <= 50001; r++) {
    for (let c = 0; c < 20; c++) {
      const mod = c % 4;
      if (mod === 0) {
        wb.setCellValue(sheet, cellRef(r, c), r * (c + 1));
      } else if (mod === 1) {
        wb.setCellValue(sheet, cellRef(r, c), `Row${r}_Col${c + 1}`);
      } else if (mod === 2) {
        wb.setCellValue(sheet, cellRef(r, c), (r * (c + 1)) / 100.0);
      } else {
        wb.setCellValue(sheet, cellRef(r, c), r % 2 === 0);
      }
    }
  }

  wb.saveSync(join(FIXTURES_DIR, 'large-data.xlsx'));
  console.log('  Done.');
}

function generateHeavyStyles() {
  console.log('Generating heavy-styles.xlsx (5000 rows, rich formatting)...');
  const wb = new Workbook();
  const sheet = 'Sheet1';

  // Pre-register styles
  const boldStyle = wb.addStyle({
    font: { bold: true, size: 12, name: 'Arial' },
    fill: { pattern: 'solid', fgColor: 'FF4472C4' },
    border: {
      top: { style: 'thin', color: 'FF000000' },
      bottom: { style: 'thin', color: 'FF000000' },
      left: { style: 'thin', color: 'FF000000' },
      right: { style: 'thin', color: 'FF000000' },
    },
    alignment: { horizontal: 'center', vertical: 'center' },
  });

  const numberStyle = wb.addStyle({
    numFmtId: 4, // #,##0.00
    font: { name: 'Calibri', size: 11 },
    border: {
      bottom: { style: 'hair', color: 'FFCCCCCC' },
    },
  });

  const percentStyle = wb.addStyle({
    numFmtId: 10, // 0.00%
    font: { name: 'Calibri', size: 11, italic: true },
    fill: { pattern: 'solid', fgColor: 'FFFFE699' },
  });

  const dateStyle = wb.addStyle({
    customNumFmt: 'yyyy-mm-dd',
    font: { name: 'Calibri', size: 11 },
  });

  const wrapStyle = wb.addStyle({
    alignment: { wrapText: true },
    font: { name: 'Calibri', size: 10 },
    fill: { pattern: 'solid', fgColor: 'FFF2F2F2' },
  });

  const styles = [boldStyle, numberStyle, percentStyle, dateStyle, wrapStyle];

  // Header
  const headers = [
    'ID', 'Name', 'Amount', 'Rate', 'Date',
    'Category', 'Score', 'Percent', 'Notes', 'Status',
  ];
  for (let c = 0; c < headers.length; c++) {
    wb.setCellValue(sheet, cellRef(1, c), headers[c]);
    wb.setCellStyle(sheet, cellRef(1, c), boldStyle);
  }

  for (let r = 2; r <= 5001; r++) {
    for (let c = 0; c < 10; c++) {
      const mod = c % 5;
      if (mod === 0) {
        wb.setCellValue(sheet, cellRef(r, c), r - 1);
      } else if (mod === 1) {
        wb.setCellValue(sheet, cellRef(r, c), `Employee_${r - 1}`);
      } else if (mod === 2) {
        wb.setCellValue(sheet, cellRef(r, c), (r * 123.45) % 99999);
      } else if (mod === 3) {
        wb.setCellValue(sheet, cellRef(r, c), (r % 100) / 100.0);
      } else {
        wb.setCellValue(sheet, cellRef(r, c), `Notes for row ${r - 1}`);
      }
      wb.setCellStyle(sheet, cellRef(r, c), styles[mod]);
    }
  }

  wb.saveSync(join(FIXTURES_DIR, 'heavy-styles.xlsx'));
  console.log('  Done.');
}

function generateMultiSheet() {
  console.log('Generating multi-sheet.xlsx (10 sheets x 5000 rows)...');
  const wb = new Workbook();

  for (let s = 0; s < 10; s++) {
    const sheetName = s === 0 ? 'Sheet1' : `Sheet${s + 1}`;
    if (s > 0) wb.newSheet(sheetName);

    for (let c = 0; c < 10; c++) {
      wb.setCellValue(sheetName, cellRef(1, c), `Header_${c + 1}`);
    }

    for (let r = 2; r <= 5001; r++) {
      for (let c = 0; c < 10; c++) {
        if (c % 2 === 0) {
          wb.setCellValue(sheetName, cellRef(r, c), r * (c + 1) + s * 1000);
        } else {
          wb.setCellValue(sheetName, cellRef(r, c), `S${s + 1}_R${r}_C${c + 1}`);
        }
      }
    }
  }

  wb.saveSync(join(FIXTURES_DIR, 'multi-sheet.xlsx'));
  console.log('  Done.');
}

function generateFormulas() {
  console.log('Generating formulas.xlsx (10000 rows with formulas)...');
  const wb = new Workbook();
  const sheet = 'Sheet1';

  const headers = ['Value_A', 'Value_B', 'Sum', 'Product', 'Average', 'Max', 'Condition'];
  for (let c = 0; c < headers.length; c++) {
    wb.setCellValue(sheet, cellRef(1, c), headers[c]);
  }

  for (let r = 2; r <= 10001; r++) {
    wb.setCellValue(sheet, cellRef(r, 0), r * 1.5);
    wb.setCellValue(sheet, cellRef(r, 1), (r % 100) + 0.5);
    wb.setCellFormula(sheet, cellRef(r, 2), `A${r}+B${r}`);
    wb.setCellFormula(sheet, cellRef(r, 3), `A${r}*B${r}`);
    wb.setCellFormula(sheet, cellRef(r, 4), `AVERAGE(A${r},B${r})`);
    wb.setCellFormula(sheet, cellRef(r, 5), `MAX(A${r},B${r})`);
    wb.setCellFormula(sheet, cellRef(r, 6), `IF(A${r}>B${r},"A","B")`);
  }

  wb.saveSync(join(FIXTURES_DIR, 'formulas.xlsx'));
  console.log('  Done.');
}

function generateStrings() {
  console.log('Generating strings.xlsx (20000 rows of text)...');
  const wb = new Workbook();
  const sheet = 'Sheet1';

  const words = [
    'alpha', 'bravo', 'charlie', 'delta', 'echo',
    'foxtrot', 'golf', 'hotel', 'india', 'juliet',
    'kilo', 'lima', 'mike', 'november', 'oscar',
    'papa', 'quebec', 'romeo', 'sierra', 'tango',
  ];

  const headers = ['Full_Name', 'Email', 'Department', 'Address', 'Description',
    'City', 'Country', 'Phone', 'Title', 'Bio'];
  for (let c = 0; c < headers.length; c++) {
    wb.setCellValue(sheet, cellRef(1, c), headers[c]);
  }

  for (let r = 2; r <= 20001; r++) {
    const w1 = words[(r * 3) % words.length];
    const w2 = words[(r * 7) % words.length];
    const w3 = words[(r * 11) % words.length];

    wb.setCellValue(sheet, cellRef(r, 0), `${w1} ${w2} ${w3}`);
    wb.setCellValue(sheet, cellRef(r, 1), `${w1}.${w2}@example.com`);
    wb.setCellValue(sheet, cellRef(r, 2), `Department_${(r % 20) + 1}`);
    wb.setCellValue(sheet, cellRef(r, 3), `${r} ${w3} Street, Suite ${r % 500}`);
    wb.setCellValue(sheet, cellRef(r, 4), `Description for entry ${r}: ${w1} ${w2} ${w3} related work`);
    wb.setCellValue(sheet, cellRef(r, 5), `City_${w1}`);
    wb.setCellValue(sheet, cellRef(r, 6), `Country_${w2}`);
    wb.setCellValue(sheet, cellRef(r, 7), `+1-555-${String(r % 10000).padStart(4, '0')}`);
    wb.setCellValue(sheet, cellRef(r, 8), `${w1} ${w2} Specialist`);
    wb.setCellValue(sheet, cellRef(r, 9), `Experienced ${w3} professional with expertise in ${w1} and ${w2} domains.`);
  }

  wb.saveSync(join(FIXTURES_DIR, 'strings.xlsx'));
  console.log('  Done.');
}

// Main
if (!existsSync(FIXTURES_DIR)) {
  mkdirSync(FIXTURES_DIR, { recursive: true });
}

console.log(`Generating fixtures in ${FIXTURES_DIR}\n`);

const start = performance.now();
generateLargeData();
generateHeavyStyles();
generateMultiSheet();
generateFormulas();
generateStrings();

const elapsed = ((performance.now() - start) / 1000).toFixed(1);
console.log(`\nAll fixtures generated in ${elapsed}s`);

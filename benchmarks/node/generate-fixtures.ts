/**
 * Generates test Excel files for benchmarking.
 *
 * Fixture types:
 * 1. large-data: 50,000 rows x 20 columns of mixed types
 * 2. heavy-styles: 5,000 rows with various cell formatting
 * 3. multi-sheet: 10 sheets, each with 5,000 rows
 * 4. formulas: 10,000 rows with computed columns
 * 5. strings: 20,000 rows of text-heavy data (SST stress test)
 * 6. data-validation: 5,000 rows with various validation rules
 * 7. comments: 2,000 rows with cell comments
 * 8. merged-cells: worksheet with many merged regions
 * 9. mixed-workload: realistic ERP-style document with all features
 * 10. scale-1k / scale-10k / scale-100k: scaling benchmarks
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

  for (let c = 0; c < 20; c++) {
    wb.setCellValue(sheet, cellRef(1, c), `Column_${c + 1}`);
  }

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
    numFmtId: 4,
    font: { name: 'Calibri', size: 11 },
    border: {
      bottom: { style: 'hair', color: 'FFCCCCCC' },
    },
  });

  const percentStyle = wb.addStyle({
    numFmtId: 10,
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

function generateDataValidation() {
  console.log('Generating data-validation.xlsx (5000 rows with validation rules)...');
  const wb = new Workbook();
  const sheet = 'Sheet1';

  const headers = [
    'Status', 'Priority', 'Score', 'Rate', 'Quantity',
    'Department', 'Code', 'Notes',
  ];
  for (let c = 0; c < headers.length; c++) {
    wb.setCellValue(sheet, cellRef(1, c), headers[c]);
  }

  // List validation on column A (Status)
  wb.addDataValidation(sheet, {
    sqref: `A2:A5001`,
    validationType: 'list',
    formula1: '"Active,Inactive,Pending,Closed,Archived"',
    errorTitle: 'Invalid Status',
    errorMessage: 'Select a valid status value.',
    showErrorMessage: true,
  });

  // List validation on column B (Priority)
  wb.addDataValidation(sheet, {
    sqref: `B2:B5001`,
    validationType: 'list',
    formula1: '"Critical,High,Medium,Low"',
  });

  // Whole number validation on column C (Score: 0-100)
  wb.addDataValidation(sheet, {
    sqref: `C2:C5001`,
    validationType: 'whole',
    operator: 'between',
    formula1: '0',
    formula2: '100',
    errorTitle: 'Invalid Score',
    errorMessage: 'Score must be between 0 and 100.',
    showErrorMessage: true,
  });

  // Decimal validation on column D (Rate: 0.0-1.0)
  wb.addDataValidation(sheet, {
    sqref: `D2:D5001`,
    validationType: 'decimal',
    operator: 'between',
    formula1: '0',
    formula2: '1',
    promptTitle: 'Rate',
    promptMessage: 'Enter a decimal between 0 and 1.',
    showInputMessage: true,
  });

  // Whole number validation on column E (Quantity: >= 0)
  wb.addDataValidation(sheet, {
    sqref: `E2:E5001`,
    validationType: 'whole',
    operator: 'greaterThanOrEqual',
    formula1: '0',
  });

  // Text length validation on column F (Department: max 30 chars)
  wb.addDataValidation(sheet, {
    sqref: `F2:F5001`,
    validationType: 'textLength',
    operator: 'lessThanOrEqual',
    formula1: '30',
    errorTitle: 'Too Long',
    errorMessage: 'Department name must be 30 characters or less.',
    showErrorMessage: true,
  });

  // Custom formula validation on column G (Code: must start with letter)
  wb.addDataValidation(sheet, {
    sqref: `G2:G5001`,
    validationType: 'custom',
    formula1: 'ISNUMBER(FIND(LEFT(G2,1),"ABCDEFGHIJKLMNOPQRSTUVWXYZ"))',
  });

  // Text length validation on column H (Notes: 0-500 chars)
  wb.addDataValidation(sheet, {
    sqref: `H2:H5001`,
    validationType: 'textLength',
    operator: 'between',
    formula1: '0',
    formula2: '500',
  });

  // Fill data
  const statuses = ['Active', 'Inactive', 'Pending', 'Closed', 'Archived'];
  const priorities = ['Critical', 'High', 'Medium', 'Low'];
  const depts = ['Engineering', 'Sales', 'Marketing', 'Finance', 'Support', 'HR', 'Legal', 'Operations'];

  for (let r = 2; r <= 5001; r++) {
    wb.setCellValue(sheet, cellRef(r, 0), statuses[r % statuses.length]);
    wb.setCellValue(sheet, cellRef(r, 1), priorities[r % priorities.length]);
    wb.setCellValue(sheet, cellRef(r, 2), r % 101);
    wb.setCellValue(sheet, cellRef(r, 3), (r % 100) / 100.0);
    wb.setCellValue(sheet, cellRef(r, 4), (r * 7) % 500);
    wb.setCellValue(sheet, cellRef(r, 5), depts[r % depts.length]);
    wb.setCellValue(sheet, cellRef(r, 6), `X${String(r).padStart(5, '0')}`);
    wb.setCellValue(sheet, cellRef(r, 7), `Notes for item ${r}`);
  }

  wb.saveSync(join(FIXTURES_DIR, 'data-validation.xlsx'));
  console.log('  Done.');
}

function generateComments() {
  console.log('Generating comments.xlsx (2000 rows with comments)...');
  const wb = new Workbook();
  const sheet = 'Sheet1';

  const headers = ['ID', 'Name', 'Score', 'Status', 'Review_Notes'];
  for (let c = 0; c < headers.length; c++) {
    wb.setCellValue(sheet, cellRef(1, c), headers[c]);
  }

  for (let r = 2; r <= 2001; r++) {
    wb.setCellValue(sheet, cellRef(r, 0), r - 1);
    wb.setCellValue(sheet, cellRef(r, 1), `Employee_${r - 1}`);
    wb.setCellValue(sheet, cellRef(r, 2), (r * 17) % 100);
    wb.setCellValue(sheet, cellRef(r, 3), r % 3 === 0 ? 'Flagged' : 'OK');
    wb.setCellValue(sheet, cellRef(r, 4), `Review note for row ${r}`);

    // Add comment on every row's score cell
    wb.addComment(sheet, {
      cell: cellRef(r, 2),
      author: 'Reviewer',
      text: `Score reviewed on 2025-01-${String((r % 28) + 1).padStart(2, '0')}. ${r % 3 === 0 ? 'Needs attention.' : 'Acceptable.'}`,
    });

    // Add comment on flagged rows' status cell
    if (r % 3 === 0) {
      wb.addComment(sheet, {
        cell: cellRef(r, 3),
        author: 'Manager',
        text: `Flagged for review. Priority: ${r % 2 === 0 ? 'High' : 'Medium'}.`,
      });
    }
  }

  wb.saveSync(join(FIXTURES_DIR, 'comments.xlsx'));
  console.log('  Done.');
}

function generateMergedCells() {
  console.log('Generating merged-cells.xlsx (500 merged regions)...');
  const wb = new Workbook();
  const sheet = 'Sheet1';

  let currentRow = 1;

  // Generate 100 section headers (each is a merged title row spanning 8 columns)
  for (let section = 0; section < 100; section++) {
    // Title row: merge A:H
    wb.setCellValue(sheet, cellRef(currentRow, 0), `Section ${section + 1}: Category Report`);
    wb.mergeCells(sheet, cellRef(currentRow, 0), cellRef(currentRow, 7));
    currentRow++;

    // Sub-headers: merge pairs (A:B, C:D, E:F, G:H)
    for (let pair = 0; pair < 4; pair++) {
      const leftCol = pair * 2;
      wb.setCellValue(sheet, cellRef(currentRow, leftCol), `SubHeader_${pair + 1}`);
      wb.mergeCells(sheet, cellRef(currentRow, leftCol), cellRef(currentRow, leftCol + 1));
    }
    currentRow++;

    // Data rows (3 rows per section)
    for (let dr = 0; dr < 3; dr++) {
      for (let c = 0; c < 8; c++) {
        wb.setCellValue(sheet, cellRef(currentRow, c), section * 3 + dr + c);
      }
      currentRow++;
    }
  }

  wb.saveSync(join(FIXTURES_DIR, 'merged-cells.xlsx'));
  console.log('  Done.');
}

function generateMixedWorkload() {
  console.log('Generating mixed-workload.xlsx (realistic ERP document)...');
  const wb = new Workbook();

  // Sheet 1: Invoice list with styles, formulas, validation
  const invoiceSheet = 'Sheet1';
  const boldId = wb.addStyle({
    font: { bold: true, size: 12, name: 'Arial', color: 'FFFFFFFF' },
    fill: { pattern: 'solid', fgColor: 'FF2F5496' },
    alignment: { horizontal: 'center' },
    border: {
      top: { style: 'thin', color: 'FF000000' },
      bottom: { style: 'thin', color: 'FF000000' },
      left: { style: 'thin', color: 'FF000000' },
      right: { style: 'thin', color: 'FF000000' },
    },
  });
  const currencyId = wb.addStyle({ customNumFmt: '$#,##0.00' });
  const pctId = wb.addStyle({ numFmtId: 10 });

  // Title (merged)
  wb.setCellValue(invoiceSheet, 'A1', 'Invoice Summary Report');
  wb.mergeCells(invoiceSheet, 'A1', 'H1');
  wb.setCellStyle(invoiceSheet, 'A1', boldId);

  const invHeaders = ['Invoice_ID', 'Customer', 'Amount', 'Tax', 'Total', 'Status', 'Due_Date', 'Notes'];
  for (let c = 0; c < invHeaders.length; c++) {
    wb.setCellValue(invoiceSheet, cellRef(2, c), invHeaders[c]);
    wb.setCellStyle(invoiceSheet, cellRef(2, c), boldId);
  }

  wb.addDataValidation(invoiceSheet, {
    sqref: 'F3:F3002',
    validationType: 'list',
    formula1: '"Paid,Unpaid,Overdue,Cancelled"',
  });

  const customers = ['Acme Corp', 'GlobalTech', 'Initech', 'Umbrella Inc', 'Stark Industries',
    'Wayne Enterprises', 'Oscorp', 'LexCorp', 'Massive Dynamic', 'Soylent Corp'];
  const statuses = ['Paid', 'Unpaid', 'Overdue', 'Cancelled'];

  for (let r = 3; r <= 3002; r++) {
    const i = r - 2;
    wb.setCellValue(invoiceSheet, cellRef(r, 0), `INV-${String(i).padStart(5, '0')}`);
    wb.setCellValue(invoiceSheet, cellRef(r, 1), customers[i % customers.length]);
    wb.setCellValue(invoiceSheet, cellRef(r, 2), Math.round((i * 123.45) % 9999 * 100) / 100);
    wb.setCellStyle(invoiceSheet, cellRef(r, 2), currencyId);
    wb.setCellFormula(invoiceSheet, cellRef(r, 3), `C${r}*0.1`);
    wb.setCellStyle(invoiceSheet, cellRef(r, 3), currencyId);
    wb.setCellFormula(invoiceSheet, cellRef(r, 4), `C${r}+D${r}`);
    wb.setCellStyle(invoiceSheet, cellRef(r, 4), currencyId);
    wb.setCellValue(invoiceSheet, cellRef(r, 5), statuses[i % statuses.length]);
    wb.setCellValue(invoiceSheet, cellRef(r, 6), 45000 + (i % 365));
    wb.setCellValue(invoiceSheet, cellRef(r, 7), i % 5 === 0 ? `Special terms for invoice ${i}` : '');
  }

  // Comments on overdue invoices
  for (let r = 3; r <= 3002; r++) {
    if ((r - 2) % statuses.length === 2) {
      wb.addComment(invoiceSheet, {
        cell: cellRef(r, 5),
        author: 'Accounts',
        text: `Overdue since ${30 + ((r * 3) % 60)} days. Follow up required.`,
      });
    }
  }

  // Sheet 2: Employee directory with many columns
  const empSheet = 'Employees';
  wb.newSheet(empSheet);

  const empHeaders = ['EmpID', 'First_Name', 'Last_Name', 'Email', 'Department',
    'Title', 'Salary', 'Bonus_Rate', 'Total_Comp', 'Start_Date',
    'Manager', 'Location', 'Phone', 'Status', 'Notes'];
  for (let c = 0; c < empHeaders.length; c++) {
    wb.setCellValue(empSheet, cellRef(1, c), empHeaders[c]);
    wb.setCellStyle(empSheet, cellRef(1, c), boldId);
  }

  wb.addDataValidation(empSheet, {
    sqref: 'E2:E2001',
    validationType: 'list',
    formula1: '"Engineering,Sales,Marketing,Finance,HR,Legal,Operations,Support"',
  });

  wb.addDataValidation(empSheet, {
    sqref: 'G2:G2001',
    validationType: 'decimal',
    operator: 'between',
    formula1: '30000',
    formula2: '500000',
  });

  wb.addDataValidation(empSheet, {
    sqref: 'H2:H2001',
    validationType: 'decimal',
    operator: 'between',
    formula1: '0',
    formula2: '0.5',
  });

  const firstNames = ['James', 'Mary', 'Robert', 'Patricia', 'John', 'Jennifer', 'Michael', 'Linda'];
  const lastNames = ['Smith', 'Johnson', 'Williams', 'Brown', 'Jones', 'Garcia', 'Miller', 'Davis'];
  const depts = ['Engineering', 'Sales', 'Marketing', 'Finance', 'HR', 'Legal', 'Operations', 'Support'];
  const titles = ['Analyst', 'Engineer', 'Manager', 'Director', 'VP', 'Lead', 'Specialist', 'Coordinator'];
  const locations = ['New York', 'San Francisco', 'London', 'Tokyo', 'Berlin', 'Sydney', 'Toronto', 'Singapore'];

  for (let r = 2; r <= 2001; r++) {
    const i = r - 1;
    const fn = firstNames[i % firstNames.length];
    const ln = lastNames[(i * 3) % lastNames.length];
    wb.setCellValue(empSheet, cellRef(r, 0), `E${String(i).padStart(5, '0')}`);
    wb.setCellValue(empSheet, cellRef(r, 1), fn);
    wb.setCellValue(empSheet, cellRef(r, 2), ln);
    wb.setCellValue(empSheet, cellRef(r, 3), `${fn.toLowerCase()}.${ln.toLowerCase()}@company.com`);
    wb.setCellValue(empSheet, cellRef(r, 4), depts[i % depts.length]);
    wb.setCellValue(empSheet, cellRef(r, 5), `${titles[i % titles.length]}`);
    wb.setCellValue(empSheet, cellRef(r, 6), 50000 + (i * 137) % 150000);
    wb.setCellStyle(empSheet, cellRef(r, 6), currencyId);
    wb.setCellValue(empSheet, cellRef(r, 7), ((i % 20) + 5) / 100);
    wb.setCellStyle(empSheet, cellRef(r, 7), pctId);
    wb.setCellFormula(empSheet, cellRef(r, 8), `G${r}*(1+H${r})`);
    wb.setCellStyle(empSheet, cellRef(r, 8), currencyId);
    wb.setCellValue(empSheet, cellRef(r, 9), 43000 + (i % 1500));
    wb.setCellValue(empSheet, cellRef(r, 10), `MGR_${Math.floor(i / 10)}`);
    wb.setCellValue(empSheet, cellRef(r, 11), locations[i % locations.length]);
    wb.setCellValue(empSheet, cellRef(r, 12), `+1-555-${String(i % 10000).padStart(4, '0')}`);
    wb.setCellValue(empSheet, cellRef(r, 13), i % 20 === 0 ? 'Inactive' : 'Active');
    if (i % 10 === 0) {
      wb.setCellValue(empSheet, cellRef(r, 14), `Performance review pending for ${fn} ${ln}`);
    }
  }

  // Sheet 3: Summary dashboard with formulas
  const summarySheet = 'Summary';
  wb.newSheet(summarySheet);

  wb.setCellValue(summarySheet, 'A1', 'Dashboard Summary');
  wb.mergeCells(summarySheet, 'A1', 'D1');
  wb.setCellStyle(summarySheet, 'A1', boldId);

  const metrics = [
    ['Total Invoices', '=COUNTA(Sheet1!A3:A3002)'],
    ['Total Revenue', '=SUM(Sheet1!E3:E3002)'],
    ['Average Invoice', '=AVERAGE(Sheet1!E3:E3002)'],
    ['Max Invoice', '=MAX(Sheet1!E3:E3002)'],
    ['Employee Count', '=COUNTA(Employees!A2:A2001)'],
    ['Avg Salary', '=AVERAGE(Employees!G2:G2001)'],
    ['Total Payroll', '=SUM(Employees!I2:I2001)'],
  ];

  for (let i = 0; i < metrics.length; i++) {
    wb.setCellValue(summarySheet, cellRef(i + 3, 0), metrics[i][0]);
    wb.setCellFormula(summarySheet, cellRef(i + 3, 1), metrics[i][1]);
  }

  wb.saveSync(join(FIXTURES_DIR, 'mixed-workload.xlsx'));
  console.log('  Done.');
}

function generateScaleFixture(rows: number, label: string) {
  console.log(`Generating scale-${label}.xlsx (${rows} rows x 10 cols)...`);
  const wb = new Workbook();
  const sheet = 'Sheet1';

  for (let c = 0; c < 10; c++) {
    wb.setCellValue(sheet, cellRef(1, c), `Col_${c + 1}`);
  }

  for (let r = 2; r <= rows + 1; r++) {
    for (let c = 0; c < 10; c++) {
      const mod = c % 3;
      if (mod === 0) wb.setCellValue(sheet, cellRef(r, c), r * (c + 1));
      else if (mod === 1) wb.setCellValue(sheet, cellRef(r, c), `R${r}C${c}`);
      else wb.setCellValue(sheet, cellRef(r, c), (r * c) / 100);
    }
  }

  wb.saveSync(join(FIXTURES_DIR, `scale-${label}.xlsx`));
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
generateDataValidation();
generateComments();
generateMergedCells();
generateMixedWorkload();
generateScaleFixture(1_000, '1k');
generateScaleFixture(10_000, '10k');
generateScaleFixture(100_000, '100k');

const elapsed = ((performance.now() - start) / 1000).toFixed(1);
console.log(`\nAll fixtures generated in ${elapsed}s`);

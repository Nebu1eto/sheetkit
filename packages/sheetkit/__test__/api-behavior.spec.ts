import { existsSync, unlinkSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { Workbook } from '../index.js';

const TEST_DIR = import.meta.dirname;

function tmpFile(name: string) {
  return join(TEST_DIR, name);
}

function cleanup(...files: string[]) {
  for (const f of files) {
    if (existsSync(f)) unlinkSync(f);
  }
}

describe('Workbook API behavior', () => {
  const out = tmpFile('test-api-behavior.xlsx');
  afterEach(() => cleanup(out));

  it('should support openSync/saveSync', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'sync');
    wb.saveSync(out);
    const wb2 = Workbook.openSync(out);
    expect(wb2.getCellValue('Sheet1', 'A1')).toBe('sync');
  });

  it('should throw on unknown chart type', () => {
    const wb = new Workbook();
    expect(() =>
      wb.addChart('Sheet1', 'A1', 'D10', {
        chartType: 'unknown-chart',
        series: [{ name: 'S1', categories: 'Sheet1!$A$1:$A$1', values: 'Sheet1!$B$1:$B$1' }],
      }),
    ).toThrow();
  });

  it('should throw on unknown validation type', () => {
    const wb = new Workbook();
    expect(() =>
      wb.addDataValidation('Sheet1', {
        sqref: 'A1:A2',
        validationType: 'unknown-validation',
      }),
    ).toThrow();
  });

  it('should throw on unknown pivot aggregate function', () => {
    const wb = new Workbook();
    expect(() =>
      wb.addPivotTable({
        name: 'PT1',
        sourceSheet: 'Sheet1',
        sourceRange: 'A1:B2',
        targetSheet: 'Sheet1',
        targetCell: 'D1',
        rows: [],
        columns: [],
        data: [{ name: 'Amount', function: 'unknown-fn' }],
      }),
    ).toThrow();
  });
});

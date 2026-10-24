import { describe, expect, it } from 'vitest';
import { type FormulaLimits, type FormulaValue, isResourceLimitError, Workbook } from '../index.js';
import { captureError } from './helpers/errors.js';

function formula(value: string): FormulaValue {
  return { type: 'formula', formula: value };
}

function expectFormulaLimit(
  error: Error,
  dimension: 'input_bytes' | 'output_bytes' | 'range_cell_count',
  limit: bigint,
  observed: bigint,
  stage: 'validate' | 'serialize',
): void {
  expect(isResourceLimitError(error)).toBe(true);

  if (!isResourceLimitError(error)) throw new Error('Expected a resource limit error');
  expect(error.operation).toBe('formula_mutate');
  expect(error.category).toBe('formula');
  expect(error.dimension).toBe(dimension);
  expect(error.limit).toBe(limit);
  expect(error.observed).toBe(observed);
  expect(error.stage).toBe(stage);
}

describe('formula mutation limits', () => {
  it('rejects a formula over the secure input default without mutation', () => {
    const workbook = new Workbook();
    workbook.setCellValue('Sheet1', 'A1', 'original');

    const error = captureError(() => workbook.setCellFormula('Sheet1', 'A1', 'x'.repeat(32_769)));

    expectFormulaLimit(error, 'input_bytes', 32_768n, 32_769n, 'validate');
    expect(workbook.getCellValue('Sheet1', 'A1')).toBe('original');
  });

  it('accepts an opaque Excel formula at exact caller-selected limits', () => {
    const workbook = new Workbook();
    const value = '_xlfn.UNSUPPORTED(A1)';

    const limits: FormulaLimits = {
      perFormulaInputBytes: value.length,
      operationInputBytes: value.length,
      operationOutputBytes: value.length,
    };

    workbook.setCellFormula('Sheet1', 'A1', value, limits);

    expect(workbook.getCellValue('Sheet1', 'A1')).toEqual({ type: 'formula', formula: value });
  });

  it('enforces fill range exact and one-over limits atomically', () => {
    const accepted = new Workbook();
    accepted.fillFormula('Sheet1', 'D1:D2', 'A1', { operationRangeCells: 2 });
    expect(accepted.getCellValue('Sheet1', 'D2')).toEqual({ type: 'formula', formula: 'A2' });

    const rejected = new Workbook();
    rejected.setCellValue('Sheet1', 'D1', 'first');
    rejected.setCellValue('Sheet1', 'D2', 'second');

    const error = captureError(() =>
      rejected.fillFormula('Sheet1', 'D1:D2', 'A1', { operationRangeCells: 1 }),
    );

    expectFormulaLimit(error, 'range_cell_count', 1n, 2n, 'validate');
    expect(rejected.getCellValue('Sheet1', 'D1')).toBe('first');
    expect(rejected.getCellValue('Sheet1', 'D2')).toBe('second');
  });

  it('preflights cumulative shifted output before mutation', () => {
    const accepted = new Workbook();
    accepted.fillFormula('Sheet1', 'D1:D2', 'A9', { operationOutputBytes: 5 });
    expect(accepted.getCellValue('Sheet1', 'D2')).toEqual({ type: 'formula', formula: 'A10' });

    const rejected = new Workbook();
    rejected.setCellValue('Sheet1', 'D1', 'first');
    rejected.setCellValue('Sheet1', 'D2', 'second');

    const error = captureError(() =>
      rejected.fillFormula('Sheet1', 'D1:D2', 'A9', { operationOutputBytes: 4 }),
    );

    expectFormulaLimit(error, 'output_bytes', 4n, 5n, 'serialize');
    expect(rejected.getCellValue('Sheet1', 'D1')).toBe('first');
    expect(rejected.getCellValue('Sheet1', 'D2')).toBe('second');
  });

  it('preflights cumulative shift work before mutation', () => {
    const workbook = new Workbook();
    workbook.setCellValue('Sheet1', 'D1', 'first');
    workbook.setCellValue('Sheet1', 'D2', 'second');
    workbook.setCellValue('Sheet1', 'D3', 'third');

    const error = captureError(() =>
      workbook.fillFormula('Sheet1', 'D1:D3', 'A1', { operationInputBytes: 5 }),
    );

    expectFormulaLimit(error, 'input_bytes', 5n, 14n, 'validate');
    expect(workbook.getCellValue('Sheet1', 'D1')).toBe('first');
    expect(workbook.getCellValue('Sheet1', 'D3')).toBe('third');
  });

  it.each([
    [
      'setCellValue',
      (workbook: Workbook, oversized: FormulaValue) =>
        workbook.setCellValue('Sheet1', 'A1', oversized),
    ],
    [
      'setCellValues',
      (workbook: Workbook, oversized: FormulaValue) =>
        workbook.setCellValues('Sheet1', [
          { cell: 'A1', value: formula('1+1') },
          { cell: 'B1', value: oversized },
        ]),
    ],
    [
      'setRowValues',
      (workbook: Workbook, oversized: FormulaValue) =>
        workbook.setRowValues('Sheet1', 1, 'A', [formula('1+1'), oversized]),
    ],
    [
      'setSheetData',
      (workbook: Workbook, oversized: FormulaValue) =>
        workbook.setSheetData('Sheet1', [[formula('1+1'), oversized]], 'A1'),
    ],
  ])('secures the %s FormulaValue adapter without partial mutation', (_name, mutate) => {
    const workbook = new Workbook();
    workbook.setCellValue('Sheet1', 'A1', 'first');
    workbook.setCellValue('Sheet1', 'B1', 'second');
    const oversized = formula('x'.repeat(32_769));

    const error = captureError(() => mutate(workbook, oversized));

    expectFormulaLimit(error, 'input_bytes', 32_768n, 32_769n, 'validate');
    expect(workbook.getCellValue('Sheet1', 'A1')).toBe('first');
    expect(workbook.getCellValue('Sheet1', 'B1')).toBe('second');
  });

  it('bounds cached formula results before a batch mutation', () => {
    const accepted = new Workbook();
    accepted.setCellValues('Sheet1', [
      {
        cell: 'A1',
        value: {
          type: 'formula',
          formula: '1+1',
          result: { valueType: 'string', value: 'x'.repeat(32_767) },
        },
      },
    ]);
    expect(accepted.getCellValue('Sheet1', 'A1')).toEqual({
      type: 'formula',
      formula: '1+1',
      result: { valueType: 'string', value: 'x'.repeat(32_767) },
    });

    const rejected = new Workbook();
    rejected.setCellValue('Sheet1', 'A1', 'first');
    rejected.setCellValue('Sheet1', 'B1', 'second');
    expect(() =>
      rejected.setCellValues('Sheet1', [
        {
          cell: 'A1',
          value: {
            type: 'formula',
            formula: '1+1',
            result: { valueType: 'string', value: 'ok' },
          },
        },
        {
          cell: 'B1',
          value: {
            type: 'formula',
            formula: '2+2',
            result: { valueType: 'error', value: 'x'.repeat(32_768) },
          },
        },
      ]),
    ).toThrow(/cell value too long/i);
    expect(rejected.getCellValue('Sheet1', 'A1')).toBe('first');
    expect(rejected.getCellValue('Sheet1', 'B1')).toBe('second');
  });

  it('enforces the cumulative cached-result output budget atomically', () => {
    const cachedText = 'x'.repeat(32_767);

    const exactValues: FormulaValue[] = Array.from({ length: 512 }, () => ({
      type: 'formula',
      formula: 'x',
      result: { valueType: 'string', value: cachedText },
    }));

    const accepted = new Workbook();
    accepted.setRowValues('Sheet1', 1, 'A', exactValues);
    expect(accepted.getCellValue('Sheet1', 'A1')).toEqual({
      type: 'formula',
      formula: 'x',
      result: { valueType: 'string', value: cachedText },
    });

    const rejected = new Workbook();
    rejected.setCellValue('Sheet1', 'A1', 'first');
    rejected.setCellValue('Sheet1', 'B1', 'second');

    const error = captureError(() =>
      rejected.setRowValues('Sheet1', 1, 'A', [...exactValues, formula('x')]),
    );

    expectFormulaLimit(error, 'output_bytes', 16_777_216n, 16_777_217n, 'serialize');
    expect(rejected.getCellValue('Sheet1', 'A1')).toBe('first');
    expect(rejected.getCellValue('Sheet1', 'B1')).toBe('second');
  });
});

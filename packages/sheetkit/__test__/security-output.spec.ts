import { describe, expect, it } from 'vitest';
import {
  isResourceLimitError,
  MAX_CSV_OUTPUT_BYTES,
  type ToCsvOptions,
  Workbook,
} from '../index.js';
import { withForeignField } from './helpers/contract-violation.js';
import { captureError } from './helpers/errors.js';

describe('CSV formula policy', () => {
  it('exports formula expressions with exact Rust and CLI record bytes', () => {
    const workbook = new Workbook();
    workbook.setSheetData('Sheet1', [
      [
        {
          type: 'formula',
          formula: '1+1',
          result: { valueType: 'number', numberValue: 2 },
        },
      ],
      [{ type: 'formula', formula: '2+2' }],
    ]);

    expect(workbook.toCSV('Sheet1')).toBe("'=1+1\n'=2+2");
    expect(workbook.toCSV('Sheet1', { formulaPolicy: 'raw-exchange' })).toBe('=1+1\n=2+2');
  });

  it('uses spreadsheet-safe output by default without changing numbers', () => {
    const workbook = new Workbook();
    workbook.setSheetData('Sheet1', [['=cmd', ' \t\uFEFF@SUM(A1)', '+1', '-1', 42, -42, 'plain']]);

    expect(workbook.toCSV('Sheet1')).toBe("'=cmd,' \t\uFEFF@SUM(A1),'+1,'-1,42,-42,plain");
  });

  it('requires the named raw-exchange policy to preserve formula-like text', () => {
    const workbook = new Workbook();
    workbook.setSheetData('Sheet1', [['=cmd', '+1', '-1', '@SUM(A1)', 'plain']]);

    expect(workbook.toCSV('Sheet1', { formulaPolicy: 'raw-exchange' })).toBe(
      '=cmd,+1,-1,@SUM(A1),plain',
    );
  });

  it('rejects unknown formula policies instead of treating them as raw exchange', () => {
    const workbook = new Workbook();
    workbook.setSheetData('Sheet1', [['=cmd']]);

    expect(() =>
      workbook.toCSV('Sheet1', withForeignField<ToCsvOptions>({}, 'formulaPolicy', 'unsafe')),
    ).toThrow(/formulaPolicy/);
  });

  it('rejects a quote that would become a formula trigger before neutralized text', () => {
    const workbook = new Workbook();
    workbook.setSheetData('Sheet1', [['=cmd,still-one-field']]);

    expect(() => workbook.toCSV('Sheet1', { quote: '=' })).toThrow(/syntax/);
  });

  it('rejects delimiters and line endings that can introduce formula prefixes', () => {
    const workbook = new Workbook();
    workbook.setSheetData('Sheet1', [[null, '=cmd'], ['plain']]);

    expect(() => workbook.toCSV('Sheet1', { delimiter: '=' })).toThrow(/syntax/);
    expect(() => workbook.toCSV('Sheet1', { lineEnding: '\n=' })).toThrow(/lineEnding/);
  });

  it('rejects ambiguous or control-based CSV structure', () => {
    const workbook = new Workbook();
    workbook.setSheetData('Sheet1', [['plain']]);

    expect(() => workbook.toCSV('Sheet1', { delimiter: '||' })).toThrow(/one character/);
    expect(() => workbook.toCSV('Sheet1', { delimiter: '\t' })).toThrow(/non-control/);
    expect(() => workbook.toCSV('Sheet1', { delimiter: ';', quote: ';' })).toThrow(/distinct/);
    expect(() => workbook.toCSV('Sheet1', { quote: '\r' })).toThrow(/non-control/);
    expect(() => workbook.toCSV('Sheet1', { lineEnding: '\r' })).toThrow(/lineEnding/);
  });

  it('preflights exact escaped UTF-8 output and rejects cumulative plus one', () => {
    const workbook = new Workbook();
    const quoteHeavy = `한글${'"'.repeat(1_000)}`;
    workbook.setSheetData('Sheet1', [[quoteHeavy], [quoteHeavy]]);
    const expected = workbook.toCSV('Sheet1');
    const exactBytes = Buffer.byteLength(expected);

    expect(workbook.toCSV('Sheet1', { limits: { maxOutputBytes: exactBytes } })).toBe(expected);

    const error = captureError(() =>
      workbook.toCSV('Sheet1', { limits: { maxOutputBytes: exactBytes - 1 } }),
    );

    expect(isResourceLimitError(error)).toBe(true);

    if (!isResourceLimitError(error)) throw new Error('Expected a resource limit error');
    expect(error.operation).toBe('export');
    expect(error.category).toBe('output');
    expect(error.dimension).toBe('output_bytes');
    expect(error.stage).toBe('serialize');
    expect(error.limit).toBe(BigInt(exactBytes - 1));
    expect(error.observed).toBe(BigInt(exactBytes));
    expect(MAX_CSV_OUTPUT_BYTES).toBe(64 * 1024 * 1024);
  });
});

import { describe, expect, it } from 'vitest';
import { isResourceLimitError, MAX_CSV_OUTPUT_BYTES, SheetData, Workbook } from '../index.js';
import { captureError } from './helpers/errors.js';

const sparseLimits = {
  maxRows: 2,
  maxColumns: 1,
  maxCells: 2,
  maxOutputBytes: 1_024,
};

function sparseExtremeWorkbook(): Workbook {
  const workbook = new Workbook();
  workbook.setCellValue('Sheet1', 'A1', 'first');
  workbook.setCellValue('Sheet1', 'A1048576', 'last');

  return workbook;
}

function expectTraversalLimit(error: Error): void {
  expect(isResourceLimitError(error)).toBe(true);

  if (!isResourceLimitError(error)) throw new Error('Expected a resource limit error');
  expect(error.operation).toBe('data_traverse');
  expect(error.category).toBe('traversal');
  expect(error.dimension).toBe('row_count');
  expect(error.limit).toBe(1n);
  expect(error.observed).toBe(2n);
  expect(error.stage).toBe('serialize');
}

describe('secure worksheet traversal', () => {
  it('keeps sparse output independent from the distance between occupied rows', () => {
    const workbook = sparseExtremeWorkbook();
    const options = { mode: 'sparse' as const, limits: sparseLimits };
    const first = workbook.getRowsBufferV2('Sheet1', options);
    const second = workbook.getRowsBufferV2('Sheet1', options);
    const data = new SheetData(first);

    expect(first.equals(second)).toBe(true);
    expect(data.rowCount).toBe(2);
    expect(data.minRow).toBe(1);
    expect(data.maxRow).toBe(1_048_576);
    expect(data.rowSpan).toBe(1_048_576);
    expect([...data.rows()].map(({ row }) => row)).toEqual([1, 1_048_576]);
    expect(workbook.getRows('Sheet1', options).map(({ row }) => row)).toEqual([1, 1_048_576]);
    expect(workbook.getRowsRaw('Sheet1', options).rowNumbers).toEqual(
      new Uint32Array([1, 1_048_576]),
    );
    expect([...workbook.getRowsIterator('Sheet1', options)].map(({ row }) => row)).toEqual([
      1, 1_048_576,
    ]);
  });

  it('uses sparse output by default for JSON, CSV, and HTML', () => {
    const workbook = sparseExtremeWorkbook();

    expect(workbook.toJSON('Sheet1', { header: false })).toHaveLength(2);
    expect(workbook.toCSV('Sheet1').split('\n')).toEqual(['first', 'last']);
    expect(workbook.toHTML('Sheet1').match(/<tr>/g)).toHaveLength(2);
  });

  it('keeps sparse exports proportional across extreme column gaps', () => {
    const workbook = new Workbook();
    const rowCount = 32;

    for (let row = 1; row <= rowCount; row++) {
      workbook.setCellValue('Sheet1', `A${row}`, `left-${row}`);
      workbook.setCellValue('Sheet1', `XFD${row}`, `right-${row}`);
    }

    const traversal = {
      mode: 'sparse' as const,
      limits: {
        maxRows: rowCount,
        maxColumns: 16_384,
        maxCells: rowCount * 2,
        maxOutputBytes: 1_048_576,
      },
    };

    const data = new SheetData(workbook.getRowsBufferV2('Sheet1', traversal));
    const v1Data = new SheetData(workbook.getRowsBuffer('Sheet1', traversal));

    expect(data.getRow(1)).toEqual(['left-1', 'right-1']);
    expect(v1Data.getRow(1)).toEqual(['left-1', 'right-1']);
    expect(data.toArray()).toHaveLength(rowCount);
    expect([...data.occupiedRows()][0]).toEqual({
      row: 1,
      cells: [
        { column: 1, value: 'left-1' },
        { column: 16_384, value: 'right-1' },
      ],
    });

    const json = workbook.toJSON('Sheet1', { header: false, ...traversal });
    expect(json).toHaveLength(rowCount);
    expect(json[0]).toEqual({ A: 'left-1', XFD: 'right-1' });
    expect(json.every((row) => Object.keys(row).length === 2)).toBe(true);
    expect(workbook.toJSON('Sheet1', traversal)[0]).toEqual({
      'left-1': 'left-2',
      'right-1': 'right-2',
    });

    const csvLines = workbook.toCSV('Sheet1', traversal).split('\n');
    expect(csvLines).toHaveLength(rowCount);
    expect(csvLines[0]).toBe('left-1,right-1');
    expect(csvLines[rowCount - 1]).toBe(`left-${rowCount},right-${rowCount}`);

    const html = workbook.toHTML('Sheet1', traversal);
    expect(html.match(/<tr>/g)).toHaveLength(rowCount);
    expect(html.match(/<td>/g)).toHaveLength(rowCount * 2);
    expect(html).not.toContain('<td></td>');
  });

  it('retains rectangular compatibility only for admitted dense exports', () => {
    const workbook = new Workbook();
    workbook.setCellValue('Sheet1', 'A1', 'left');
    workbook.setCellValue('Sheet1', 'C1', 'right');
    workbook.setCellValue('Sheet1', 'A2', 'next');

    const dense = {
      mode: 'dense' as const,
      limits: { maxRows: 2, maxColumns: 3, maxCells: 6, maxOutputBytes: 4_096 },
    };

    expect(workbook.toJSON('Sheet1', { header: false, ...dense })).toEqual([
      { A: 'left', B: null, C: 'right' },
      { A: 'next', B: null, C: null },
    ]);
    expect(workbook.toCSV('Sheet1', dense)).toBe('left,,right\nnext,,');
    expect(workbook.toHTML('Sheet1', dense)).toBe(
      '<table><tr><td>left</td><td></td><td>right</td></tr>' +
        '<tr><td>next</td><td></td><td></td></tr></table>',
    );
  });

  it('applies the CSV byte limit to compact output, not its transfer buffer', () => {
    const workbook = new Workbook();
    workbook.setCellValue('Sheet1', 'A1', 'stale'.repeat(20));
    workbook.setCellValue('Sheet1', 'A1', 'x');
    expect(workbook.toCSV('Sheet1', { limits: { maxOutputBytes: 1 } })).toBe('x');
    expect(
      workbook.toCSV('Sheet1', {
        mode: 'dense',
        limits: { maxRows: 1, maxColumns: 1, maxCells: 1, maxOutputBytes: 1 },
      }),
    ).toBe('x');

    workbook.setCellValue('Sheet1', 'A1', 'xx');
    const error = captureError(() => workbook.toCSV('Sheet1', { limits: { maxOutputBytes: 1 } }));
    expect(isResourceLimitError(error)).toBe(true);

    if (!isResourceLimitError(error)) throw new Error('Expected a resource limit error');
    expect(error.operation).toBe('export');
    expect(error.category).toBe('output');
    expect(error.dimension).toBe('output_bytes');
    expect(error.limit).toBe(1n);
    expect(error.observed).toBe(2n);
    expect(error.stage).toBe('serialize');
  });

  it('derives a finite transfer ceiling from a tiny final CSV limit', () => {
    const workbook = new Workbook();
    const large = 'x'.repeat(32_767);
    workbook.setSheetData(
      'Sheet1',
      Array.from({ length: 64 }, () => [large]),
    );

    const error = captureError(() => workbook.toCSV('Sheet1', { limits: { maxOutputBytes: 1 } }));

    expect(isResourceLimitError(error)).toBe(true);

    if (!isResourceLimitError(error)) throw new Error('Expected a resource limit error');
    expect(error.operation).toBe('data_traverse');
    expect(error.category).toBe('traversal');
    expect(error.dimension).toBe('output_bytes');
    expect(error.limit).toBe(85n);
    expect(error.observed).toBeGreaterThan(error.limit);
  });

  it('accepts valid CSV output near the public cap despite transfer framing', () => {
    const workbook = new Workbook();
    const value = 'x'.repeat(32_767);
    const rows = Array.from({ length: 2_048 }, () => [value]);
    workbook.setSheetData('Sheet1', rows);
    const expectedBytes = rows.length * value.length + rows.length - 1;

    const csv = workbook.toCSV('Sheet1', {
      limits: { maxOutputBytes: MAX_CSV_OUTPUT_BYTES },
    });

    expect(Buffer.byteLength(csv)).toBe(expectedBytes);
    expect(expectedBytes).toBe(MAX_CSV_OUTPUT_BYTES - 1);
  });

  it('preserves typed exact and plus-one errors across all public paths', () => {
    const workbook = sparseExtremeWorkbook();
    const exact = { mode: 'sparse' as const, limits: sparseLimits };
    const rejected = { ...exact, limits: { ...sparseLimits, maxRows: 1 } };

    expect(workbook.getRowsBuffer('Sheet1', exact).length).toBeGreaterThan(0);
    expect(workbook.getRowsBufferV2('Sheet1', exact).length).toBeGreaterThan(0);

    const actions = [
      () => workbook.getRowsBuffer('Sheet1', rejected),
      () => workbook.getRowsBufferV2('Sheet1', rejected),
      () => workbook.getRows('Sheet1', rejected),
      () => workbook.getRowsRaw('Sheet1', rejected),
      () => [...workbook.getRowsIterator('Sheet1', rejected)],
      () => workbook.getCols('Sheet1', rejected),
      () => workbook.toJSON('Sheet1', { header: false, ...rejected }),
      () => workbook.toCSV('Sheet1', rejected),
      () => workbook.toHTML('Sheet1', rejected),
    ];

    for (const action of actions) expectTraversalLimit(captureError(action));

    const hardCeiling = captureError(() =>
      workbook.getRowsBufferV2('Sheet1', {
        limits: { maxCells: 100_000_001 },
      }),
    );

    expect(isResourceLimitError(hardCeiling)).toBe(true);

    if (!isResourceLimitError(hardCeiling)) throw new Error('Expected a resource limit error');
    expect(hardCeiling.dimension).toBe('cell_count');
    expect(hardCeiling.limit).toBe(100_000_000n);
    expect(hardCeiling.observed).toBe(100_000_001n);

    expect(() =>
      workbook.getRowsBufferV2('Sheet1', { limits: { maxOutputBytes: Number.POSITIVE_INFINITY } }),
    ).toThrow('finite');
  });

  it('allows finite dense v1 compatibility and rejects dense v2', () => {
    const workbook = new Workbook();
    workbook.setCellValue('Sheet1', 'A1', 'first');
    workbook.setCellValue('Sheet1', 'A3', 'last');

    const dense = {
      mode: 'dense' as const,
      limits: { maxRows: 3, maxColumns: 1, maxCells: 3, maxOutputBytes: 1_024 },
    };

    const data = new SheetData(workbook.getRowsBuffer('Sheet1', dense));
    expect(data.rowCount).toBe(2);
    expect(data.toArray()).toEqual([['first'], [null], ['last']]);
    expect(workbook.toCSV('Sheet1', dense)).toBe('first\n\nlast');
    expect(() => workbook.getRowsBufferV2('Sheet1', dense)).toThrow(
      'raw transfer v2 supports sparse traversal only',
    );
    expect(() => workbook.getCols('Sheet1', dense)).toThrow(
      'getCols supports sparse traversal only',
    );
  });

  it('retains safe own properties in traversal-aware JSON output', () => {
    const workbook = new Workbook();
    workbook.setCellValue('Sheet1', 'A1', '__proto__');
    workbook.setCellValue('Sheet1', 'A2', 'value');

    const [row] = workbook.toJSON('Sheet1', { mode: 'sparse' });
    expect(Object.hasOwn(row, '__proto__')).toBe(true);
    expect(Object.getPrototypeOf(row)).toBe(Object.prototype);
  });
});

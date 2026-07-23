import { describe, expect, it } from 'vitest';
import { isResourceLimitError, Workbook } from '../index.js';

function captureError(action: () => unknown): unknown {
  try {
    action();
  } catch (error) {
    return error;
  }
  throw new Error('Expected action to throw');
}

function expectResourceError(
  error: unknown,
  dimension: 'row_count' | 'column_count' | 'cell_count' | 'retained_bytes',
  limit: bigint,
  observed: bigint,
): void {
  expect(isResourceLimitError(error)).toBe(true);
  if (!isResourceLimitError(error)) throw new Error('Expected a resource limit error');
  expect(error.operation).toBe('data_traverse');
  expect(error.category).toBe('structure');
  expect(error.dimension).toBe(dimension);
  expect(error.limit).toBe(limit);
  expect(error.observed).toBe(observed);
  expect(error.stage).toBe('validate');
}

describe('fromJSON finite import boundary', () => {
  it('accepts exact row, column, and sparse cell limits', () => {
    const workbook = new Workbook();
    workbook.fromJSON('Sheet1', [{ A: 1, B: 2 }, { A: 3 }], {
      header: false,
      limits: {
        maxRows: 2,
        maxColumns: 2,
        maxCells: 3,
        maxRetainedBytes: 1_024,
      },
    });

    expect(workbook.getCellValue('Sheet1', 'A1')).toBe(1);
    expect(workbook.getCellValue('Sheet1', 'B1')).toBe(2);
    expect(workbook.getCellValue('Sheet1', 'A2')).toBe(3);
    expect(workbook.getCellValue('Sheet1', 'B2')).toBeNull();
  });

  it('rejects one-over row, column, and sparse cell limits before mutation', () => {
    const cases = [
      {
        data: [{ A: 1 }, { A: 2 }],
        limits: { maxRows: 1 },
        dimension: 'row_count' as const,
        limit: 1n,
        observed: 2n,
      },
      {
        data: [{ A: 1, B: 2 }],
        limits: { maxColumns: 1 },
        dimension: 'column_count' as const,
        limit: 1n,
        observed: 2n,
      },
      {
        data: [{ A: 1, B: 2 }],
        limits: { maxCells: 1 },
        dimension: 'cell_count' as const,
        limit: 1n,
        observed: 2n,
      },
    ];

    for (const testCase of cases) {
      const workbook = new Workbook();
      workbook.setCellValue('Sheet1', 'A1', 'sentinel');
      const error = captureError(() =>
        workbook.fromJSON('Sheet1', testCase.data, {
          header: false,
          limits: { ...testCase.limits, maxRetainedBytes: 1_024 },
        }),
      );
      expectResourceError(error, testCase.dimension, testCase.limit, testCase.observed);
      expect(workbook.getCellValue('Sheet1', 'A1')).toBe('sentinel');
      expect(workbook.getCellValue('Sheet1', 'B1')).toBeNull();
    }
  });

  it('rejects hard Excel dimensions before iterating sparse arrays', () => {
    const rows = new Array(1_048_577) as Record<string, never>[];
    const rowError = captureError(() => new Workbook().fromJSON('Sheet1', rows, { header: false }));
    expectResourceError(rowError, 'row_count', 1_048_576n, 1_048_577n);

    const columns = new Array(16_385) as string[];
    const columnError = captureError(() =>
      new Workbook().fromJSON('Sheet1', [{}], {
        header: columns,
      }),
    );
    expectResourceError(columnError, 'column_count', 16_384n, 16_385n);
  });

  it('ignores caller-controlled array iterators and reads bounded indexed data properties', () => {
    const data = [{ A: 1 }];
    Object.defineProperty(data, Symbol.iterator, {
      value: function* infiniteRows() {
        while (true) yield { A: 1 };
      },
    });
    const headers = ['A'];
    Object.defineProperty(headers, Symbol.iterator, {
      value: function* infiniteHeaders() {
        let index = 0;
        while (true) yield `H${index++}`;
      },
    });

    const workbook = new Workbook();
    workbook.fromJSON('Sheet1', data, { header: headers });
    expect(workbook.getCellValue('Sheet1', 'A1')).toBe('A');
    expect(workbook.getCellValue('Sheet1', 'A2')).toBe(1);
    expect(workbook.getCellValue('Sheet1', 'B1')).toBeNull();
  });

  it('counts header cells and keeps finite overrides local to one call', () => {
    const workbook = new Workbook();
    const error = captureError(() =>
      workbook.fromJSON('Sheet1', [{ A: 1 }], {
        limits: { maxCells: 1 },
      }),
    );
    expectResourceError(error, 'cell_count', 1n, 2n);
    expect(workbook.getCellValue('Sheet1', 'A1')).toBeNull();

    workbook.fromJSON('Sheet1', [{ A: 1 }]);
    expect(workbook.getCellValue('Sheet1', 'A1')).toBe('A');
    expect(workbook.getCellValue('Sheet1', 'A2')).toBe(1);
  });

  it('applies retained-byte limits cumulatively and accepts the exact boundary', () => {
    const exact = new Workbook();
    exact.fromJSON('Sheet1', [{ A: 'x' }], {
      header: false,
      limits: {
        maxRows: 1,
        maxColumns: 1,
        maxCells: 1,
        maxRetainedBytes: 84,
      },
    });
    expect(exact.getCellValue('Sheet1', 'A1')).toBe('x');

    const rejected = new Workbook();
    rejected.setCellValue('Sheet1', 'A1', 'sentinel');
    const error = captureError(() =>
      rejected.fromJSON('Sheet1', [{ A: 'x' }], {
        header: false,
        limits: {
          maxRows: 1,
          maxColumns: 1,
          maxCells: 1,
          maxRetainedBytes: 83,
        },
      }),
    );
    expectResourceError(error, 'retained_bytes', 83n, 84n);
    expect(rejected.getCellValue('Sheet1', 'A1')).toBe('sentinel');

    const cumulative = captureError(() =>
      new Workbook().fromJSON('Sheet1', [{ A: 'x' }, { A: 'y' }], {
        header: false,
        limits: {
          maxRows: 2,
          maxColumns: 1,
          maxCells: 2,
          maxRetainedBytes: 150,
        },
      }),
    );
    expect(isResourceLimitError(cumulative)).toBe(true);
    if (!isResourceLimitError(cumulative)) throw new Error('Expected a resource limit error');
    expect(cumulative.dimension).toBe('retained_bytes');
    expect(cumulative.limit).toBe(150n);
    expect(cumulative.observed).toBeGreaterThan(150n);
  });

  it('allows one sparse cell at the last Excel coordinate', () => {
    const workbook = new Workbook();
    workbook.fromJSON('Sheet1', [{ Value: 42 }], {
      header: false,
      startCell: 'XFD1048576',
      limits: {
        maxRows: 1,
        maxColumns: 1,
        maxCells: 1,
        maxRetainedBytes: 1_024,
      },
    });
    expect(workbook.getCellValue('Sheet1', 'XFD1048576')).toBe(42);
  });

  it('rejects a staged rectangle that crosses the Excel boundary', () => {
    const workbook = new Workbook();
    workbook.setCellValue('Sheet1', 'XFD1048576', 'sentinel');

    const columnError = captureError(() =>
      workbook.fromJSON('Sheet1', [{ A: 1, B: 2 }], {
        header: false,
        startCell: 'XFD1048576',
      }),
    );
    expectResourceError(columnError, 'column_count', 16_384n, 16_385n);

    const rowError = captureError(() =>
      workbook.fromJSON('Sheet1', [{ A: 1 }, { A: 2 }], {
        header: false,
        startCell: 'XFD1048576',
      }),
    );
    expectResourceError(rowError, 'row_count', 1_048_576n, 1_048_577n);
    expect(workbook.getCellValue('Sheet1', 'XFD1048576')).toBe('sentinel');
  });

  it('rejects invalid and out-of-bounds references without echoing them', () => {
    for (const startCell of ['A0', 'XFE1', 'A1048577', 'A1:Z9', 'AAAA999999999999999999']) {
      const workbook = new Workbook();
      workbook.setCellValue('Sheet1', 'A1', 'sentinel');
      const error = captureError(() =>
        workbook.fromJSON('Sheet1', [{ A: 1 }], {
          header: false,
          startCell,
        }),
      );
      expect(String(error)).not.toContain(startCell);
      expect(workbook.getCellValue('Sheet1', 'A1')).toBe('sentinel');
    }
  });

  it('rejects duplicate custom headers before mutation', () => {
    const workbook = new Workbook();
    workbook.setCellValue('Sheet1', 'A1', 'sentinel');
    expect(() =>
      workbook.fromJSON('Sheet1', [{ A: 1 }], {
        header: ['A', 'A'],
      }),
    ).toThrow('header names must be unique');
    expect(workbook.getCellValue('Sheet1', 'A1')).toBe('sentinel');
  });

  it('rejects unmatched custom-header keys on the first inspected property', () => {
    const workbook = new Workbook();
    workbook.setCellValue('Sheet1', 'A1', 'sentinel');
    let matchedGetterCalled = false;
    const record: Record<string, unknown> = { Unexpected: 1 };
    Object.defineProperty(record, 'A', {
      enumerable: true,
      get() {
        matchedGetterCalled = true;
        return 2;
      },
    });

    expect(() =>
      workbook.fromJSON('Sheet1', [record as never], {
        header: ['A'],
      }),
    ).toThrow('must match custom headers');
    expect(matchedGetterCalled).toBe(false);
    expect(workbook.getCellValue('Sheet1', 'A1')).toBe('sentinel');
  });

  it('rejects custom prototypes without traversing inherited enumerable properties', () => {
    let prototypeReads = 0;
    const prototype = new Proxy(
      { inherited: 1 },
      {
        ownKeys(target) {
          prototypeReads++;
          return Reflect.ownKeys(target);
        },
        getOwnPropertyDescriptor(target, key) {
          prototypeReads++;
          return Reflect.getOwnPropertyDescriptor(target, key);
        },
      },
    );
    const record = Object.assign(Object.create(prototype) as Record<string, unknown>, { A: 1 });

    expect(() =>
      new Workbook().fromJSON('Sheet1', [record as never], {
        header: false,
        limits: {
          maxRows: 1,
          maxColumns: 1,
          maxCells: 1,
          maxRetainedBytes: 1_024,
        },
      }),
    ).toThrow('plain objects');
    expect(prototypeReads).toBe(0);
  });

  it('rejects shifted columns before inspecting later records', () => {
    let laterRecordReads = 0;
    const laterRecord = new Proxy(
      { A: 3 },
      {
        ownKeys(target) {
          laterRecordReads++;
          return Reflect.ownKeys(target);
        },
        getOwnPropertyDescriptor(target, key) {
          laterRecordReads++;
          return Reflect.getOwnPropertyDescriptor(target, key);
        },
      },
    );

    const error = captureError(() =>
      new Workbook().fromJSON('Sheet1', [{ A: 1, B: 2 }, laterRecord], {
        header: false,
        startCell: 'XFD1',
      }),
    );
    expectResourceError(error, 'column_count', 16_384n, 16_385n);
    expect(laterRecordReads).toBe(0);
  });

  it('imports special JSON property names as ordinary headers', () => {
    const workbook = new Workbook();
    const record = JSON.parse(
      '{"__proto__":"prototype","constructor":"constructor","prototype":"value"}',
    ) as Record<string, string>;
    workbook.fromJSON('Sheet1', [record]);

    expect(workbook.getCellValue('Sheet1', 'A1')).toBe('__proto__');
    expect(workbook.getCellValue('Sheet1', 'B1')).toBe('constructor');
    expect(workbook.getCellValue('Sheet1', 'C1')).toBe('prototype');
    expect(workbook.getCellValue('Sheet1', 'A2')).toBe('prototype');
    expect(workbook.getCellValue('Sheet1', 'B2')).toBe('constructor');
    expect(workbook.getCellValue('Sheet1', 'C2')).toBe('value');
    expect(Object.prototype).not.toHaveProperty('value');
  });

  it('rejects non-finite values atomically', () => {
    for (const value of [
      Number.NaN,
      Number.POSITIVE_INFINITY,
      { type: 'date', serial: Number.NEGATIVE_INFINITY, iso: null },
      {
        type: 'formula',
        formula: '1',
        result: { valueType: 'number', numberValue: Number.NaN },
      },
    ]) {
      const workbook = new Workbook();
      workbook.setCellValue('Sheet1', 'A1', 'sentinel');
      expect(() =>
        workbook.fromJSON('Sheet1', [{ A: 'accepted-first' }, { A: value as never }], {
          header: false,
        }),
      ).toThrow();
      expect(workbook.getCellValue('Sheet1', 'A1')).toBe('sentinel');
      expect(workbook.getCellValue('Sheet1', 'A2')).toBeNull();
    }
  });

  it('rejects accessors and ambiguous cached formula results without evaluating them', () => {
    const workbook = new Workbook();
    workbook.setCellValue('Sheet1', 'A1', 'sentinel');
    let getterCalled = false;
    const accessorRecord: Record<string, unknown> = {};
    Object.defineProperty(accessorRecord, 'A', {
      enumerable: true,
      get() {
        getterCalled = true;
        workbook.setCellValue('Sheet1', 'A1', 'side effect');
        return 1;
      },
    });

    expect(() => workbook.fromJSON('Sheet1', [accessorRecord as never], { header: false })).toThrow(
      'enumerable data properties',
    );
    expect(getterCalled).toBe(false);
    expect(workbook.getCellValue('Sheet1', 'A1')).toBe('sentinel');

    expect(() =>
      workbook.fromJSON(
        'Sheet1',
        [
          {
            A: {
              type: 'formula',
              formula: '1',
              result: { valueType: 'empty', value: 'unused' },
            } as never,
          },
        ],
        { header: false },
      ),
    ).toThrow('Invalid JSON formula result');
    expect(workbook.getCellValue('Sheet1', 'A1')).toBe('sentinel');
  });

  it('rejects stateful nested value accessors without re-reading them for N-API', () => {
    const workbook = new Workbook();
    workbook.setCellValue('Sheet1', 'A1', 'sentinel');
    let formulaGetterCalled = false;
    const formula = { type: 'formula' } as Record<string, unknown>;
    Object.defineProperty(formula, 'formula', {
      enumerable: true,
      get() {
        formulaGetterCalled = true;
        return formulaGetterCalled ? '1'.repeat(1_024) : '1';
      },
    });

    expect(() =>
      workbook.fromJSON('Sheet1', [{ A: formula as never }], {
        header: false,
        limits: { maxRetainedBytes: 100 },
      }),
    ).toThrow('data properties');
    expect(formulaGetterCalled).toBe(false);
    expect(workbook.getCellValue('Sheet1', 'A1')).toBe('sentinel');
  });

  it('rejects malformed runtime header and limit options', () => {
    const workbook = new Workbook();
    expect(() => workbook.fromJSON('Sheet1', [{ A: 1 }], { header: 0 as never })).toThrow(
      'header must be',
    );
    expect(() => workbook.fromJSON('Sheet1', [{ A: 1 }], { limits: 'unbounded' as never })).toThrow(
      'limits must be an object',
    );
    expect(() =>
      workbook.fromJSON('Sheet1', [{ A: 1 }], {
        limits: { maxCells: Number.POSITIVE_INFINITY },
      }),
    ).toThrow('safe integers');
    expect(workbook.getCellValue('Sheet1', 'A1')).toBeNull();
  });

  it('rejects oversized cell strings before the native call', () => {
    const workbook = new Workbook();
    workbook.setCellValue('Sheet1', 'A1', 'sentinel');
    expect(() =>
      workbook.fromJSON('Sheet1', [{ A: 'accepted-first' }, { A: 'x'.repeat(32_768) }], {
        header: false,
        limits: { maxRetainedBytes: 1_024 * 1_024 },
      }),
    ).toThrow('Excel limit');
    expect(workbook.getCellValue('Sheet1', 'A1')).toBe('sentinel');
    expect(workbook.getCellValue('Sheet1', 'A2')).toBeNull();
  });

  it('stages only present properties and still clears explicit nulls', () => {
    const workbook = new Workbook();
    workbook.setCellValue('Sheet1', 'B2', 'preserved');
    workbook.setCellValue('Sheet1', 'C2', 'cleared');

    workbook.fromJSON(
      'Sheet1',
      [
        { A: 1, B: 2, C: 3 },
        { A: 4, C: null },
      ],
      { header: false },
    );

    expect(workbook.getCellValue('Sheet1', 'A2')).toBe(4);
    expect(workbook.getCellValue('Sheet1', 'B2')).toBe('preserved');
    expect(workbook.getCellValue('Sheet1', 'C2')).toBeNull();
  });

  it('rejects an oversized formula through the native typed budget atomically', () => {
    const workbook = new Workbook();
    workbook.setCellValue('Sheet1', 'A1', 'sentinel');
    const error = captureError(() =>
      workbook.fromJSON(
        'Sheet1',
        [{ A: 'accepted-first' }, { A: { type: 'formula', formula: '1'.repeat(32_769) } }],
        {
          header: false,
          limits: { maxRetainedBytes: 1_024 * 1_024 },
        },
      ),
    );

    expect(isResourceLimitError(error)).toBe(true);
    if (!isResourceLimitError(error)) throw new Error('Expected a resource limit error');
    expect(error.category).toBe('formula');
    expect(error.dimension).toBe('input_bytes');
    expect(error.limit).toBe(32_768n);
    expect(error.observed).toBe(32_769n);
    expect(workbook.getCellValue('Sheet1', 'A1')).toBe('sentinel');
    expect(workbook.getCellValue('Sheet1', 'A2')).toBeNull();
  });

  it('preserves ordinary fromJSON and toJSON roundtrips', () => {
    const workbook = new Workbook();
    const data = [
      { Name: 'Alice', Score: 95, Pass: true },
      { Name: 'Bob', Score: 80, Pass: false },
    ];
    workbook.fromJSON('Sheet1', data);
    expect(workbook.toJSON('Sheet1')).toEqual(data);
  });
});

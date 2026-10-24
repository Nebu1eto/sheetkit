import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { type JsOpenOptions, Workbook as NativeWorkbook } from '../binding.js';
import { isResourceLimitError, Workbook } from '../index.js';
import { captureError, captureRejection } from './helpers/errors.js';

const encryptedFixture = readFileSync(
  join(
    dirname(fileURLToPath(import.meta.url)),
    '../../../crates/sheetkit-core/tests/fixtures/encryption/msoffcrypto-tool-v6.0.0/agile-password.xlsx',
  ),
);

function expectResourceError(
  error: Error,
  expected: {
    operation:
      | 'render'
      | 'data_traverse'
      | 'workbook_open'
      | 'decrypt'
      | 'stream_read'
      | 'formula_mutate'
      | 'workbook_save'
      | 'stream_write';
    category: 'traversal' | 'output' | 'archive' | 'container' | 'streaming' | 'formula';
    dimension: 'cell_count' | 'row_count' | 'output_bytes' | 'input_bytes' | 'xml_bytes';
    stage: 'render' | 'serialize' | 'write' | 'inspect' | 'parse' | 'validate';
    limit?: bigint;
    observed?: bigint;
  },
): void {
  expect(isResourceLimitError(error)).toBe(true);

  if (!isResourceLimitError(error)) throw new Error('Expected a resource limit error');
  const { limit, observed, ...classification } = expected;
  expect(error).toMatchObject(classification);
  expect(error.limit).toBe(limit ?? error.limit);
  expect(error.observed).toBe(observed ?? error.observed);
  expect(error.observed).toBeGreaterThan(error.limit);
}

describe('Node resource error parity', () => {
  it('keeps save and writeBuffer sync/async limits structured and path output atomic', async () => {
    const workbook = new Workbook();
    workbook.setCellValue('Sheet1', 'A1', 'payload');
    const limits = { outputBytes: 1 };

    const syncBufferError = captureError(() => workbook.writeBufferSync(limits));
    expectResourceError(syncBufferError, {
      operation: 'workbook_save',
      category: 'output',
      dimension: 'output_bytes',
      stage: 'write',
      limit: 1n,
      observed: 2n,
    });

    const asyncBufferError = await captureRejection(workbook.writeBuffer(limits));
    expectResourceError(asyncBufferError, {
      operation: 'workbook_save',
      category: 'output',
      dimension: 'output_bytes',
      stage: 'write',
      limit: 1n,
      observed: 2n,
    });

    const directory = mkdtempSync(join(tmpdir(), 'sheetkit-save-limit-'));
    const path = join(directory, 'atomic.xlsx');
    const sentinel = Buffer.from('existing destination');
    writeFileSync(path, sentinel);

    try {
      const syncSaveError = captureError(() => workbook.saveSync(path, limits));
      expectResourceError(syncSaveError, {
        operation: 'workbook_save',
        category: 'output',
        dimension: 'output_bytes',
        stage: 'write',
        limit: 1n,
        observed: 2n,
      });
      expect(readFileSync(path)).toEqual(sentinel);

      const asyncSaveError = await captureRejection(workbook.save(path, limits));
      expectResourceError(asyncSaveError, {
        operation: 'workbook_save',
        category: 'output',
        dimension: 'output_bytes',
        stage: 'write',
        limit: 1n,
        observed: 2n,
      });
      expect(readFileSync(path)).toEqual(sentinel);

      workbook.setCellValue('Sheet1', 'B1', 'task released');
      expect(workbook.getCellValue('Sheet1', 'B1')).toBe('task released');
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });

  it('maps StreamWriter row limits without publishing the rejected row', () => {
    const workbook = new Workbook();

    const writer = workbook.newStreamWriter('Bounded', {
      rowCount: 1,
      cellCount: 1,
      rowXmlBytes: 1_024,
      outputBytes: 1_024,
    });

    writer.writeRow(1, ['accepted']);

    const error = captureError(() => writer.writeRow(2, ['rejected']));
    expectResourceError(error, {
      operation: 'stream_write',
      category: 'streaming',
      dimension: 'row_count',
      stage: 'serialize',
      limit: 1n,
      observed: 2n,
    });

    workbook.applyStreamWriter(writer);
    const reopened = Workbook.openBufferSync(workbook.writeBufferSync());
    expect(reopened.getCellValue('Bounded', 'A1')).toBe('accepted');
    expect(reopened.getCellValue('Bounded', 'A2')).toBeNull();
  });

  it('preserves typed lazy hydration failures across cell and feature methods', () => {
    const source = new Workbook();

    for (let row = 1; row <= 512; row++) {
      source.setCellValue('Sheet1', `A${row}`, row);
    }

    const data = source.writeBufferSync();
    const options = { readMode: 'lazy' as const, maxXmlPartBytes: 4_096 };

    for (const access of [
      (workbook: Workbook) => workbook.getCellValue('Sheet1', 'A1'),
      (workbook: Workbook) => workbook.getDataValidations('Sheet1'),
    ]) {
      const workbook = Workbook.openBufferSync(data, options);
      expect(workbook.sheetNames).toEqual(['Sheet1']);
      const error = captureError(() => access(workbook));
      expectResourceError(error, {
        operation: 'stream_read',
        category: 'streaming',
        dimension: 'xml_bytes',
        stage: 'parse',
        limit: 4_096n,
        observed: 4_097n,
      });
    }

    const accepted = Workbook.openBufferSync(data, {
      readMode: 'lazy',
      maxXmlPartBytes: 64 * 1_024,
    });

    expect(accepted.getCellValue('Sheet1', 'A1')).toBe(1);
    expect(accepted.getDataValidations('Sheet1')).toEqual([]);
  });

  it('keeps ordinary core failures outside the resource-error contract', () => {
    const workbook = new Workbook();
    const error = captureError(() => workbook.getCellValue('Missing', 'A1'));
    expect(isResourceLimitError(error)).toBe(false);
    expect(String(error)).toMatch(/sheet.*(?:not found|does not exist)/i);
  });

  it('preserves typed formula limits for direct evaluation', () => {
    const workbook = new Workbook();
    expect(
      workbook.evaluateFormula('Sheet1', '1+1', {
        perFormulaInputBytes: 3,
        operationInputBytes: 3,
      }),
    ).toBe(2);

    const error = captureError(() =>
      workbook.evaluateFormula('Sheet1', '1+1', {
        perFormulaInputBytes: 2,
        operationInputBytes: 3,
      }),
    );

    expect(isResourceLimitError(error)).toBe(true);

    if (!isResourceLimitError(error)) throw new Error('Expected a resource limit error');
    expect(error.operation).toBe('formula_calculate');
    expect(error.category).toBe('formula');
    expect(error.dimension).toBe('input_bytes');
    expect(error.limit).toBe(2n);
    expect(error.observed).toBe(3n);
    expect(error.stage).toBe('parse');
  });

  it('preserves typed formula limits and workbook state for calculateAll', () => {
    const workbook = new Workbook();
    workbook.setCellFormula('Sheet1', 'A1', '1+1');
    const before = workbook.writeBufferSync();

    const error = captureError(() =>
      workbook.calculateAll({
        perFormulaInputBytes: 3,
        operationInputBytes: 2,
      }),
    );

    expect(isResourceLimitError(error)).toBe(true);

    if (!isResourceLimitError(error)) throw new Error('Expected a resource limit error');
    expect(error.operation).toBe('formula_calculate');
    expect(error.category).toBe('formula');
    expect(error.dimension).toBe('input_bytes');
    expect(error.limit).toBe(2n);
    expect(error.observed).toBe(3n);
    expect(error.stage).toBe('parse');
    expect(workbook.writeBufferSync()).toEqual(before);

    workbook.calculateAll({
      perFormulaInputBytes: 3,
      operationInputBytes: 3,
    });
    expect(workbook.writeBufferSync()).not.toEqual(before);
  });

  it('bounds native row and column collectors with the same options', () => {
    const native = new NativeWorkbook();
    native.setCellValue('Sheet1', 'A1', 'first');
    native.setCellValue('Sheet1', 'B2', 'second');

    const exact = {
      mode: 'sparse',
      limits: { maxRows: 2, maxColumns: 2, maxCells: 2, maxOutputBytes: 1_024 },
    };

    expect(native.getRows('Sheet1', exact)).toHaveLength(2);
    expect(native.getCols('Sheet1', exact)).toHaveLength(2);

    for (const collect of [
      () => native.getRows('Sheet1', { ...exact, limits: { ...exact.limits, maxRows: 1 } }),
      () => native.getCols('Sheet1', { ...exact, limits: { ...exact.limits, maxRows: 1 } }),
    ]) {
      const error = captureError(collect);
      expect(isResourceLimitError(error)).toBe(true);

      if (!isResourceLimitError(error)) throw new Error('Expected a resource limit error');
      expect(error.operation).toBe('data_traverse');
      expect(error.category).toBe('traversal');
      expect(error.dimension).toBe('row_count');
      expect(error.limit).toBe(1n);
      expect(error.observed).toBe(2n);
      expect(error.stage).toBe('serialize');
    }

    for (const collect of [
      () => native.getRows('Sheet1', { ...exact, limits: { ...exact.limits, maxColumns: 1 } }),
      () => native.getCols('Sheet1', { ...exact, limits: { ...exact.limits, maxColumns: 1 } }),
    ]) {
      const error = captureError(collect);
      expect(isResourceLimitError(error)).toBe(true);

      if (!isResourceLimitError(error)) throw new Error('Expected a resource limit error');
      expect(error.dimension).toBe('column_count');
      expect(error.limit).toBe(1n);
      expect(error.observed).toBe(2n);
    }
  });

  it('preflights shifted buffer coordinates before any workbook mutation', () => {
    const source = new Workbook();
    source.setCellValue('Sheet1', 'A1', 'first');
    source.setCellValue('Sheet1', 'A2', 'second');
    const buffer = source.getRowsBuffer('Sheet1');
    const oversizedFormulaSource = new Workbook();
    oversizedFormulaSource.setCellValue('Sheet1', 'A1', 'accepted-first');
    oversizedFormulaSource.setCellFormula('Sheet1', 'A2', 'x'.repeat(32_769), {
      perFormulaInputBytes: 32_769,
    });
    const oversizedFormulaBuffer = oversizedFormulaSource.getRowsBuffer('Sheet1');

    const target = new Workbook();
    target.setCellValue('Sheet1', 'A1', 'sentinel');
    target.setCellValue('Sheet1', 'A1048576', 'tail-sentinel');
    const before = target.writeBufferSync();

    const error = captureError(() => target.setSheetDataBuffer('Sheet1', buffer, 'A1048576'));
    expect(isResourceLimitError(error)).toBe(false);
    expect(String(error)).toMatch(/row|range/i);
    expect(target.writeBufferSync()).toEqual(before);
    expect(target.getCellValue('Sheet1', 'A1048576')).toBe('tail-sentinel');

    const missingSheetError = captureError(() =>
      target.setSheetDataBuffer('Missing', oversizedFormulaBuffer),
    );

    expect(isResourceLimitError(missingSheetError)).toBe(false);
    expect(String(missingSheetError)).toMatch(/sheet.*(?:not found|does not exist)/i);

    const shiftedFormulaError = captureError(() =>
      target.setSheetDataBuffer('Sheet1', oversizedFormulaBuffer, 'A1048576'),
    );

    expect(isResourceLimitError(shiftedFormulaError)).toBe(false);
    expect(String(shiftedFormulaError)).toMatch(/row|range/i);
    expect(target.writeBufferSync()).toEqual(before);

    const streamed = new Workbook();
    const writer = streamed.newStreamWriter('Streamed');
    writer.writeRow(1, ['sentinel']);
    streamed.applyStreamWriter(writer);

    const streamedError = captureError(() =>
      streamed.setSheetDataBuffer('Streamed', oversizedFormulaBuffer),
    );

    expect(isResourceLimitError(streamedError)).toBe(false);
    expect(String(streamedError)).toMatch(/streamed sheet/i);

    target.setSheetDataBuffer('Sheet1', buffer, 'C3');
    expect(target.getCellValue('Sheet1', 'C3')).toBe('first');
    expect(target.getCellValue('Sheet1', 'C4')).toBe('second');
  });

  it('preflights v1 and v2 transfer formulas at the secure input boundary', () => {
    const exactFormula = 'x'.repeat(32_768);
    const oversizedFormula = 'x'.repeat(32_769);

    const exactSource = new Workbook();
    exactSource.setCellFormula('Sheet1', 'A1', exactFormula);

    const oversizedSource = new Workbook();
    oversizedSource.setCellFormula('Sheet1', 'A1', 'accepted-first');
    oversizedSource.setCellFormula('Sheet1', 'B1', oversizedFormula, {
      perFormulaInputBytes: 32_769,
    });

    for (const [exactBuffer, oversizedBuffer] of [
      [exactSource.getRowsBuffer('Sheet1'), oversizedSource.getRowsBuffer('Sheet1')],
      [exactSource.getRowsBufferV2('Sheet1'), oversizedSource.getRowsBufferV2('Sheet1')],
    ]) {
      const accepted = new Workbook();
      accepted.setSheetDataBuffer('Sheet1', exactBuffer);
      expect(accepted.getCellValue('Sheet1', 'A1')).toEqual({
        type: 'formula',
        formula: exactFormula,
      });

      const rejected = new Workbook();
      rejected.setCellValue('Sheet1', 'A1', 'first-sentinel');
      rejected.setCellValue('Sheet1', 'B1', 'second-sentinel');
      const before = rejected.writeBufferSync();

      const error = captureError(() => rejected.setSheetDataBuffer('Sheet1', oversizedBuffer));
      expectResourceError(error, {
        operation: 'formula_mutate',
        category: 'formula',
        dimension: 'input_bytes',
        stage: 'validate',
        limit: 32_768n,
        observed: 32_769n,
      });
      expect(rejected.writeBufferSync()).toEqual(before);
      expect(rejected.getCellValue('Sheet1', 'A1')).toBe('first-sentinel');
      expect(rejected.getCellValue('Sheet1', 'B1')).toBe('second-sentinel');

      rejected.setSheetDataBuffer('Sheet1', exactBuffer);
      expect(rejected.getCellValue('Sheet1', 'A1')).toEqual({
        type: 'formula',
        formula: exactFormula,
      });
      expect(rejected.getCellValue('Sheet1', 'B1')).toBe('second-sentinel');
    }
  });

  it('shares the cumulative formula budget across v1 and v2 transfer imports', () => {
    const formula = 'x'.repeat(32_768);
    const exactSource = new Workbook();

    for (let row = 1; row <= 128; row++) {
      exactSource.setCellFormula('Sheet1', `A${row}`, formula);
    }

    const oversizedSource = new Workbook();

    for (let row = 1; row <= 128; row++) {
      oversizedSource.setCellFormula('Sheet1', `A${row}`, formula);
    }

    oversizedSource.setCellFormula('Sheet1', 'A129', 'x');

    const reusableSource = new Workbook();
    reusableSource.setCellValue('Sheet1', 'A1', 'reused');

    for (const [exactBuffer, oversizedBuffer, reusableBuffer] of [
      [
        exactSource.getRowsBuffer('Sheet1'),
        oversizedSource.getRowsBuffer('Sheet1'),
        reusableSource.getRowsBuffer('Sheet1'),
      ],
      [
        exactSource.getRowsBufferV2('Sheet1'),
        oversizedSource.getRowsBufferV2('Sheet1'),
        reusableSource.getRowsBufferV2('Sheet1'),
      ],
    ]) {
      const accepted = new Workbook();
      accepted.setSheetDataBuffer('Sheet1', exactBuffer);
      expect(accepted.getCellValue('Sheet1', 'A128')).toEqual({ type: 'formula', formula });

      const rejected = new Workbook();
      rejected.setCellValue('Sheet1', 'A1', 'first-sentinel');
      rejected.setCellValue('Sheet1', 'A129', 'last-sentinel');
      const before = rejected.writeBufferSync();

      const error = captureError(() => rejected.setSheetDataBuffer('Sheet1', oversizedBuffer));
      expectResourceError(error, {
        operation: 'formula_mutate',
        category: 'formula',
        dimension: 'input_bytes',
        stage: 'validate',
        limit: 4_194_304n,
        observed: 4_194_305n,
      });
      expect(rejected.writeBufferSync()).toEqual(before);
      expect(rejected.getCellValue('Sheet1', 'A1')).toBe('first-sentinel');
      expect(rejected.getCellValue('Sheet1', 'A129')).toBe('last-sentinel');

      rejected.setSheetDataBuffer('Sheet1', reusableBuffer);
      expect(rejected.getCellValue('Sheet1', 'A1')).toBe('reused');
      expect(rejected.getCellValue('Sheet1', 'A129')).toBe('last-sentinel');
    }
  });

  it('rejects malformed compatibility options instead of silently defaulting', async () => {
    const data = new Workbook().writeBufferSync();

    const invalidOptions: JsOpenOptions[] = [
      { readMode: 'unknown' },
      { auxParts: 'unknown' },
      { dateInterpretation: 'unknown' },
      { maxOuterInputBytes: Number.NaN },
      { maxUnzipSize: 1.5 },
      { cryptLimits: { kdfIterations: Number.POSITIVE_INFINITY } },
      { vbaLimits: { inputBytes: -1 } },
    ];

    for (const options of invalidOptions) {
      const syncError = captureError(() => Workbook.openBufferSync(data, options));

      expect(isResourceLimitError(syncError)).toBe(false);
      expect(String(syncError)).toMatch(/must be|readMode|auxParts|dateInterpretation/);

      const asyncError = await captureRejection(Workbook.openBuffer(data, options));

      expect(isResourceLimitError(asyncError)).toBe(false);
      expect(String(asyncError)).toMatch(/must be|readMode|auxParts|dateInterpretation/);
    }
  });

  it('checks async buffer outer limits before copying with exact parity', async () => {
    const data = new Workbook().writeBufferSync();
    const exact = { maxOuterInputBytes: data.length };

    expect(Workbook.openBufferSync(data, exact).sheetNames).toEqual(['Sheet1']);
    await expect(Workbook.openBuffer(data, exact)).resolves.toBeInstanceOf(Workbook);

    const rejected = { maxOuterInputBytes: data.length - 1 };
    const syncError = captureError(() => Workbook.openBufferSync(data, rejected));
    const asyncError = await captureRejection(Workbook.openBuffer(data, rejected));

    for (const error of [syncError, asyncError]) {
      expectResourceError(error, {
        operation: 'workbook_open',
        category: 'archive',
        dimension: 'input_bytes',
        stage: 'inspect',
        limit: BigInt(data.length - 1),
        observed: BigInt(data.length),
      });
    }
  });

  it('checks encrypted async buffer limits before copying with typed parity', async () => {
    const options = {
      maxOuterInputBytes: encryptedFixture.length,
      cryptLimits: { outerCfbBytes: encryptedFixture.length - 1 },
    };

    const syncError = captureError(() =>
      Workbook.openBufferWithPasswordSync(encryptedFixture, 'unused', options),
    );

    const asyncError = await captureRejection(
      Workbook.openBufferWithPassword(encryptedFixture, 'unused', options),
    );

    for (const error of [syncError, asyncError]) {
      expect(isResourceLimitError(error)).toBe(true);

      if (!isResourceLimitError(error)) throw new Error('Expected a resource limit error');
      expect(error.operation).toBe('decrypt');
      expect(error.category).toBe('container');
      expect(error.dimension).toBe('input_bytes');
      expect(error.stage).toBe('inspect');
      expect(error.limit).toBe(BigInt(encryptedFixture.length - 1));
      expect(error.observed).toBe(BigInt(encryptedFixture.length));
    }
  });

  it('preserves encrypted save exact and plus-one typed parity', async () => {
    const directory = mkdtempSync(join(tmpdir(), 'sheetkit-encrypted-save-parity-'));
    const path = join(directory, 'workbook.xlsx');
    const workbook = new Workbook();
    workbook.setCellValue('Sheet1', 'A1', 'value');

    try {
      workbook.saveWithPasswordSync(path, 'secret', {
        kdfIterations: 100_000,
        kdfWork: 300_000,
      });
      const accepted = readFileSync(path);
      expect(accepted.length).toBeGreaterThan(0);

      const rejected = { kdfIterations: 99_999, kdfWork: 300_000 };
      const syncError = captureError(() => workbook.saveWithPasswordSync(path, 'secret', rejected));

      const asyncError = await captureRejection(
        workbook.saveWithPassword(path, 'secret', rejected),
      );

      for (const error of [syncError, asyncError]) {
        expect(isResourceLimitError(error)).toBe(true);

        if (!isResourceLimitError(error)) throw new Error('Expected a resource limit error');
        expect(error.operation).toBe('encrypt');
        expect(error.category).toBe('encryption');
        expect(error.dimension).toBe('iteration_count');
        expect(error.stage).toBe('inspect');
        expect(error.limit).toBe(99_999n);
        expect(error.observed).toBe(100_000n);
      }

      expect(readFileSync(path)).toEqual(accepted);
      workbook.setCellValue('Sheet1', 'B1', 'task released');
      expect(workbook.getCellValue('Sheet1', 'B1')).toBe('task released');
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  }, 15_000);

  it('preserves typed render range exact and plus-one metadata', () => {
    const workbook = new Workbook();
    workbook.setCellValue('Sheet1', 'A1', 'first');
    workbook.setCellValue('Sheet1', 'B1', 'second');

    const exact = {
      mode: 'dense' as const,
      limits: { maxRows: 1, maxColumns: 2, maxCells: 2 },
    };

    expect(workbook.renderToSvg({ sheetName: 'Sheet1', range: 'A1:B1' }, exact)).toMatch(/^<svg/);

    const error = captureError(() =>
      workbook.renderToSvg(
        { sheetName: 'Sheet1', range: 'A1:B1' },
        { ...exact, limits: { ...exact.limits, maxCells: 1 } },
      ),
    );

    expectResourceError(error, {
      operation: 'render',
      category: 'traversal',
      dimension: 'cell_count',
      stage: 'render',
      limit: 1n,
      observed: 2n,
    });
  });

  it('preserves typed render output and default oversized-range errors', () => {
    const workbook = new Workbook();
    workbook.setCellValue('Sheet1', 'A1', 'value');

    const outputError = captureError(() =>
      workbook.renderToSvg(
        { sheetName: 'Sheet1', range: 'A1:A1' },
        { mode: 'dense', limits: { maxOutputBytes: 16 } },
      ),
    );

    expectResourceError(outputError, {
      operation: 'render',
      category: 'output',
      dimension: 'output_bytes',
      stage: 'render',
      limit: 16n,
    });

    const rangeError = captureError(() =>
      workbook.renderToSvg({ sheetName: 'Sheet1', range: 'A1:XFD1048576' }),
    );

    expectResourceError(rangeError, {
      operation: 'render',
      category: 'traversal',
      dimension: 'cell_count',
      stage: 'render',
    });
  });

  it('preserves typed pivot exact and plus-one errors without mutation', () => {
    const config = {
      name: 'PivotTable1',
      sourceSheet: 'Sheet1',
      sourceRange: 'A1:B2',
      targetSheet: 'PivotSheet',
      targetCell: 'A1',
      rows: [{ name: 'Region' }],
      columns: [],
      data: [{ name: 'Sales', function: 'sum' }],
    };

    const makeWorkbook = () => {
      const workbook = new Workbook();
      workbook.setRowValues('Sheet1', 1, 'A', ['Region', 'Sales']);
      workbook.setRowValues('Sheet1', 2, 'A', ['East', 10]);
      workbook.newSheet('PivotSheet');

      return workbook;
    };

    const exact = {
      mode: 'dense' as const,
      limits: { maxRows: 2, maxColumns: 2, maxCells: 4 },
    };

    const accepted = makeWorkbook();
    accepted.addPivotTable(config, exact);
    expect(accepted.getPivotTables()).toHaveLength(1);

    const rejected = makeWorkbook();

    const error = captureError(() =>
      rejected.addPivotTable(config, {
        ...exact,
        limits: { ...exact.limits, maxCells: 3 },
      }),
    );

    expectResourceError(error, {
      operation: 'data_traverse',
      category: 'traversal',
      dimension: 'cell_count',
      stage: 'serialize',
      limit: 3n,
      observed: 4n,
    });
    expect(rejected.getPivotTables()).toEqual([]);
  });

  it('rejects oversized pivot config allocations before core validation', () => {
    const oversized = 'x'.repeat(4096);

    const noColumns: { name: string }[] = [];

    const base = {
      name: 'PivotTable1',
      sourceSheet: 'MissingSheet',
      sourceRange: 'A1:B2',
      targetSheet: 'PivotSheet',
      targetCell: 'A1',
      rows: [{ name: 'Region' }],
      columns: noColumns,
      data: [{ name: 'Sales', function: 'sum' as const }],
    };

    const cases = [
      { ...base, name: oversized },
      { ...base, data: [{ ...base.data[0], displayName: oversized }] },
      {
        ...base,
        rows: Array.from({ length: 128 }, (_, index) => ({ name: `Field${index}` })),
      },
    ];

    for (const config of cases) {
      const workbook = new Workbook();
      workbook.newSheet('PivotSheet');

      const error = captureError(() =>
        workbook.addPivotTable(config, {
          mode: 'dense',
          limits: { maxOutputBytes: 512 },
        }),
      );

      expectResourceError(error, {
        operation: 'data_traverse',
        category: 'traversal',
        dimension: 'output_bytes',
        stage: 'serialize',
        limit: 512n,
        observed: 513n,
      });
      expect(workbook.getPivotTables()).toEqual([]);
    }
  });

  it('preserves typed pivot errors under secure finite defaults', () => {
    const workbook = new Workbook();
    workbook.newSheet('PivotSheet');

    const error = captureError(() =>
      workbook.addPivotTable({
        name: 'OversizedPivot',
        sourceSheet: 'Sheet1',
        sourceRange: 'A1:XFD1000',
        targetSheet: 'PivotSheet',
        targetCell: 'A1',
        rows: [],
        columns: [],
        data: [],
      }),
    );

    expectResourceError(error, {
      operation: 'data_traverse',
      category: 'traversal',
      dimension: 'cell_count',
      stage: 'serialize',
    });
    expect(workbook.getPivotTables()).toEqual([]);
  });
});

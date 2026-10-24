import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { Workbook as NativeWorkbook } from '../binding.js';
import { isResourceLimitError, Workbook } from '../index.js';
import { captureError, captureRejection } from './helpers/errors.js';

function expectBatchLimitError(error: Error, observed = 100_001n): void {
  expect(isResourceLimitError(error)).toBe(true);

  if (!isResourceLimitError(error)) throw new Error('Expected a resource limit error');
  expect(error.operation).toBe('stream_read');
  expect(error.category).toBe('streaming');
  expect(error.dimension).toBe('row_count');
  expect(error.stage).toBe('validate');
  expect(error.limit).toBe(100_000n);
  expect(error.observed).toBe(observed);
}

describe('Node streaming resource boundaries', () => {
  it('validates batch zero, exact, and plus one for path and buffer readers', async () => {
    const directory = mkdtempSync(join(tmpdir(), 'sheetkit-stream-policy-'));
    const path = join(directory, 'workbook.xlsx');
    const created = new Workbook();
    created.setCellValue('Sheet1', 'A1', 'value');
    created.saveSync(path);
    const buffer = created.writeBufferSync();

    try {
      const workbooks = [Workbook.openSync(path), Workbook.openBufferSync(buffer)];

      for (const workbook of workbooks) {
        const zero = await captureRejection(
          workbook.openSheetReader('MissingSheet', { batchSize: 0 }),
        );

        expect(isResourceLimitError(zero)).toBe(false);
        expect(String(zero)).toContain('greater than zero');

        const plusOne = await captureRejection(
          workbook.openSheetReader('MissingSheet', { batchSize: 100_001 }),
        );

        expectBatchLimitError(plusOne);
        expectBatchLimitError(
          await captureRejection(
            workbook.openSheetReader('MissingSheet', { batchSize: 4_294_967_297 }),
          ),
          4_294_967_297n,
        );

        for (const invalid of [-1, 1.5, Number.NaN, Number.POSITIVE_INFINITY]) {
          const error = await captureRejection(
            workbook.openSheetReader('MissingSheet', { batchSize: invalid }),
          );

          expect(isResourceLimitError(error)).toBe(false);
          expect(String(error)).toContain('positive safe integer');
        }

        const reader = await workbook.openSheetReader('Sheet1', { batchSize: 100_000 });
        const batch = await reader.next();
        expect(batch).toHaveLength(1);
        await reader.close();
      }
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });

  it('rejects per-read invalid batches before consuming reader state', async () => {
    const workbook = new Workbook();
    workbook.setCellValue('Sheet1', 'A1', 'first');
    workbook.setCellValue('Sheet1', 'A2', 'second');
    const reader = await workbook.openSheetReader('Sheet1', { batchSize: 1 });

    const zero = await captureRejection(reader.next(0));
    expect(isResourceLimitError(zero)).toBe(false);
    expect(String(zero)).toContain('greater than zero');
    expectBatchLimitError(await captureRejection(reader.next(100_001)));
    expectBatchLimitError(await captureRejection(reader.next(4_294_967_297)), 4_294_967_297n);

    for (const invalid of [-1, 1.5, Number.NaN, Number.NEGATIVE_INFINITY]) {
      const error = await captureRejection(reader.next(invalid));
      expect(isResourceLimitError(error)).toBe(false);
      expect(String(error)).toContain('positive safe integer');
    }

    expect((await reader.next())?.[0]?.row).toBe(1);
    expect((await reader.next())?.[0]?.row).toBe(2);
    expect(await reader.next()).toBeNull();

    await reader.close();
    expectBatchLimitError(await captureRejection(reader.next(100_001)));
  });

  it('validates direct native calls before lookup or reader state', () => {
    const created = new Workbook();
    created.setCellValue('Sheet1', 'A1', 'first');
    const native = NativeWorkbook.openBufferSync(created.writeBufferSync());

    expectBatchLimitError(
      captureError(() => native.openSheetReader('MissingSheet', 4_294_967_297)),
      4_294_967_297n,
    );
    expect(String(captureError(() => native.openSheetReader('MissingSheet', 1.5)))).toContain(
      'positive safe integer',
    );

    const reader = native.openSheetReader('Sheet1', 1);
    expect(String(captureError(() => reader.nextBatch(Number.NaN)))).toContain(
      'positive safe integer',
    );
    expect(reader.nextBatch(1)?.[0]?.row).toBe(1);
  });
});

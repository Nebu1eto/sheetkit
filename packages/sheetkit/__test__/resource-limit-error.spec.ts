import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { isResourceLimitError, Workbook } from '../index.js';
import { captureError, captureRejection } from './helpers/errors.js';

function expectOpenResourceError(error: Error) {
  expect(isResourceLimitError(error)).toBe(true);

  if (!isResourceLimitError(error)) throw new Error('Expected a resource limit error');
  expect(error).toBeInstanceOf(Error);
  expect(error.code).toBe('resource_limit_exceeded');
  expect(error.operation).toBe('workbook_open');
  expect(error.category).toBe('archive');
  expect(error.dimension).toBe('entry_count');
  expect(error.limit).toBeTypeOf('bigint');
  expect(error.observed).toBeTypeOf('bigint');
  expect(error.observed).toBeGreaterThan(error.limit);
  expect(error.stage).toBe('inspect');

  return error;
}

describe('resource limit errors', () => {
  it('recognizes only the closed public error contract', () => {
    const error = Object.assign(new Error('resource limit exceeded'), {
      code: 'resource_limit_exceeded',
      operation: 'workbook_open',
      category: 'archive',
      dimension: 'entry_count',
      limit: 10n,
      observed: 11n,
      stage: 'inspect',
    });

    expect(isResourceLimitError(error)).toBe(true);
    expect(isResourceLimitError({ ...error, operation: 'unknown' })).toBe(false);
    expect(isResourceLimitError({ ...error, limit: 10 })).toBe(false);
    expect(isResourceLimitError({ ...error, observed: 10n })).toBe(false);
    expect(isResourceLimitError({ ...error, stage: 'unknown' })).toBe(false);
    expect(isResourceLimitError({ ...error, stage: null })).toBe(true);

    const poisoned = Object.defineProperty({}, 'code', {
      get() {
        throw new Error('poisoned getter');
      },
    });

    expect(() => isResourceLimitError(poisoned)).not.toThrow();
    expect(isResourceLimitError(poisoned)).toBe(false);
  });

  it('preserves structured sync open metadata and bigint counters', () => {
    const workbook = new Workbook();
    const data = workbook.writeBufferSync();
    const error = captureError(() => Workbook.openBufferSync(data, { maxZipEntries: 1 }));

    expectOpenResourceError(error);
  });

  it('preserves structured async rejection metadata', async () => {
    const workbook = new Workbook();
    const data = workbook.writeBufferSync();

    await expect(Workbook.openBuffer(data, { maxZipEntries: 1 })).rejects.toSatisfy(
      isResourceLimitError,
    );
  });

  it('preserves structured sync and async path metadata', async () => {
    const directory = mkdtempSync(join(tmpdir(), 'sheetkit-resource-error-'));
    const path = join(directory, 'workbook.xlsx');

    try {
      new Workbook().saveSync(path);

      const syncError = expectOpenResourceError(
        captureError(() => Workbook.openSync(path, { maxZipEntries: 1 })),
      );

      const asyncError = expectOpenResourceError(
        await captureRejection(Workbook.open(path, { maxZipEntries: 1 })),
      );

      expect(asyncError.limit).toBe(syncError.limit);
      expect(asyncError.observed).toBe(syncError.observed);
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });
});

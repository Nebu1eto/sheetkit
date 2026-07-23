import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { isResourceLimitError, Workbook } from '../index.js';

const password = 'Password1234_';
const fixturePath = join(
  dirname(fileURLToPath(import.meta.url)),
  '../../../crates/sheetkit-core/tests/fixtures/encryption/msoffcrypto-tool-v6.0.0/agile-password.xlsx',
);
const encrypted = readFileSync(fixturePath);

function captureError(action: () => unknown): unknown {
  try {
    action();
  } catch (error) {
    return error;
  }
  throw new Error('Expected action to throw');
}

async function captureRejection(action: Promise<unknown>): Promise<unknown> {
  try {
    await action;
  } catch (error) {
    return error;
  }
  throw new Error('Expected promise to reject');
}

function expectResource(
  error: unknown,
  operation: string,
  category: string,
  dimension: string,
): void {
  expect(isResourceLimitError(error)).toBe(true);
  if (!isResourceLimitError(error)) throw new Error('Expected resource limit error');
  expect(error.operation).toBe(operation);
  expect(error.category).toBe(category);
  expect(error.dimension).toBe(dimension);
  expect(typeof error.limit).toBe('bigint');
  expect(typeof error.observed).toBe('bigint');
  expect(error.observed).toBeGreaterThan(error.limit);
}

function crc32(data: Buffer): number {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function appendStoredZipEntry(source: Buffer, path: string, data: Buffer): Buffer {
  let eocdOffset = -1;
  for (let offset = source.length - 22; offset >= 0; offset -= 1) {
    if (source.readUInt32LE(offset) === 0x06054b50) {
      eocdOffset = offset;
      break;
    }
  }
  if (eocdOffset < 0) throw new Error('EOCD not found');

  const entryCount = source.readUInt16LE(eocdOffset + 10);
  const centralSize = source.readUInt32LE(eocdOffset + 12);
  const centralOffset = source.readUInt32LE(eocdOffset + 16);
  const commentLength = source.readUInt16LE(eocdOffset + 20);
  const name = Buffer.from(path, 'utf8');
  const crc = crc32(data);

  const local = Buffer.alloc(30 + name.length + data.length);
  local.writeUInt32LE(0x04034b50, 0);
  local.writeUInt16LE(20, 4);
  local.writeUInt16LE(0, 6);
  local.writeUInt16LE(0, 8);
  local.writeUInt32LE(crc, 14);
  local.writeUInt32LE(data.length, 18);
  local.writeUInt32LE(data.length, 22);
  local.writeUInt16LE(name.length, 26);
  name.copy(local, 30);
  data.copy(local, 30 + name.length);

  const central = Buffer.alloc(46 + name.length);
  central.writeUInt32LE(0x02014b50, 0);
  central.writeUInt16LE(20, 4);
  central.writeUInt16LE(20, 6);
  central.writeUInt16LE(0, 8);
  central.writeUInt16LE(0, 10);
  central.writeUInt32LE(crc, 16);
  central.writeUInt32LE(data.length, 20);
  central.writeUInt32LE(data.length, 24);
  central.writeUInt16LE(name.length, 28);
  central.writeUInt32LE(centralOffset, 42);
  name.copy(central, 46);

  const eocd = Buffer.alloc(22 + commentLength);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(entryCount + 1, 8);
  eocd.writeUInt16LE(entryCount + 1, 10);
  eocd.writeUInt32LE(centralSize + central.length, 12);
  eocd.writeUInt32LE(centralOffset + local.length, 16);
  eocd.writeUInt16LE(commentLength, 20);
  source.copy(eocd, 22, eocdOffset + 22, eocdOffset + 22 + commentLength);

  return Buffer.concat([
    source.subarray(0, centralOffset),
    local,
    source.subarray(centralOffset, eocdOffset),
    central,
    eocd,
  ]);
}

describe('encrypted and VBA resource boundaries', () => {
  // Four 100K-iteration KDF opens need runtime tolerance, not a performance threshold.
  it('keeps password path and buffer sync/async behavior in parity', async () => {
    const pathSync = Workbook.openWithPasswordSync(fixturePath, password);
    const pathAsync = await Workbook.openWithPassword(fixturePath, password);
    const bufferSync = Workbook.openBufferWithPasswordSync(encrypted, password);
    const bufferAsync = await Workbook.openBufferWithPassword(encrypted, password);

    expect(pathAsync.sheetNames).toEqual(pathSync.sheetNames);
    expect(bufferSync.sheetNames).toEqual(pathSync.sheetNames);
    expect(bufferAsync.sheetNames).toEqual(pathSync.sheetNames);
  }, 15_000);

  it('preserves typed sync/async container limit errors', async () => {
    const options = { cryptLimits: { outerCfbBytes: encrypted.length - 1 } };
    const errors = [
      captureError(() => Workbook.openWithPasswordSync(fixturePath, password, options)),
      await captureRejection(Workbook.openWithPassword(fixturePath, password, options)),
      captureError(() => Workbook.openBufferWithPasswordSync(encrypted, password, options)),
      await captureRejection(Workbook.openBufferWithPassword(encrypted, password, options)),
    ];
    for (const error of errors) expectResource(error, 'decrypt', 'container', 'input_bytes');
  });

  // Multiple 100K-iteration KDF paths need runtime tolerance, not a performance threshold.
  it('keeps stream, KDF, and decrypted ZIP sync/async errors in parity', async () => {
    const cases = [
      {
        options: { cryptLimits: { namedStreamCount: 1 } },
        operation: 'decrypt',
        category: 'container',
        dimension: 'stream_count',
      },
      {
        options: { cryptLimits: { kdfIterations: 1 } },
        operation: 'decrypt',
        category: 'encryption',
        dimension: 'iteration_count',
      },
      {
        options: { maxUnzipSize: 1 },
        operation: 'workbook_open',
        category: 'archive',
        dimension: 'aggregate_declared_bytes',
      },
    ];
    for (const testCase of cases) {
      const syncError = captureError(() =>
        Workbook.openBufferWithPasswordSync(encrypted, password, testCase.options),
      );
      const asyncError = await captureRejection(
        Workbook.openBufferWithPassword(encrypted, password, testCase.options),
      );
      expectResource(syncError, testCase.operation, testCase.category, testCase.dimension);
      expectResource(asyncError, testCase.operation, testCase.category, testCase.dimension);
      if (!isResourceLimitError(syncError) || !isResourceLimitError(asyncError)) {
        throw new Error('Expected paired resource limit errors');
      }
      expect(asyncError.limit).toBe(syncError.limit);
      expect(asyncError.observed).toBe(syncError.observed);
    }

    expect(Workbook.openBufferWithPasswordSync(encrypted, password).sheetNames.length).toBe(1);
  }, 15_000);

  it('keeps wrong-password errors distinct from resource errors', async () => {
    const syncError = captureError(() =>
      Workbook.openBufferWithPasswordSync(encrypted, 'wrong password'),
    );
    const asyncError = await captureRejection(
      Workbook.openBufferWithPassword(encrypted, 'wrong password'),
    );
    expect(isResourceLimitError(syncError)).toBe(false);
    expect(isResourceLimitError(asyncError)).toBe(false);
  });

  it('applies exact and plus-one VBA input limits and typed module limits', () => {
    const vba = Buffer.from('bounded-vba-project');
    const xlsm = appendStoredZipEntry(new Workbook().writeBufferSync(), 'xl/vbaProject.bin', vba);
    const exact = Workbook.openBufferSync(xlsm, { vbaLimits: { inputBytes: vba.length } });
    expect(exact.getVbaProject()?.equals(vba)).toBe(true);

    expectResource(
      captureError(() =>
        Workbook.openBufferSync(xlsm, { vbaLimits: { inputBytes: vba.length - 1 } }),
      ),
      'macro_read',
      'macro',
      'input_bytes',
    );
    expectResource(
      captureError(() => exact.getVbaModules({ inputBytes: vba.length - 1 })),
      'macro_read',
      'macro',
      'input_bytes',
    );
  });

  it('delegates hard-ceiling validation to core limits', () => {
    expectResource(
      captureError(() =>
        Workbook.openBufferWithPasswordSync(encrypted, password, {
          cryptLimits: { kdfIterations: Number.MAX_SAFE_INTEGER },
        }),
      ),
      'decrypt',
      'encryption',
      'iteration_count',
    );
  });
});

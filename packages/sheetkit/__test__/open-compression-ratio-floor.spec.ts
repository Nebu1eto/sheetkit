import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { isResourceLimitError, type OpenOptions, Workbook } from '../index.js';

const MIB = 1024 * 1024;
const CEILING = 1024 * MIB;

const csvFixture = readFileSync(
  join(
    dirname(fileURLToPath(import.meta.url)),
    '../../../crates/sheetkit/tests/fixtures/csv-output-limit.xlsx',
  ),
);

/** SheetKit output whose repetitive parts exceed the default 100:1 ratio. */
function repetitiveOutput(): Buffer {
  return Workbook.openBufferSync(csvFixture).writeBufferSync();
}

function dimensionOf(action: () => unknown): string {
  try {
    action();
  } catch (error) {
    if (!isResourceLimitError(error)) throw error;
    return error.dimension;
  }
  throw new Error('Expected a resource limit error');
}

function messageOf(action: () => unknown): string {
  try {
    action();
  } catch (error) {
    expect(isResourceLimitError(error)).toBe(false);
    return String(error);
  }
  throw new Error('Expected an argument error');
}

describe('compressionRatioFloorBytes open option', () => {
  const data = repetitiveOutput();

  it.each([
    'eager',
    'lazy',
    'stream',
  ] as const)('reopens repetitive library output at defaults in %s mode', async (readMode) => {
    const sync = Workbook.openBufferSync(data, { readMode });
    expect(sync.getRows('Sheet1').length).toBeGreaterThan(0);
    const async = await Workbook.openBuffer(data, { readMode });
    expect(async.sheetNames).toEqual(['Sheet1']);
  });

  it('enforces the ratio for every entry when set to zero', async () => {
    const strict: OpenOptions = { compressionRatioFloorBytes: 0 };
    expect(dimensionOf(() => Workbook.openBufferSync(data, strict))).toBe('compression_ratio');
    await expect(Workbook.openBuffer(data, strict)).rejects.toSatisfy(
      (error: unknown) => isResourceLimitError(error) && error.dimension === 'compression_ratio',
    );
  });

  it('accepts the documented default and ceiling values', () => {
    for (const value of [8 * MIB, CEILING]) {
      const workbook = Workbook.openBufferSync(data, { compressionRatioFloorBytes: value });
      expect(workbook.sheetNames).toEqual(['Sheet1']);
    }
    expect(
      dimensionOf(() => Workbook.openBufferSync(data, { compressionRatioFloorBytes: CEILING + 1 })),
    ).toBe('compression_ratio_floor_bytes');
  });

  it.each([
    ['fractional', 1.5],
    ['negative', -1],
    ['NaN', Number.NaN],
    ['Infinity', Number.POSITIVE_INFINITY],
    ['2^53', 2 ** 53],
  ])('rejects %s values before opening', async (_name, value) => {
    const options: OpenOptions = { compressionRatioFloorBytes: value };
    expect(messageOf(() => Workbook.openBufferSync(data, options))).toMatch(
      /compressionRatioFloorBytes must be a finite non-negative safe integer/,
    );
    await expect(Workbook.openBuffer(data, options)).rejects.toThrow(
      /compressionRatioFloorBytes must be a finite non-negative safe integer/,
    );
  });
});

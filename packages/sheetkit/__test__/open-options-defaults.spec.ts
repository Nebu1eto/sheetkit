import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { Workbook } from '../index.js';
import { toNativeOpenOptions } from '../open-options.js';

const DEFAULTS = { readMode: 'lazy', auxParts: 'deferred' } as const;

describe('OpenOptions wrapper defaults', () => {
  it('uses readMode=lazy and auxParts=deferred when options are omitted', () => {
    expect(toNativeOpenOptions()).toEqual(DEFAULTS);
    expect(toNativeOpenOptions(undefined)).toEqual(DEFAULTS);
  });

  it('applies defaults when options are partial', () => {
    expect(toNativeOpenOptions({ sheetRows: 5 })).toEqual({
      sheetRows: 5,
      readMode: 'lazy',
      auxParts: 'deferred',
    });
    expect(toNativeOpenOptions({ auxParts: 'eager' })).toEqual({
      auxParts: 'eager',
      readMode: 'lazy',
    });
    expect(toNativeOpenOptions({ readMode: 'eager' })).toEqual({
      readMode: 'eager',
      auxParts: 'deferred',
    });
    expect(toNativeOpenOptions({ sheets: ['Sheet1'] })).toEqual({
      sheets: ['Sheet1'],
      readMode: 'lazy',
      auxParts: 'deferred',
    });
  });

  it('applies defaults when a partial option explicitly contains undefined', () => {
    expect(toNativeOpenOptions({ readMode: undefined })).toEqual({
      readMode: 'lazy',
      auxParts: 'deferred',
    });
    expect(toNativeOpenOptions({ auxParts: undefined })).toEqual({
      auxParts: 'deferred',
      readMode: 'lazy',
    });
  });

  it('applies defaults when options are null', () => {
    expect(toNativeOpenOptions(null)).toEqual(DEFAULTS);
  });

  it('preserves legacy parseMode-only options', () => {
    expect(toNativeOpenOptions({ parseMode: 'full' })).toEqual({ parseMode: 'full' });
  });
});

describe('OpenOptions defaults through the public open entry points', () => {
  let directory = '';
  let path = '';
  let buffer: Buffer = Buffer.alloc(0);

  beforeAll(() => {
    directory = mkdtempSync(join(tmpdir(), 'sheetkit-open-defaults-'));
    path = join(directory, 'tiny.xlsx');
    const workbook = new Workbook();
    workbook.setCellValue('Sheet1', 'A1', 'tiny');
    workbook.saveSync(path);
    buffer = workbook.writeBufferSync();
  });

  afterAll(() => {
    rmSync(directory, { recursive: true, force: true });
  });

  it('opens a real workbook from every entry point with omitted and null options', async () => {
    const opened = [
      Workbook.openSync(path),
      Workbook.openSync(path, null),
      await Workbook.open(path),
      Workbook.openBufferSync(buffer),
      await Workbook.openBuffer(buffer, null),
    ];

    for (const workbook of opened) {
      expect(workbook.sheetNames).toEqual(['Sheet1']);
      expect(workbook.getCellValue('Sheet1', 'A1')).toBe('tiny');
    }
  });
});

import { mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { type SaveLimits, type SaveOptions, Workbook } from '../index.js';
import { withForeignField } from './helpers/contract-violation.js';
import { captureError } from './helpers/errors.js';

function sampleWorkbook(): Workbook {
  const workbook = new Workbook();
  workbook.setCellValue('Sheet1', 'A1', 'durable');
  workbook.setCellValue('Sheet1', 'B2', 42);

  return workbook;
}

describe('save durability options', () => {
  let directory: string;

  beforeEach(() => {
    directory = mkdtempSync(join(tmpdir(), 'sheetkit-durable-'));
  });

  afterEach(() => {
    rmSync(directory, { recursive: true, force: true });
  });

  it('writes byte-identical readable files with and without durable', () => {
    const workbook = sampleWorkbook();
    const defaultPath = join(directory, 'default.xlsx');
    const durablePath = join(directory, 'durable.xlsx');
    const explicitPath = join(directory, 'explicit.xlsx');

    workbook.saveSync(defaultPath);
    workbook.saveSync(durablePath, undefined, { durable: true });
    workbook.saveSync(explicitPath, undefined, { durable: false });

    expect(readFileSync(durablePath).equals(readFileSync(defaultPath))).toBe(true);
    expect(readFileSync(explicitPath).equals(readFileSync(defaultPath))).toBe(true);
    expect(Workbook.openSync(durablePath).getCellValue('Sheet1', 'A1')).toBe('durable');
    expect(readdirSync(directory).filter((name) => name.startsWith('.sheetkit-'))).toEqual([]);
  });

  it('replaces an existing destination in durable async save', async () => {
    const path = join(directory, 'existing.xlsx');
    writeFileSync(path, 'old contents');

    await sampleWorkbook().save(path, undefined, { durable: true });

    const reopened = await Workbook.open(path);
    expect(reopened.getCellValue('Sheet1', 'B2')).toBe(42);
    expect(readdirSync(directory)).toEqual(['existing.xlsx']);
  });

  it('honours durable for sync and async encrypted saves', async () => {
    const workbook = sampleWorkbook();
    const syncPath = join(directory, 'sync.xlsx');
    const asyncPath = join(directory, 'async.xlsx');

    workbook.saveWithPasswordSync(syncPath, 'pw', undefined, { durable: true });
    await workbook.saveWithPassword(asyncPath, 'pw', undefined, { durable: true });

    expect(Workbook.openWithPasswordSync(syncPath, 'pw').getCellValue('Sheet1', 'A1')).toBe(
      'durable',
    );
    const reopened = await Workbook.openWithPassword(asyncPath, 'pw');
    expect(reopened.getCellValue('Sheet1', 'A1')).toBe('durable');
  });

  it('rejects a non-boolean durable like other invalid option fields', async () => {
    const workbook = sampleWorkbook();
    const path = join(directory, 'invalid.xlsx');
    const invalid = withForeignField<SaveOptions>({}, 'durable', 'yes');

    const durableError = captureError(() => workbook.saveSync(path, undefined, invalid));

    const limitError = captureError(() =>
      workbook.saveSync(path, withForeignField<SaveLimits>({}, 'outputBytes', 'many')),
    );

    expect(durableError).toBeInstanceOf(Error);
    expect(limitError).toBeInstanceOf(Error);
    expect(Object.getPrototypeOf(durableError)).toBe(Object.getPrototypeOf(limitError));
    expect(String(durableError)).toMatch(/durable/);
    await expect(workbook.save(path, undefined, invalid)).rejects.toThrow(/durable/);
    expect(() => workbook.saveWithPasswordSync(path, 'pw', undefined, invalid)).toThrow(/durable/);
    expect(readdirSync(directory)).toEqual([]);
  });
});

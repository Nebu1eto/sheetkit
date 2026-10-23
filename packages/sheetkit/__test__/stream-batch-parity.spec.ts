import assert from 'node:assert';
import { existsSync } from 'node:fs';
import { unlink } from 'node:fs/promises';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { type JsOpenOptions, Workbook as NativeWorkbook } from '../binding.js';
import { Workbook } from '../index.js';

const TEST_DIR = import.meta.dirname;
const FIXTURE_DIR = join(TEST_DIR, '..', '..', '..', 'benchmarks', 'node', 'fixtures');

type Batch = unknown[] | null;

/** Serialize with key order preserved so property order differences fail. */
function ordered(value: unknown): string {
  return JSON.stringify(value, (_key, entry) =>
    typeof entry === 'number' && Object.is(entry, -0) ? '-0' : entry,
  );
}

async function readPublic(
  path: string,
  batchSize: number,
  options: JsOpenOptions,
): Promise<Batch[]> {
  const workbook = await Workbook.open(path, options);
  const reader = await workbook.openSheetReader('Sheet1', { batchSize });
  const batches: Batch[] = [];
  for (;;) {
    const batch = await reader.next();
    batches.push(batch);
    if (batch === null) break;
  }
  await reader.close();
  return batches;
}

function readNative(path: string, batchSize: number, options: JsOpenOptions): Batch[] {
  const workbook = NativeWorkbook.openSync(path, options);
  const reader = workbook.openSheetReader('Sheet1', batchSize);
  const batches: Batch[] = [];
  for (;;) {
    const batch = reader.nextBatch(batchSize);
    batches.push(batch);
    if (batch === null) break;
  }
  reader.close();
  return batches;
}

async function expectParity(
  path: string,
  batchSize: number,
  options: JsOpenOptions,
): Promise<number> {
  const decoded = await readPublic(path, batchSize, options);
  const native = readNative(path, batchSize, options);
  expect(decoded).toStrictEqual(native);
  expect(ordered(decoded)).toBe(ordered(native));
  return decoded.reduce((total, batch) => total + (batch?.length ?? 0), 0);
}

describe('stream batch transfer parity', () => {
  const out = join(TEST_DIR, 'test-stream-batch-parity.xlsx');
  afterEach(async () => {
    if (existsSync(out)) await unlink(out);
  });

  it('decodes every cell shape identically to native row objects', async () => {
    const wb = new Workbook();
    const dateStyle = wb.addStyle({ numFmtId: 14 });
    const timeStyle = wb.addStyle({ customNumFmt: 'yyyy-mm-dd hh:mm:ss' });
    wb.setCellValue('Sheet1', 'A1', 'plain');
    wb.setCellValue('Sheet1', 'B1', 'escapes &<>"\' \\ \t tab');
    wb.setCellValue('Sheet1', 'C1', 'unicode \u00e9\u4e2d\ud55c\uae00 \ud83d\ude00');
    wb.setCellValue('Sheet1', 'D1', '');
    wb.setCellValue('Sheet1', 'E1', 0);
    wb.setCellValue('Sheet1', 'F1', -1.5e-300);
    wb.setCellValue('Sheet1', 'G1', 1.7976931348623157e308);
    wb.setCellValue('Sheet1', 'H1', true);
    wb.setCellValue('Sheet1', 'I1', false);
    wb.setCellValue('Sheet1', 'XFD1', 'last column');
    wb.setCellValue('Sheet1', 'A2', 46127);
    wb.setCellStyle('Sheet1', 'A2', dateStyle);
    wb.setCellValue('Sheet1', 'B2', 46127.75);
    wb.setCellStyle('Sheet1', 'B2', timeStyle);
    wb.setCellValue('Sheet1', 'C2', { type: 'date', serial: 45000, iso: '2023-03-15' });
    wb.setCellValue('Sheet1', 'D2', {
      type: 'formula',
      formula: 'A2+1',
      result: { valueType: 'number', numberValue: 46128 },
    });
    wb.setCellValue('Sheet1', 'E2', {
      type: 'formula',
      formula: '"x"&"y"',
      result: { valueType: 'string', value: 'xy' },
    });
    wb.setCellValue('Sheet1', 'F2', {
      type: 'formula',
      formula: 'TRUE()',
      result: { valueType: 'boolean', boolValue: true },
    });
    wb.setCellValue('Sheet1', 'G2', {
      type: 'formula',
      formula: '1/0',
      result: { valueType: 'error', value: '#DIV/0!' },
    });
    wb.setCellValue('Sheet1', 'H2', { type: 'formula', formula: 'SUM(A1:A9)' });
    wb.setCellRichText('Sheet1', 'I2', [{ text: 'rich ' }, { text: 'text', bold: true }]);
    for (let row = 3; row <= 40; row += 1) {
      wb.setCellValue('Sheet1', `A${row}`, `row ${row}`);
      wb.setCellValue('Sheet1', `C${row}`, row * 1.25);
    }
    wb.setCellValue('Sheet1', 'B1000', 'sparse row');
    await wb.save(out);

    for (const dateInterpretation of ['numFmt', 'cellType']) {
      for (const batchSize of [1, 3, 1000]) {
        const rows = await expectParity(out, batchSize, {
          readMode: 'stream',
          dateInterpretation,
        });
        expect(rows).toBe(41);
      }
    }
    const first = (await readPublic(out, 1000, { readMode: 'stream' }))[0];
    assert(first != null);
    expect(ordered(first).length).toBeGreaterThan(0);
  });

  it.each([
    'formulas.xlsx',
    'strings.xlsx',
    'comments.xlsx',
    'data-validation.xlsx',
    'scale-1k.xlsx',
  ])('matches native row objects for %s', async (fixture) => {
    const path = join(FIXTURE_DIR, fixture);
    const rows = await expectParity(path, 1000, { readMode: 'stream' });
    expect(rows).toBeGreaterThan(0);
  });

  it('reports the same batch-size errors as the native reader', async () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 1);
    await wb.save(out);
    const workbook = await Workbook.open(out, { readMode: 'stream' });
    const reader = await workbook.openSheetReader('Sheet1');
    await expect(reader.next(0)).rejects.toThrow();
    await expect(reader.next(100_001)).rejects.toThrow();
    const batch = await reader.next(1);
    expect(batch).toHaveLength(1);
    expect(await reader.next(1)).toBeNull();
  });
});

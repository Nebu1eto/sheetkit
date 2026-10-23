import assert from 'node:assert';
import { createHash } from 'node:crypto';
import { existsSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { decodeRowsBuffer } from '../buffer-codec.js';
import { Workbook } from '../index.js';

const FIXTURES = join(
  dirname(fileURLToPath(import.meta.url)),
  '..',
  '..',
  '..',
  'benchmarks',
  'node',
  'fixtures',
);

/**
 * Content hashes recorded from 4e02d37 for each benchmark fixture sheet:
 * decodeRowsBuffer(getRowsBufferV2), toJSON(), toJSON({ header: false }),
 * and toJSON({ mode: 'dense' }). Each hash is the first 16 hex digits of the
 * SHA-256 of the JSON serialization, which fixes values and key order.
 */
const EXPECTED: ReadonlyArray<readonly [string, string, string, string, string]> = [
  [
    'comments.xlsx/Sheet1',
    '210b582c96fd16d8',
    'd90e2cbbcd3dff5f',
    '170a7c3ee78d8aca',
    'd90e2cbbcd3dff5f',
  ],
  [
    'data-validation.xlsx/Sheet1',
    '34b705c8caedafb6',
    '6fe4bece88d39f65',
    '61ab7dc49405a1eb',
    '6fe4bece88d39f65',
  ],
  [
    'formulas.xlsx/Sheet1',
    '606a8a973046541a',
    '3ddd6891fc2b8395',
    'f7fb58611bf392df',
    '3ddd6891fc2b8395',
  ],
  [
    'heavy-styles.xlsx/Sheet1',
    'bef6a15fec1f84aa',
    'b8da97220743b2e1',
    'e930f2b08faf1cb2',
    'b8da97220743b2e1',
  ],
  [
    'large-data.xlsx/Sheet1',
    'ad6acb76f37af193',
    '2706c938784e8b44',
    '34c8ebbdbb9b52ee',
    '2706c938784e8b44',
  ],
  [
    'merged-cells.xlsx/Sheet1',
    '0a50a9ec93b1be67',
    'e166a5e1ea07720a',
    'bd61535e51c1ab3a',
    '982d63b1f546f6f9',
  ],
  [
    'mixed-workload.xlsx/Sheet1',
    'f8a0f6fbcad95367',
    'b85582c6c70de090',
    'ef11f6b563e368be',
    '1874e2bfc829c496',
  ],
  [
    'mixed-workload.xlsx/Employees',
    '2979b3af4b97808c',
    '873cc1c6fa27466f',
    'a3d2c2cf57d56beb',
    '13bacb94e5800cc0',
  ],
  [
    'mixed-workload.xlsx/Summary',
    '31ef1c6169caf5d2',
    '8465a5fbc5029673',
    'fa1fcd6b672f225d',
    '1962e30cf10537b4',
  ],
  [
    'multi-sheet.xlsx/Sheet1',
    '8f57584ab1fe2cbb',
    '3eccaae2d6e4e0c8',
    '3393d89156728749',
    '3eccaae2d6e4e0c8',
  ],
  [
    'multi-sheet.xlsx/Sheet2',
    'c78b498c048d59d7',
    '972741c6c2465e5b',
    '1ddf4194bb80de15',
    '972741c6c2465e5b',
  ],
  [
    'multi-sheet.xlsx/Sheet3',
    'c257a5a81103198e',
    'fda6f8d5d5c79dc4',
    '75a7edeffc1c1fc6',
    'fda6f8d5d5c79dc4',
  ],
  [
    'multi-sheet.xlsx/Sheet4',
    '5415b1a1d5b40a7f',
    '79fbf5f4b83a0efd',
    '1b21b5204d5d62f6',
    '79fbf5f4b83a0efd',
  ],
  [
    'multi-sheet.xlsx/Sheet5',
    '1aa27e610aba9a82',
    '61666c946738cf84',
    '87a1bcc148bec5e5',
    '61666c946738cf84',
  ],
  [
    'multi-sheet.xlsx/Sheet6',
    'a06f283ea0ae8a32',
    'a36633cacf34b847',
    '2378d5a0d93c1db7',
    'a36633cacf34b847',
  ],
  [
    'multi-sheet.xlsx/Sheet7',
    '15a55b5f20aa9260',
    '13e9bfd66a8536b4',
    '591c886f4b32c8a7',
    '13e9bfd66a8536b4',
  ],
  [
    'multi-sheet.xlsx/Sheet8',
    '65e4ac0a8752c3f5',
    '5a7e49e6e7c749f9',
    '3de9ac26462936a3',
    '5a7e49e6e7c749f9',
  ],
  [
    'multi-sheet.xlsx/Sheet9',
    '05e165326deca070',
    'badb33d59eab2769',
    'dd03c7942716aaa5',
    'badb33d59eab2769',
  ],
  [
    'multi-sheet.xlsx/Sheet10',
    '8c71d41c68ebc5c5',
    '3c02e2eaed3445eb',
    '46427f99e04357fb',
    '3c02e2eaed3445eb',
  ],
  [
    'scale-100k.xlsx/Sheet1',
    '2c12e53498feeeb2',
    'beaa1542ffa0819f',
    '04ef62cb092b10ec',
    'beaa1542ffa0819f',
  ],
  [
    'scale-10k.xlsx/Sheet1',
    '894ea4dafad58c4b',
    'c73cd618141d41e2',
    '608219f9d19f911c',
    'c73cd618141d41e2',
  ],
  [
    'scale-1k.xlsx/Sheet1',
    '8b22378ddefcc673',
    '43e8348699139bb4',
    '3ece88eeea50630b',
    '43e8348699139bb4',
  ],
  [
    'strings.xlsx/Sheet1',
    'ae06382fcd589d3d',
    'e1f9cd7c0102f554',
    '868893d51339907e',
    'e1f9cd7c0102f554',
  ],
];

function contentHash(value: unknown): string {
  return createHash('sha256').update(JSON.stringify(value)).digest('hex').slice(0, 16);
}

function expectPlainRecords(rows: Record<string, unknown>[]): void {
  for (const row of rows) {
    expect(Object.getPrototypeOf(row)).toBe(Object.prototype);
  }
}

describe.skipIf(!existsSync(FIXTURES))('benchmark fixture output identity', () => {
  it.each(EXPECTED)(
    '%s',
    (name, rowsHash, jsonHash, noHeaderHash, denseHash) => {
      const [file, sheet] = name.split('/');
      assert(file != null && sheet != null);
      const workbook = Workbook.openSync(join(FIXTURES, file));

      expect(contentHash(decodeRowsBuffer(workbook.getRowsBufferV2(sheet)))).toBe(rowsHash);
      const json = workbook.toJSON(sheet);
      const noHeader = workbook.toJSON(sheet, { header: false });
      const dense = workbook.toJSON(sheet, { mode: 'dense' });
      expectPlainRecords(json);
      expectPlainRecords(noHeader);
      expectPlainRecords(dense);
      expect([contentHash(json), contentHash(noHeader), contentHash(dense)]).toEqual([
        jsonHash,
        noHeaderHash,
        denseHash,
      ]);
    },
    120_000,
  );
});

describe('toJSON workbook-derived keys', () => {
  it('keeps prototype-named header cells as own data properties', () => {
    const workbook = new Workbook();
    workbook.setRowValues('Sheet1', 1, 'A', ['__proto__', 'constructor', 'toString', 'name']);
    workbook.setRowValues('Sheet1', 2, 'A', ['proto-value', 'ctor-value', 'string-value', 'v']);

    for (const mode of ['sparse', 'dense'] as const) {
      const [row] = workbook.toJSON('Sheet1', { mode });
      assert(row != null);
      expect(Object.getPrototypeOf(row)).toBe(Object.prototype);
      expect(Object.keys(row)).toEqual(['__proto__', 'constructor', 'toString', 'name']);
      for (const key of Object.keys(row)) {
        const descriptor = Object.getOwnPropertyDescriptor(row, key);
        assert(descriptor != null);
        expect(descriptor).toMatchObject({ writable: true, enumerable: true, configurable: true });
      }
      expect(Object.getOwnPropertyDescriptor(row, '__proto__')?.value).toBe('proto-value');
    }
  });

  it('does not invoke inherited setters for header cells or column letters', () => {
    const headerKey = '__sheetkit_header_setter__';
    let setterCalls = 0;
    const descriptor = {
      configurable: true,
      set: () => {
        setterCalls++;
      },
    };
    Object.defineProperty(Object.prototype, headerKey, descriptor);
    Object.defineProperty(Object.prototype, 'B', descriptor);
    try {
      const workbook = new Workbook();
      workbook.setRowValues('Sheet1', 1, 'A', [headerKey, 'plain']);
      workbook.setRowValues('Sheet1', 2, 'A', ['first', 'second']);

      for (const mode of ['sparse', 'dense'] as const) {
        const [row] = workbook.toJSON('Sheet1', { mode });
        assert(row != null);
        expect(Object.getOwnPropertyDescriptor(row, headerKey)?.value).toBe('first');
        expect(Object.getOwnPropertyDescriptor(row, 'plain')?.value).toBe('second');

        const [lettered] = workbook.toJSON('Sheet1', { mode, header: false });
        assert(lettered != null);
        expect(Object.getOwnPropertyDescriptor(lettered, 'A')?.value).toBe(headerKey);
        expect(Object.getOwnPropertyDescriptor(lettered, 'B')?.value).toBe('plain');
      }
      expect(setterCalls).toBe(0);
    } finally {
      Reflect.deleteProperty(Object.prototype, headerKey);
      Reflect.deleteProperty(Object.prototype, 'B');
    }
  });

  it('decodes ASCII and non-ASCII string cells identically in both modes', () => {
    const workbook = new Workbook();
    const values = ['plain', 'tab\tdel\u007f', 'caf\u00e9', '\ud55c\uae00', 'emoji \ud83d\ude00'];
    workbook.setRowValues('Sheet1', 1, 'A', ['key']);
    values.forEach((value, index) => {
      workbook.setCellValue('Sheet1', `A${index + 2}`, value);
    });
    for (const mode of ['sparse', 'dense'] as const) {
      expect(workbook.toJSON('Sheet1', { mode }).map((row) => row.key)).toEqual(values);
    }
  });
});

import { describe, expect, it } from 'vitest';
import { decodeRowsBuffer } from '../buffer-codec.js';
import { Workbook } from '../index.js';

/** Return the absolute offset of the first occurrence of `needle` in `buf`. */
function offsetOf(buf: Buffer, needle: string): number {
  const offset = buf.indexOf(Buffer.from(needle, 'utf8'));
  expect(offset).toBeGreaterThan(0);
  return offset;
}

describe('decodeRowsBuffer v2 string decoding', () => {
  const values = [
    'plain ascii',
    'tab\tand\u007fdel',
    'caf\u00e9',
    '\ud55c\uae00',
    'emoji \ud83d\ude00',
  ];

  function sourceBuffer(): Buffer {
    const wb = new Workbook();
    values.forEach((value, index) => {
      wb.setCellValue('Sheet1', `A${index + 1}`, value);
    });
    return Buffer.from(wb.getRowsBufferV2('Sheet1'));
  }

  it('decodes ASCII and non-ASCII strings exactly', () => {
    const rows = decodeRowsBuffer(sourceBuffer());
    const decoded = rows.map((row) => row.cells[0]?.value);
    expect(decoded).toStrictEqual(values);
  });

  it('still rejects malformed UTF-8 in otherwise ASCII text', () => {
    const buf = sourceBuffer();
    buf[offsetOf(buf, 'plain ascii') + 3] = 0xff;
    expect(() => decodeRowsBuffer(buf)).toThrow('malformed UTF-8 string');
  });

  it('still rejects a truncated multi-byte sequence', () => {
    const buf = sourceBuffer();
    const start = offsetOf(buf, 'caf\u00e9');
    // Replace the continuation byte of U+00E9 with an ASCII byte.
    buf[start + 4] = 0x41;
    expect(() => decodeRowsBuffer(buf)).toThrow('malformed UTF-8 string');
  });
});

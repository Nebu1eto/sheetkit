import assert from 'node:assert';
import { unlink } from 'node:fs/promises';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { Workbook } from '../index.js';

const TEST_DIR = import.meta.dirname;
const tmpFile = (name: string) => join(TEST_DIR, name);
const cleanup = async (...files: string[]) => {
  for (const f of files) await unlink(f).catch(() => {});
};

/**
 * Writes a workbook that mixes date-styled and non-date-styled number cells,
 * so every test can share the same fixture and assert on the four columns
 * individually.
 */
async function writeMixedStyleWorkbook(out: string): Promise<void> {
  const wb = new Workbook();
  const builtinDateStyle = wb.addStyle({ numFmtId: 14 }); // m/d/yyyy
  const customDateStyle = wb.addStyle({ customNumFmt: 'yyyy-mm-dd hh:mm' });
  const decimalStyle = wb.addStyle({ numFmtId: 2 }); // 0.00

  // A1: number + built-in date numFmt
  wb.setCellValue('Sheet1', 'A1', 46127);
  wb.setCellStyle('Sheet1', 'A1', builtinDateStyle);
  // B1: number + custom date format code
  wb.setCellValue('Sheet1', 'B1', 46127.9993);
  wb.setCellStyle('Sheet1', 'B1', customDateStyle);
  // C1: number + non-date format (decimal 2)
  wb.setCellValue('Sheet1', 'C1', 2.5);
  wb.setCellStyle('Sheet1', 'C1', decimalStyle);
  // D1: number without any explicit style
  wb.setCellValue('Sheet1', 'D1', 42);

  await wb.save(out);
}

describe('dateInterpretation option', () => {
  const out = tmpFile('test-date-interpretation.xlsx');
  afterEach(async () => cleanup(out));

  it('defaults to cellType: number cells stay as numbers even with date styles', async () => {
    await writeMixedStyleWorkbook(out);

    const wb = await Workbook.open(out, { readMode: 'lazy' });
    const reader = await wb.openSheetReader('Sheet1');
    const batch = await reader.next();
    assert(batch != null);

    expect(batch[0].cells).toHaveLength(4);
    for (const cell of batch[0].cells) {
      expect(cell.valueType).toBe('number');
    }
    expect(batch[0].cells[0].numberValue).toBe(46127);
    expect(batch[0].cells[2].numberValue).toBe(2.5);
    expect(batch[0].cells[3].numberValue).toBe(42);
    await reader.close();
  });

  it('numFmt mode promotes number cells with built-in or custom date formats to date', async () => {
    await writeMixedStyleWorkbook(out);

    const wb = await Workbook.open(out, {
      readMode: 'lazy',
      dateInterpretation: 'numFmt',
    });
    const reader = await wb.openSheetReader('Sheet1');
    const batch = await reader.next();
    assert(batch != null);

    // A1: built-in date numFmtId 14 -> promoted.
    expect(batch[0].cells[0].valueType).toBe('date');
    expect(batch[0].cells[0].numberValue).toBe(46127);

    // B1: custom date format code -> promoted.
    expect(batch[0].cells[1].valueType).toBe('date');
    expect(batch[0].cells[1].numberValue).toBeCloseTo(46127.9993, 4);

    // C1: non-date format -> stays a number.
    expect(batch[0].cells[2].valueType).toBe('number');
    expect(batch[0].cells[2].numberValue).toBe(2.5);

    // D1: no style -> stays a number.
    expect(batch[0].cells[3].valueType).toBe('number');
    expect(batch[0].cells[3].numberValue).toBe(42);

    await reader.close();
  });

  it('explicit cellType matches the default behavior', async () => {
    await writeMixedStyleWorkbook(out);

    const wb = await Workbook.open(out, {
      readMode: 'lazy',
      dateInterpretation: 'cellType',
    });
    const reader = await wb.openSheetReader('Sheet1');
    const batch = await reader.next();
    assert(batch != null);

    for (const cell of batch[0].cells) {
      expect(cell.valueType).toBe('number');
    }

    await reader.close();
  });
});

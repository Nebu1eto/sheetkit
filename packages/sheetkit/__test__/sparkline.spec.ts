import { unlink } from 'node:fs/promises';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { Workbook } from '../index.js';

const TEST_DIR = import.meta.dirname;

function tmpFile(name: string) {
  return join(TEST_DIR, name);
}

async function cleanup(...files: string[]) {
  for (const f of files) {
    await unlink(f).catch(() => {});
  }
}

describe('Sparklines', () => {
  const out = tmpFile('test-sparklines.xlsx');
  afterEach(async () => cleanup(out));

  it('should add and get sparklines', () => {
    const wb = new Workbook();
    wb.addSparkline('Sheet1', {
      dataRange: 'Sheet1!A1:A10',
      location: 'B1',
    });

    const sparklines = wb.getSparklines('Sheet1');
    expect(sparklines).toHaveLength(1);
    expect(sparklines[0].dataRange).toBe('Sheet1!A1:A10');
    expect(sparklines[0].location).toBe('B1');
    expect(sparklines[0].sparklineType).toBe('line');
  });

  it('should add multiple sparklines to one sheet', () => {
    const wb = new Workbook();
    wb.addSparkline('Sheet1', {
      dataRange: 'Sheet1!A1:A10',
      location: 'B1',
    });
    wb.addSparkline('Sheet1', {
      dataRange: 'Sheet1!A1:A10',
      location: 'B2',
      sparklineType: 'column',
      markers: true,
      highPoint: true,
    });

    const sparklines = wb.getSparklines('Sheet1');
    expect(sparklines).toHaveLength(2);
    expect(sparklines[1].sparklineType).toBe('column');
    expect(sparklines[1].markers).toBe(true);
    expect(sparklines[1].highPoint).toBe(true);
  });

  it('should remove sparkline by location', () => {
    const wb = new Workbook();
    wb.addSparkline('Sheet1', {
      dataRange: 'Sheet1!A1:A10',
      location: 'B1',
    });
    wb.addSparkline('Sheet1', {
      dataRange: 'Sheet1!A1:A10',
      location: 'B2',
    });

    wb.removeSparkline('Sheet1', 'B1');

    const sparklines = wb.getSparklines('Sheet1');
    expect(sparklines).toHaveLength(1);
    expect(sparklines[0].location).toBe('B2');
  });

  it('should throw on non-existent sheet', () => {
    const wb = new Workbook();
    expect(() =>
      wb.addSparkline('NoSheet', {
        dataRange: 'Sheet1!A1:A10',
        location: 'B1',
      }),
    ).toThrow();
    expect(() => wb.getSparklines('NoSheet')).toThrow();
    expect(() => wb.removeSparkline('NoSheet', 'B1')).toThrow();
  });

  it('should preserve sparklines through save/open roundtrip', async () => {
    const wb = new Workbook();
    for (let i = 1; i <= 10; i++) {
      wb.setCellValue('Sheet1', `A${i}`, i * 10);
    }

    wb.addSparkline('Sheet1', {
      dataRange: 'Sheet1!A1:A10',
      location: 'B1',
      sparklineType: 'column',
      markers: true,
      highPoint: true,
      lineWeight: 1.5,
    });

    wb.addSparkline('Sheet1', {
      dataRange: 'Sheet1!A1:A5',
      location: 'C1',
      sparklineType: 'line',
    });

    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const sparklines = wb2.getSparklines('Sheet1');
    expect(sparklines).toHaveLength(2);
    expect(sparklines[0].dataRange).toBe('Sheet1!A1:A10');
    expect(sparklines[0].location).toBe('B1');
    expect(sparklines[0].sparklineType).toBe('column');
    expect(sparklines[0].markers).toBe(true);
    expect(sparklines[0].highPoint).toBe(true);
    expect(sparklines[0].lineWeight).toBe(1.5);
    expect(sparklines[1].dataRange).toBe('Sheet1!A1:A5');
    expect(sparklines[1].location).toBe('C1');
  });

  it('should return empty array for sheet with no sparklines', () => {
    const wb = new Workbook();
    const sparklines = wb.getSparklines('Sheet1');
    expect(sparklines).toHaveLength(0);
  });
});

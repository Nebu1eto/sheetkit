import { describe, expect, it } from 'vitest';
import { Workbook as NativeWorkbook } from '../binding.js';
import { isResourceLimitError, Workbook } from '../index.js';

const INVALID_INTEGERS: readonly [string, number][] = [
  ['fractional', 1.9],
  ['negative', -1],
  ['negative wrap', -4_294_967_295],
  ['NaN', Number.NaN],
  ['Infinity', Number.POSITIVE_INFINITY],
  ['-Infinity', Number.NEGATIVE_INFINITY],
  ['2^32 + 1', 4_294_967_297],
  ['2^53', 2 ** 53],
];

const INTEGER_ERROR = /must be an integer/;

interface RowTarget {
  setCellValue(sheet: string, cell: string, value: string): void;
  getCellValue(sheet: string, cell: string): unknown;
}

function seed(workbook: RowTarget): void {
  workbook.setCellValue('Sheet1', 'A1', 'r1');
  workbook.setCellValue('Sheet1', 'A2', 'r2');
  workbook.setCellValue('Sheet1', 'B1', 'b1');
}

function snapshot(workbook: RowTarget): unknown[] {
  return ['A1', 'A2', 'A3', 'B1', 'C1'].map((cell) => workbook.getCellValue('Sheet1', cell));
}

const SEEDED = ['r1', 'r2', null, 'b1', null];

function seededBuffer(rows: number): Buffer {
  const workbook = new Workbook();
  for (let row = 1; row <= rows; row++) {
    workbook.setCellValue('Sheet1', `A${row}`, `r${row}`);
  }
  return workbook.writeBufferSync();
}

const factories = [
  ['public wrapper', () => new Workbook()],
  ['raw binding', () => new NativeWorkbook()],
] as const;

describe.each(factories)('numeric argument validation via %s', (_label, create) => {
  it.each(INVALID_INTEGERS)('removeRow rejects %s rows before mutation', (_name, row) => {
    const workbook = create();
    seed(workbook);
    expect(() => workbook.removeRow('Sheet1', row)).toThrow(INTEGER_ERROR);
    expect(snapshot(workbook)).toEqual(SEEDED);
  });

  it.each(INVALID_INTEGERS)('insertRows rejects %s start rows and counts', (_name, value) => {
    const workbook = create();
    seed(workbook);
    expect(() => workbook.insertRows('Sheet1', value, 1)).toThrow(INTEGER_ERROR);
    expect(() => workbook.insertRows('Sheet1', 1, value)).toThrow(INTEGER_ERROR);
    expect(snapshot(workbook)).toEqual(SEEDED);
  });

  it.each(INVALID_INTEGERS)('insertCols rejects %s counts before mutation', (_name, count) => {
    const workbook = create();
    seed(workbook);
    expect(() => workbook.insertCols('Sheet1', 'A', count)).toThrow(INTEGER_ERROR);
    expect(snapshot(workbook)).toEqual(SEEDED);
  });

  it.each(INVALID_INTEGERS)('outline setters reject %s levels and rows', (_name, value) => {
    const workbook = create();
    seed(workbook);
    expect(() => workbook.setRowOutlineLevel('Sheet1', 1, value)).toThrow(INTEGER_ERROR);
    expect(() => workbook.setRowOutlineLevel('Sheet1', value, 1)).toThrow(INTEGER_ERROR);
    expect(() => workbook.setColOutlineLevel('Sheet1', 'A', value)).toThrow(INTEGER_ERROR);
    expect(workbook.getRowOutlineLevel('Sheet1', 1)).toBe(0);
    expect(workbook.getColOutlineLevel('Sheet1', 'A')).toBe(0);
  });

  it.each(INVALID_INTEGERS)('row and style setters reject %s values', (_name, value) => {
    const workbook = create();
    seed(workbook);
    const styleId = workbook.addStyle({ font: { bold: true } });
    expect(() => workbook.setRowStyle('Sheet1', value, styleId)).toThrow(INTEGER_ERROR);
    expect(() => workbook.setRowStyle('Sheet1', 1, value)).toThrow(INTEGER_ERROR);
    expect(() => workbook.setCellStyle('Sheet1', 'A1', value)).toThrow(INTEGER_ERROR);
    expect(() => workbook.setColStyle('Sheet1', 'A', value)).toThrow(INTEGER_ERROR);
    expect(() => workbook.setRowVisible('Sheet1', value, false)).toThrow(INTEGER_ERROR);
    expect(() => workbook.setRowHeight('Sheet1', value, 20)).toThrow(INTEGER_ERROR);
    expect(() => workbook.duplicateRow('Sheet1', value)).toThrow(INTEGER_ERROR);
    expect(() => workbook.insertPageBreak('Sheet1', value)).toThrow(INTEGER_ERROR);
    expect(() => workbook.setRowValues('Sheet1', value, 'A', ['x'])).toThrow(INTEGER_ERROR);
    expect(() => workbook.getRowHeight('Sheet1', value)).toThrow(INTEGER_ERROR);
    expect(() => workbook.deleteFormControl('Sheet1', value)).toThrow(INTEGER_ERROR);
    expect(workbook.getRowStyle('Sheet1', 1)).toBe(0);
    expect(workbook.getCellStyle('Sheet1', 'A1')).toBeNull();
    expect(workbook.getRowVisible('Sheet1', 1)).toBe(true);
    expect(workbook.getPageBreaks('Sheet1')).toEqual([]);
    expect(snapshot(workbook)).toEqual(SEEDED);
  });

  it('keeps exact boundary values', () => {
    const workbook = create();
    seed(workbook);
    workbook.setRowOutlineLevel('Sheet1', 1_048_576, 7);
    expect(workbook.getRowOutlineLevel('Sheet1', 1_048_576)).toBe(7);
    workbook.setColOutlineLevel('Sheet1', 'XFD', 7);
    expect(workbook.getColOutlineLevel('Sheet1', 'XFD')).toBe(7);
    expect(() => workbook.setRowOutlineLevel('Sheet1', 1, 8)).toThrow(
      'outline level 8 exceeds maximum 7',
    );
    expect(() => workbook.setRowOutlineLevel('Sheet1', 1, 256)).toThrow(INTEGER_ERROR);
    expect(() => workbook.removeRow('Sheet1', 0)).toThrow('invalid row number: 0');
    workbook.insertRows('Sheet1', 1, 0);
    workbook.insertCols('Sheet1', 'A', 0);
    expect(snapshot(workbook)).toEqual(SEEDED);
    workbook.removeRow('Sheet1', 1);
    expect(workbook.getCellValue('Sheet1', 'A1')).toBe('r2');
  });

  it.each(
    INVALID_INTEGERS,
  )('open options reject %s sheetRows and maxZipEntries', (_name, value) => {
    const data = seededBuffer(3);
    const open = (options: { sheetRows?: number; maxZipEntries?: number }): unknown =>
      create === factories[0][1]
        ? Workbook.openBufferSync(data, options)
        : NativeWorkbook.openBufferSync(data, options);
    expect(() => open({ sheetRows: value })).toThrow(/sheetRows must be an integer/);
    // 2^32 + 1 is a safe integer, so it reaches the hard ceiling check
    // instead of wrapping to 1 as the previous u32 coercion did.
    expect(() => open({ maxZipEntries: value })).toThrow(
      /maxZipEntries must be|resource limit exceeded/,
    );
  });
});

describe('numeric option boundaries', () => {
  it('accepts exact sheetRows and maxZipEntries limits', () => {
    const data = seededBuffer(3);
    const limited = Workbook.openBufferSync(data, { sheetRows: 2 });
    expect(limited.getCellValue('Sheet1', 'A2')).toBe('r2');
    expect(limited.getCellValue('Sheet1', 'A3')).toBeNull();
    const all = Workbook.openBufferSync(data, { sheetRows: 4_294_967_295 });
    expect(all.getCellValue('Sheet1', 'A3')).toBe('r3');
    expect(() => Workbook.openBufferSync(data, { maxZipEntries: 100 })).not.toThrow();
  });

  it.each(INVALID_INTEGERS)('rejects %s numeric style and view fields', (_name, value) => {
    const workbook = new Workbook();
    expect(() => workbook.addStyle({ numFmtId: value })).toThrow(INTEGER_ERROR);
    expect(() => workbook.addStyle({ alignment: { indent: value } })).toThrow(INTEGER_ERROR);
    expect(() => workbook.addStyle({ alignment: { textRotation: value } })).toThrow(INTEGER_ERROR);
    expect(() => workbook.setSheetViewOptions('Sheet1', { zoomScale: value })).toThrow(
      INTEGER_ERROR,
    );
    expect(() =>
      workbook.setPageSetup('Sheet1', { scale: value, fitToWidth: 1, fitToHeight: 1 }),
    ).toThrow(INTEGER_ERROR);
  });

  it('rejects unsafe builtinFormatCode and getThemeColor ids', () => {
    for (const [, value] of INVALID_INTEGERS) {
      expect(() => NativeWorkbook.builtinFormatCode(value)).toThrow(INTEGER_ERROR);
      expect(() => new Workbook().getThemeColor(value)).toThrow(INTEGER_ERROR);
    }
    expect(NativeWorkbook.builtinFormatCode(0)).toBe('General');
  });
});

describe('stream writer numeric arguments', () => {
  it.each(INVALID_INTEGERS)('rejects %s rows, columns, styles, and levels', (_name, value) => {
    const workbook = new Workbook();
    const writer = workbook.newStreamWriter('Streamed');
    expect(() => writer.writeRow(value, ['x'])).toThrow(INTEGER_ERROR);
    expect(() => writer.writeRows(value, [['x']])).toThrow(INTEGER_ERROR);
    expect(() => writer.writeRowWithStyle(1, ['x'], value)).toThrow(INTEGER_ERROR);
    expect(() => writer.setColWidth(value, 10)).toThrow(INTEGER_ERROR);
    expect(() => writer.setColWidthRange(value, 2, 10)).toThrow(INTEGER_ERROR);
    expect(() => writer.setColStyle(value, 0)).toThrow(INTEGER_ERROR);
    expect(() => writer.setColVisible(value, false)).toThrow(INTEGER_ERROR);
    expect(() => writer.setColOutlineLevel(1, value)).toThrow(INTEGER_ERROR);
    writer.writeRow(1, ['ok']);
    workbook.applyStreamWriter(writer);
    const reopened = Workbook.openBufferSync(workbook.writeBufferSync());
    expect(reopened.getCellValue('Streamed', 'A1')).toBe('ok');
    expect(reopened.getCellValue('Streamed', 'A2')).toBeNull();
    expect(reopened.getColOutlineLevel('Streamed', 'A')).toBe(0);
    expect(reopened.getColVisible('Streamed', 'A')).toBe(true);
  });

  it.each(INVALID_INTEGERS)('rejects %s sheet reader batch sizes', async (_name, value) => {
    const workbook = Workbook.openBufferSync(seededBuffer(2), { readMode: 'stream' });
    await expect(workbook.openSheetReader('Sheet1', { batchSize: value })).rejects.toThrow(
      /batch size|resource limit exceeded/,
    );
  });
});

describe('enum string validation', () => {
  it('rejects unknown style enum strings with accepted values', () => {
    const workbook = new Workbook();
    expect(() => workbook.addStyle({ fill: { pattern: 'polkadot' } })).toThrow(
      /unknown fill pattern 'polkadot'.*solid/,
    );
    expect(() => workbook.addStyle({ border: { left: { style: 'wavy' } } })).toThrow(
      /unknown border style 'wavy'.*thin/,
    );
    expect(() => workbook.addStyle({ alignment: { horizontal: 'middle' } })).toThrow(
      /unknown horizontal alignment 'middle'.*center/,
    );
    expect(() => workbook.addStyle({ alignment: { vertical: 'middle' } })).toThrow(
      /unknown vertical alignment 'middle'.*center/,
    );
    expect(workbook.addStyle({ fill: { pattern: 'Solid', fgColor: 'FF0000' } })).toBeGreaterThan(0);
  });
});

describe('toHTML output limit', () => {
  it('charges final escaped HTML bytes against maxOutputBytes', () => {
    const workbook = new Workbook();
    workbook.setSheetData('Sheet1', [['<&>"'.repeat(10)], ['plain']]);
    const html = workbook.toHTML('Sheet1');
    const exact = Buffer.byteLength(html);
    expect(workbook.toHTML('Sheet1', { limits: { maxOutputBytes: exact } })).toBe(html);
    let error: unknown;
    try {
      workbook.toHTML('Sheet1', { limits: { maxOutputBytes: exact - 1 } });
    } catch (caught) {
      error = caught;
    }
    expect(isResourceLimitError(error)).toBe(true);
    expect(error).toMatchObject({
      operation: 'export',
      category: 'output',
      dimension: 'output_bytes',
      limit: BigInt(exact - 1),
      stage: 'serialize',
    });
  });
});

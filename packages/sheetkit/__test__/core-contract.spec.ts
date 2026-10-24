import { inflateRawSync } from 'node:zlib';
import { describe, expect, it, vi } from 'vitest';
import type { JsRowData } from '../binding.js';
import { Workbook as NativeWorkbook } from '../binding.js';
import { type FormulaValue, isResourceLimitError, Workbook } from '../index.js';

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

/** Build a ZIP archive with stored (uncompressed) entries. */
function storedZip(entries: [string, Buffer][]): Buffer {
  const locals: Buffer[] = [];
  const centrals: Buffer[] = [];
  let offset = 0;

  for (const [path, data] of entries) {
    const name = Buffer.from(path, 'utf8');
    const crc = crc32(data);

    const local = Buffer.alloc(30 + name.length + data.length);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
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
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(data.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(name.length, 28);
    central.writeUInt32LE(offset, 42);
    name.copy(central, 46);

    locals.push(local);
    centrals.push(central);
    offset += local.length;
  }

  const centralSize = centrals.reduce((sum, entry) => sum + entry.length, 0);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(centralSize, 12);
  eocd.writeUInt32LE(offset, 16);

  return Buffer.concat([...locals, ...centrals, eocd]);
}

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/** Read every entry of a ZIP archive, inflating deflated entries. */
function readZip(source: Buffer): [string, Buffer][] {
  let eocd = -1;

  for (let offset = source.length - 22; offset >= 0; offset -= 1) {
    if (source.readUInt32LE(offset) === 0x06054b50) {
      eocd = offset;
      break;
    }
  }

  if (eocd < 0) throw new Error('EOCD not found');
  const count = source.readUInt16LE(eocd + 10);
  let cursor = source.readUInt32LE(eocd + 16);
  const entries: [string, Buffer][] = [];

  for (let index = 0; index < count; index += 1) {
    const method = source.readUInt16LE(cursor + 10);
    const compressedSize = source.readUInt32LE(cursor + 20);
    const nameLength = source.readUInt16LE(cursor + 28);
    const extraLength = source.readUInt16LE(cursor + 30);
    const commentLength = source.readUInt16LE(cursor + 32);
    const localOffset = source.readUInt32LE(cursor + 42);
    const name = source.toString('utf8', cursor + 46, cursor + 46 + nameLength);

    const dataStart =
      localOffset +
      30 +
      source.readUInt16LE(localOffset + 26) +
      source.readUInt16LE(localOffset + 28);

    const raw = source.subarray(dataStart, dataStart + compressedSize);
    entries.push([name, method === 8 ? inflateRawSync(raw) : Buffer.from(raw)]);
    cursor += 46 + nameLength + extraLength + commentLength;
  }

  return entries;
}

function workbookWithSheetXml(sheetXml: string): Buffer {
  const base = new Workbook();
  base.setCellValue('Sheet1', 'A1', 1);

  const entries = readZip(base.writeBufferSync()).map(([name, data]): [string, Buffer] =>
    name === 'xl/worksheets/sheet1.xml' ? [name, Buffer.from(sheetXml, 'utf8')] : [name, data],
  );

  return storedZip(entries);
}

const SHEET_PREFIX =
  `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${MAIN_NS}"><sheetData>` +
  '<row r="1"><c r="A1"><v>1</v></c></row>';

function tableConfig(name: string, range: string) {
  return {
    name,
    displayName: name,
    range,
    columns: [{ name: 'Col1' }, { name: 'Col2' }],
    showHeaderRow: true,
    autoFilter: false,
  };
}

describe('table name validation', () => {
  it('rejects cell-like, invalid, and case-only duplicate names', () => {
    const wb = new Workbook();
    expect(() => wb.addTable('Sheet1', tableConfig('A1', 'A1:B5'))).toThrow(
      /^invalid argument: .*must not be a cell reference/,
    );
    expect(() => wb.addTable('Sheet1', tableConfig('Bad Name', 'A1:B5'))).toThrow(
      /^invalid argument: .*invalid character/,
    );
    expect(wb.getTables('Sheet1')).toEqual([]);

    wb.addTable('Sheet1', tableConfig('SalesTable', 'A1:B5'));
    expect(() => wb.addTable('Sheet1', tableConfig('SALESTABLE', 'D1:E5'))).toThrow(
      /already exists/,
    );
    expect(wb.getTables('Sheet1').map((table) => table.name)).toEqual(['SalesTable']);
  });
});

describe('stream writer merge cells', () => {
  it('canonicalizes ranges and rejects overlaps', () => {
    const wb = new Workbook();
    const writer = wb.newStreamWriter('Merged');
    writer.addMergeCell('$b$2:a1');
    writer.addMergeCell('D4:C3');
    expect(() => writer.addMergeCell('B2:C3')).toThrow(/overlaps with existing range/);
    writer.writeRow(1, ['x']);
    wb.applyStreamWriter(writer);

    const reopened = Workbook.openBufferSync(wb.writeBufferSync());
    expect(reopened.getMergeCells('Merged')).toEqual(['A1:B2', 'C3:D4']);
  });
});

describe('stream reader truncated XML', () => {
  it('opens a well-formed control workbook', async () => {
    const wb = Workbook.openBufferSync(
      workbookWithSheetXml(`${SHEET_PREFIX}</sheetData></worksheet>`),
    );

    const reader = await wb.openSheetReader('Sheet1');
    const batch = await reader.next();
    expect(batch?.map((row) => row.row)).toEqual([1]);
    await reader.close();
  });

  it('rejects a truncated worksheet instead of returning partial rows', async () => {
    const data = workbookWithSheetXml(`${SHEET_PREFIX}<row r="2"><c r="A2"><v>2</v>`);

    const readAll = async (): Promise<number> => {
      const wb = Workbook.openBufferSync(data);
      const reader = await wb.openSheetReader('Sheet1');
      let rows = 0;

      for (;;) {
        const batch = await reader.next();

        if (batch === null) break;
        rows += batch.length;
      }

      return rows;
    };

    await expect(readAll()).rejects.toThrow('unexpected end of worksheet XML inside an open cell');
  });
});

const FORMULA_SHEET_XML =
  `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${MAIN_NS}"><sheetData>` +
  '<row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c>' +
  '<c r="B1" t="str"><f>A1&amp;"y"</f><v>xy</v></c>' +
  '<c r="C1"><f>SUM(A2,1)</f></c>' +
  '<c r="D1"><f t="shared" ref="D1:D2" si="0">A2*2</f><v>6</v></c>' +
  '<c r="E1" t="e"><f>1/0</f><v>#DIV/0!</v></c></row>' +
  '<row r="2"><c r="A2"><v>3</v></c>' +
  '<c r="B2" t="b"><f>A2&gt;1</f><v>1</v></c>' +
  '<c r="D2"><f t="shared" si="0"/><v>8</v></c></row>' +
  '</sheetData></worksheet>';

/** Formula cells of FORMULA_SHEET_XML; D2 is a shared-formula follower and C1 has no cached value. */
const EXPECTED_FORMULAS = {
  B1: { type: 'formula', formula: 'A1&"y"', result: { valueType: 'string', value: 'xy' } },
  C1: { type: 'formula', formula: 'SUM(A2,1)' },
  D1: { type: 'formula', formula: 'A2*2', result: { valueType: 'number', numberValue: 6 } },
  E1: { type: 'formula', formula: '1/0', result: { valueType: 'error', value: '#DIV/0!' } },
  B2: { type: 'formula', formula: 'A2>1', result: { valueType: 'boolean', boolValue: true } },
  D2: { type: 'formula', formula: 'A3*2', result: { valueType: 'number', numberValue: 8 } },
} satisfies Record<string, FormulaValue>;

/** EXPECTED_FORMULAS keyed by cell reference. */
const EXPECTED_FORMULA_MAP = new Map<string, FormulaValue>(Object.entries(EXPECTED_FORMULAS));

type CellFormula = JsRowData['cells'][number]['formula'];

/** Collect `formula` by cell reference, checking that each matches the decoded `value`. */
function formulasByCell(rows: Iterable<JsRowData>): Map<string, CellFormula> {
  const found = new Map<string, CellFormula>();

  for (const row of rows) {
    for (const cell of row.cells) {
      if (cell.valueType !== 'formula') continue;
      expect(cell.formula?.formula).toBe(cell.value);
      found.set(`${cell.column}${row.row}`, cell.formula);
    }
  }

  return found;
}

describe('formula cell reads', () => {
  for (const readMode of ['lazy', 'eager'] as const) {
    it(`returns the expression and cached result on every read path (${readMode})`, async () => {
      const wb = Workbook.openBufferSync(workbookWithSheetXml(FORMULA_SHEET_XML), { readMode });

      for (const [cell, expected] of Object.entries(EXPECTED_FORMULAS)) {
        expect(wb.getCellValue('Sheet1', cell)).toEqual(expected);
      }

      expect(formulasByCell(wb.getRows('Sheet1'))).toEqual(EXPECTED_FORMULA_MAP);
      expect(formulasByCell(wb.getRows('Sheet1', { mode: 'dense' }))).toEqual(EXPECTED_FORMULA_MAP);
      expect(formulasByCell(wb.getRowsIterator('Sheet1'))).toEqual(EXPECTED_FORMULA_MAP);
      expect(formulasByCell(wb.getRowsIterator('Sheet1', { mode: 'dense' }))).toEqual(
        EXPECTED_FORMULA_MAP,
      );

      const columns = new Map<string, CellFormula>();

      for (const column of wb.getCols('Sheet1')) {
        for (const cell of column.cells) {
          if (cell.valueType === 'formula')
            columns.set(`${column.column}${cell.row}`, cell.formula);
        }
      }

      expect(columns).toEqual(EXPECTED_FORMULA_MAP);

      // The stream reader does not resolve shared-formula followers yet.
      const reader = await wb.openSheetReader('Sheet1');
      const streamed: JsRowData[] = [];

      for (let batch = await reader.next(); batch !== null; batch = await reader.next()) {
        streamed.push(...batch);
      }

      await reader.close();
      const streamedExpected = new Map(EXPECTED_FORMULA_MAP);
      streamedExpected.delete('D2');
      expect(formulasByCell(streamed)).toEqual(streamedExpected);
    });
  }

  it('skips the native formula lookup when no formula cells were decoded', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'plain');
    wb.setCellValue('Sheet1', 'B2', 2);
    const lookup = vi.spyOn(NativeWorkbook.prototype, 'getFormulaCells');

    try {
      wb.getRows('Sheet1');
      wb.getRows('Sheet1', { mode: 'dense' });
      expect([...wb.getRowsIterator('Sheet1')]).toHaveLength(2);
      expect(lookup).not.toHaveBeenCalled();

      wb.setCellValue('Sheet1', 'C3', { type: 'formula', formula: 'B2*2' });
      wb.getRows('Sheet1');
      expect([...wb.getRowsIterator('Sheet1')]).toHaveLength(3);
      expect(lookup).toHaveBeenCalledTimes(2);
    } finally {
      lookup.mockRestore();
    }
  });

  it('counts the formula lookup against the traversal output limit', () => {
    const wb = Workbook.openBufferSync(workbookWithSheetXml(FORMULA_SHEET_XML));
    const bufferBytes = wb.getRowsBufferV2('Sheet1').length;

    const failure = (() => {
      try {
        wb.getRows('Sheet1', { limits: { maxOutputBytes: bufferBytes } });
      } catch (error) {
        return error;
      }

      return null;
    })();

    expect(isResourceLimitError(failure)).toBe(true);
    expect(isResourceLimitError(failure) && failure.dimension).toBe('output_bytes');
    expect(wb.getRows('Sheet1', { limits: { maxOutputBytes: 65_536 } })).toHaveLength(2);
  });
});

import { inflateRawSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { Workbook } from '../index.js';

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
    await expect(readAll()).rejects.toThrow();
  });
});

import type { JsRowCell, JsRowData } from './binding.js';

const CELL_STRIDE = 9;
const TYPE_EMPTY = 0x00;
const TYPE_NUMBER = 0x01;
const TYPE_STRING = 0x02;
const TYPE_BOOL = 0x03;
const TYPE_DATE = 0x04;
const TYPE_ERROR = 0x05;
const TYPE_FORMULA = 0x06;
const TYPE_RICH_STRING = 0x07;
const FLAG_SPARSE = 0x01;
const HEADER_SIZE = 16;
const MAGIC = 0x534b5244;
const SPARSE_ENTRY_SIZE = 11;
const EMPTY_ROW_OFFSET = 0xffffffff;

const decoder = new TextDecoder();

interface BufferHeader {
  version: number;
  rowCount: number;
  colCount: number;
  flags: number;
  minCol: number;
  isSparse: boolean;
}

interface RowIndexEntry {
  rowNum: number;
  cellOffset: number;
}

function columnNumberToName(n: number): string {
  let name = '';
  while (n > 0) {
    n--;
    name = String.fromCharCode(65 + (n % 26)) + name;
    n = Math.floor(n / 26);
  }
  return name;
}

function readHeader(view: DataView): BufferHeader {
  const magic = view.getUint32(0, true);
  if (magic !== MAGIC) {
    throw new Error(`Invalid buffer: bad magic 0x${magic.toString(16)}`);
  }
  const version = view.getUint16(4, true);
  const rowCount = view.getUint32(6, true);
  const colCount = view.getUint16(10, true);
  const flags = view.getUint32(12, true);
  const minCol = flags >>> 16;
  const isSparse = (flags & FLAG_SPARSE) !== 0;
  return { version, rowCount, colCount, flags, minCol, isSparse };
}

function readRowIndex(
  view: DataView,
  rowCount: number,
): { entries: RowIndexEntry[]; endOffset: number } {
  const entries = new Array<RowIndexEntry>(rowCount);
  let offset = HEADER_SIZE;
  for (let i = 0; i < rowCount; i++) {
    entries[i] = {
      rowNum: view.getUint32(offset, true),
      cellOffset: view.getUint32(offset + 4, true),
    };
    offset += 8;
  }
  return { entries, endOffset: offset };
}

function readStringTable(
  buf: Buffer,
  view: DataView,
  startOffset: number,
): { strings: string[]; endOffset: number } {
  const count = view.getUint32(startOffset, true);
  const blobSize = view.getUint32(startOffset + 4, true);
  if (count === 0) {
    return { strings: [], endOffset: startOffset + 8 };
  }
  const offsetsStart = startOffset + 8;
  const blobStart = offsetsStart + count * 4;
  const strings = new Array<string>(count);
  for (let i = 0; i < count; i++) {
    const strOffset = view.getUint32(offsetsStart + i * 4, true);
    const strEnd = i + 1 < count ? view.getUint32(offsetsStart + (i + 1) * 4, true) : blobSize;
    const slice = new Uint8Array(
      buf.buffer,
      buf.byteOffset + blobStart + strOffset,
      strEnd - strOffset,
    );
    strings[i] = decoder.decode(slice);
  }
  const totalSize = 8 + count * 4 + blobSize;
  return { strings, endOffset: startOffset + totalSize };
}

function decodeCellToRowCell(
  view: DataView,
  offset: number,
  type: number,
  strings: string[],
  colName: string,
): JsRowCell | null {
  switch (type) {
    case TYPE_NUMBER: {
      const n = view.getFloat64(offset + 1, true);
      return { column: colName, valueType: 'number', numberValue: n };
    }
    case TYPE_STRING: {
      const idx = view.getUint32(offset + 1, true);
      return { column: colName, valueType: 'string', value: strings[idx] ?? '' };
    }
    case TYPE_BOOL: {
      const b = view.getUint8(offset + 1) !== 0;
      return { column: colName, valueType: 'boolean', boolValue: b };
    }
    case TYPE_DATE: {
      const serial = view.getFloat64(offset + 1, true);
      return { column: colName, valueType: 'date', numberValue: serial };
    }
    case TYPE_ERROR: {
      const idx = view.getUint32(offset + 1, true);
      return { column: colName, valueType: 'error', value: strings[idx] ?? '' };
    }
    case TYPE_FORMULA: {
      const idx = view.getUint32(offset + 1, true);
      return { column: colName, valueType: 'formula', value: strings[idx] ?? '' };
    }
    case TYPE_RICH_STRING: {
      const idx = view.getUint32(offset + 1, true);
      return { column: colName, valueType: 'string', value: strings[idx] ?? '' };
    }
    default:
      return null;
  }
}

export function decodeRowsBuffer(buf: Buffer | null): JsRowData[] {
  if (!buf || buf.length < HEADER_SIZE) {
    return [];
  }
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  const header = readHeader(view);
  if (header.rowCount === 0 || header.colCount === 0) {
    return [];
  }

  const { entries: rowIndex, endOffset: riEnd } = readRowIndex(view, header.rowCount);
  const { strings, endOffset: stEnd } = readStringTable(buf, view, riEnd);
  const cellDataStart = stEnd;
  const minCol = header.minCol || 1;

  const rows: JsRowData[] = [];

  for (let ri = 0; ri < rowIndex.length; ri++) {
    const { rowNum, cellOffset } = rowIndex[ri];
    if (cellOffset === EMPTY_ROW_OFFSET) {
      continue;
    }

    const cells: JsRowCell[] = [];

    if (header.isSparse) {
      const absOffset = cellDataStart + cellOffset;
      const cellCount = view.getUint16(absOffset, true);
      let pos = absOffset + 2;
      for (let i = 0; i < cellCount; i++) {
        const col = view.getUint16(pos, true);
        const type = view.getUint8(pos + 2);
        if (type !== TYPE_EMPTY) {
          const colName = columnNumberToName(minCol + col);
          const cell = decodeCellToRowCell(view, pos + 2, type, strings, colName);
          if (cell) {
            cells.push(cell);
          }
        }
        pos += SPARSE_ENTRY_SIZE;
      }
    } else {
      const rowStart = cellDataStart + cellOffset;
      for (let c = 0; c < header.colCount; c++) {
        const cellPos = rowStart + c * CELL_STRIDE;
        const type = view.getUint8(cellPos);
        if (type === TYPE_EMPTY) {
          continue;
        }
        const colName = columnNumberToName(minCol + c);
        const cell = decodeCellToRowCell(view, cellPos, type, strings, colName);
        if (cell) {
          cells.push(cell);
        }
      }
    }

    if (cells.length > 0) {
      rows.push({ row: rowNum, cells });
    }
  }

  return rows;
}

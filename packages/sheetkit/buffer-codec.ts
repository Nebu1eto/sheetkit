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

/** Result of getRowsRaw() -- typed arrays for minimal object creation. */
export interface RawRowsResult {
  /** 1-based row numbers for each row that has data. */
  rowNumbers: Uint32Array;
  /** Index into the cell arrays where each row's cells begin. */
  rowCellOffsets: Uint32Array;
  /** Number of cells per row. */
  rowCellCounts: Uint32Array;
  /** 1-based column number for each cell. */
  cellColumns: Uint32Array;
  /** Cell type tag (0=empty, 1=number, 2=string, 3=bool, 4=date, 5=error, 6=formula). */
  cellTypes: Uint8Array;
  /** Numeric value for number/date cells, NaN for others. */
  cellNumericValues: Float64Array;
  /** String value for string/error/formula cells, empty string for others. */
  cellStringValues: string[];
  /** Boolean value for boolean cells, false for others. */
  cellBoolValues: Uint8Array;
  /** Total number of rows with data. */
  totalRows: number;
  /** Total number of cells across all rows. */
  totalCells: number;
}

const colNameCache: string[] = [];

function cachedColumnName(n: number): string {
  if (n <= 0) return '';
  const idx = n - 1;
  if (idx < colNameCache.length) {
    return colNameCache[idx];
  }
  for (let i = colNameCache.length; i <= idx; i++) {
    let num = i + 1;
    let name = '';
    while (num > 0) {
      num--;
      name = String.fromCharCode(65 + (num % 26)) + name;
      num = Math.floor(num / 26);
    }
    colNameCache.push(name);
  }
  return colNameCache[idx];
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

/**
 * Parse buffer header and shared data structures used by all decode functions.
 * Returns null if the buffer is empty or invalid.
 */
function parseBuffer(buf: Buffer | null): {
  view: DataView;
  header: BufferHeader;
  rowIndex: RowIndexEntry[];
  strings: string[];
  cellDataStart: number;
  minCol: number;
} | null {
  if (!buf || buf.length < HEADER_SIZE) {
    return null;
  }
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  const header = readHeader(view);
  if (header.rowCount === 0 || header.colCount === 0) {
    return null;
  }
  const { entries: rowIndex, endOffset: riEnd } = readRowIndex(view, header.rowCount);
  const { strings, endOffset: stEnd } = readStringTable(buf, view, riEnd);
  const cellDataStart = stEnd;
  const minCol = header.minCol || 1;
  return { view, header, rowIndex, strings, cellDataStart, minCol };
}

/** Decode a binary buffer into JsRowData objects. */
export function decodeRowsBuffer(buf: Buffer | null): JsRowData[] {
  const parsed = parseBuffer(buf);
  if (!parsed) return [];
  const { view, header, rowIndex, strings, cellDataStart, minCol } = parsed;

  // Pre-warm the column name cache only for dense layout where every column
  // will be visited. For sparse layout, column names are resolved lazily per
  // encountered cell to avoid O(colCount) work on wide sparse sheets.
  if (!header.isSparse) {
    const maxCol = minCol + header.colCount - 1;
    cachedColumnName(maxCol);
  }

  const rows: JsRowData[] = [];

  for (let ri = 0; ri < rowIndex.length; ri++) {
    const entry = rowIndex[ri];
    if (entry.cellOffset === EMPTY_ROW_OFFSET) {
      continue;
    }

    const cells: JsRowCell[] = [];

    if (header.isSparse) {
      const absOffset = cellDataStart + entry.cellOffset;
      const cellCount = view.getUint16(absOffset, true);
      let pos = absOffset + 2;
      for (let i = 0; i < cellCount; i++) {
        const col = view.getUint16(pos, true);
        const type = view.getUint8(pos + 2);
        if (type !== TYPE_EMPTY) {
          const colName = cachedColumnName(minCol + col);
          const cell = decodeCellToRowCell(view, pos + 2, type, strings, colName);
          if (cell) {
            cells.push(cell);
          }
        }
        pos += SPARSE_ENTRY_SIZE;
      }
    } else {
      const rowStart = cellDataStart + entry.cellOffset;
      for (let c = 0; c < header.colCount; c++) {
        const cellPos = rowStart + c * CELL_STRIDE;
        const type = view.getUint8(cellPos);
        if (type === TYPE_EMPTY) {
          continue;
        }
        const colName = cachedColumnName(minCol + c);
        const cell = decodeCellToRowCell(view, cellPos, type, strings, colName);
        if (cell) {
          cells.push(cell);
        }
      }
    }

    if (cells.length > 0) {
      rows.push({ row: entry.rowNum, cells });
    }
  }

  return rows;
}

/**
 * Decode a binary buffer into typed arrays with minimal object allocation.
 * This avoids creating JsRowCell/JsRowData objects entirely.
 */
export function decodeRowsRawBuffer(buf: Buffer | null): RawRowsResult {
  const empty: RawRowsResult = {
    rowNumbers: new Uint32Array(0),
    rowCellOffsets: new Uint32Array(0),
    rowCellCounts: new Uint32Array(0),
    cellColumns: new Uint32Array(0),
    cellTypes: new Uint8Array(0),
    cellNumericValues: new Float64Array(0),
    cellStringValues: [],
    cellBoolValues: new Uint8Array(0),
    totalRows: 0,
    totalCells: 0,
  };

  const parsed = parseBuffer(buf);
  if (!parsed) return empty;
  const { view, header, rowIndex, strings, cellDataStart, minCol } = parsed;

  // First pass: count total non-empty rows and cells to pre-allocate
  let totalRows = 0;
  let totalCells = 0;

  for (let ri = 0; ri < rowIndex.length; ri++) {
    const entry = rowIndex[ri];
    if (entry.cellOffset === EMPTY_ROW_OFFSET) continue;

    let rowCells = 0;
    if (header.isSparse) {
      const absOffset = cellDataStart + entry.cellOffset;
      const cellCount = view.getUint16(absOffset, true);
      let pos = absOffset + 2;
      for (let i = 0; i < cellCount; i++) {
        const type = view.getUint8(pos + 2);
        if (type !== TYPE_EMPTY) rowCells++;
        pos += SPARSE_ENTRY_SIZE;
      }
    } else {
      const rowStart = cellDataStart + entry.cellOffset;
      for (let c = 0; c < header.colCount; c++) {
        const type = view.getUint8(rowStart + c * CELL_STRIDE);
        if (type !== TYPE_EMPTY) rowCells++;
      }
    }

    if (rowCells > 0) {
      totalRows++;
      totalCells += rowCells;
    }
  }

  if (totalRows === 0) return empty;

  // Allocate typed arrays
  const rowNumbers = new Uint32Array(totalRows);
  const rowCellOffsets = new Uint32Array(totalRows);
  const rowCellCounts = new Uint32Array(totalRows);
  const cellColumns = new Uint32Array(totalCells);
  const cellTypes = new Uint8Array(totalCells);
  const cellNumericValues = new Float64Array(totalCells);
  const cellStringValues = new Array<string>(totalCells);
  const cellBoolValues = new Uint8Array(totalCells);

  // Second pass: fill arrays
  let rowIdx = 0;
  let cellIdx = 0;

  for (let ri = 0; ri < rowIndex.length; ri++) {
    const entry = rowIndex[ri];
    if (entry.cellOffset === EMPTY_ROW_OFFSET) continue;

    const cellStart = cellIdx;

    if (header.isSparse) {
      const absOffset = cellDataStart + entry.cellOffset;
      const cellCount = view.getUint16(absOffset, true);
      let pos = absOffset + 2;
      for (let i = 0; i < cellCount; i++) {
        const col = view.getUint16(pos, true);
        const type = view.getUint8(pos + 2);
        if (type !== TYPE_EMPTY) {
          cellColumns[cellIdx] = minCol + col;
          cellTypes[cellIdx] = type;
          fillCellValue(
            view,
            pos + 3,
            type,
            strings,
            cellNumericValues,
            cellStringValues,
            cellBoolValues,
            cellIdx,
          );
          cellIdx++;
        }
        pos += SPARSE_ENTRY_SIZE;
      }
    } else {
      const rowStart = cellDataStart + entry.cellOffset;
      for (let c = 0; c < header.colCount; c++) {
        const cellPos = rowStart + c * CELL_STRIDE;
        const type = view.getUint8(cellPos);
        if (type === TYPE_EMPTY) continue;
        cellColumns[cellIdx] = minCol + c;
        cellTypes[cellIdx] = type;
        fillCellValue(
          view,
          cellPos + 1,
          type,
          strings,
          cellNumericValues,
          cellStringValues,
          cellBoolValues,
          cellIdx,
        );
        cellIdx++;
      }
    }

    const count = cellIdx - cellStart;
    if (count > 0) {
      rowNumbers[rowIdx] = entry.rowNum;
      rowCellOffsets[rowIdx] = cellStart;
      rowCellCounts[rowIdx] = count;
      rowIdx++;
    }
  }

  return {
    rowNumbers,
    rowCellOffsets,
    rowCellCounts,
    cellColumns,
    cellTypes,
    cellNumericValues,
    cellStringValues,
    cellBoolValues,
    totalRows,
    totalCells,
  };
}

function fillCellValue(
  view: DataView,
  payloadOffset: number,
  type: number,
  strings: string[],
  numericValues: Float64Array,
  stringValues: string[],
  boolValues: Uint8Array,
  idx: number,
): void {
  switch (type) {
    case TYPE_NUMBER:
    case TYPE_DATE:
      numericValues[idx] = view.getFloat64(payloadOffset, true);
      stringValues[idx] = '';
      break;
    case TYPE_STRING:
    case TYPE_RICH_STRING: {
      const si = view.getUint32(payloadOffset, true);
      numericValues[idx] = Number.NaN;
      stringValues[idx] = strings[si] ?? '';
      break;
    }
    case TYPE_BOOL:
      boolValues[idx] = view.getUint8(payloadOffset);
      numericValues[idx] = Number.NaN;
      stringValues[idx] = '';
      break;
    case TYPE_ERROR:
    case TYPE_FORMULA: {
      const si = view.getUint32(payloadOffset, true);
      numericValues[idx] = Number.NaN;
      stringValues[idx] = strings[si] ?? '';
      break;
    }
    default:
      numericValues[idx] = Number.NaN;
      stringValues[idx] = '';
      break;
  }
}

/**
 * Generator that yields one JsRowData at a time from the buffer, avoiding
 * materializing the entire result array at once.
 */
export function* decodeRowsIterator(buf: Buffer | null): Generator<JsRowData> {
  const parsed = parseBuffer(buf);
  if (!parsed) return;
  const { view, header, rowIndex, strings, cellDataStart, minCol } = parsed;

  // Pre-warm only for dense layout. For sparse layout, column names are
  // resolved lazily per cell to preserve streaming semantics and avoid
  // O(colCount) upfront work on wide sparse sheets.
  if (!header.isSparse) {
    const maxCol = minCol + header.colCount - 1;
    cachedColumnName(maxCol);
  }

  for (let ri = 0; ri < rowIndex.length; ri++) {
    const entry = rowIndex[ri];
    if (entry.cellOffset === EMPTY_ROW_OFFSET) continue;

    const cells: JsRowCell[] = [];

    if (header.isSparse) {
      const absOffset = cellDataStart + entry.cellOffset;
      const cellCount = view.getUint16(absOffset, true);
      let pos = absOffset + 2;
      for (let i = 0; i < cellCount; i++) {
        const col = view.getUint16(pos, true);
        const type = view.getUint8(pos + 2);
        if (type !== TYPE_EMPTY) {
          const colName = cachedColumnName(minCol + col);
          const cell = decodeCellToRowCell(view, pos + 2, type, strings, colName);
          if (cell) cells.push(cell);
        }
        pos += SPARSE_ENTRY_SIZE;
      }
    } else {
      const rowStart = cellDataStart + entry.cellOffset;
      for (let c = 0; c < header.colCount; c++) {
        const cellPos = rowStart + c * CELL_STRIDE;
        const type = view.getUint8(cellPos);
        if (type === TYPE_EMPTY) continue;
        const colName = cachedColumnName(minCol + c);
        const cell = decodeCellToRowCell(view, cellPos, type, strings, colName);
        if (cell) cells.push(cell);
      }
    }

    if (cells.length > 0) {
      yield { row: entry.rowNum, cells };
    }
  }
}

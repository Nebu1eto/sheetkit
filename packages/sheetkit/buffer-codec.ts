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
const ROW_INDEX_ENTRY_SIZE = 8;
const MAGIC_V1 = 0x534b5244;
const MAGIC_V2 = 0x534b5232;
const VERSION_V1 = 1;
const VERSION_V2 = 2;
const SPARSE_ENTRY_SIZE = 11;
const EMPTY_ROW_OFFSET = 0xffffffff;
const MAX_ROWS = 1_048_576;
const MAX_COLUMNS = 16_384;

const decoder = new TextDecoder('utf-8', { fatal: true });

function decodeUtf8(bytes: Uint8Array): string {
  try {
    return decoder.decode(bytes);
  } catch {
    throw new Error('Invalid buffer: malformed UTF-8 string');
  }
}

/**
 * Decode the UTF-8 bytes in `buf[start, end)`. All-ASCII input decodes to
 * the same string under Latin-1, so it skips the validating decoder; any
 * other byte goes through `decodeUtf8` with its malformed-input error.
 */
function decodeUtf8Range(buf: Buffer, start: number, end: number): string {
  for (let i = start; i < end; i++) {
    if (buf[i] > 0x7f) {
      return decodeUtf8(new Uint8Array(buf.buffer, buf.byteOffset + start, end - start));
    }
  }
  return buf.toString('latin1', start, end);
}

/**
 * Validate the UTF-8 bytes in `bytes[start, end)` without building a string.
 * All-ASCII input is valid UTF-8; any other byte goes through `decodeUtf8`
 * with its malformed-input error.
 */
function validateUtf8Range(bytes: Uint8Array, start: number, end: number): void {
  for (let i = start; i < end; i++) {
    if (bytes[i] > 0x7f) {
      decodeUtf8(bytes.subarray(start, end));
      return;
    }
  }
}

export interface ValidatedStringTable {
  readonly count: number;
  get(index: number): string;
}

class LazyValidatedStringTable implements ValidatedStringTable {
  readonly count: number;
  readonly #buf: Buffer;
  readonly #view: DataView;
  readonly #blobSize: number;
  readonly #offsetsStart: number;
  readonly #blobStart: number;
  readonly #cache = new Map<number, string>();

  constructor(
    buf: Buffer,
    view: DataView,
    count: number,
    blobSize: number,
    offsetsStart: number,
    blobStart: number,
  ) {
    this.#buf = buf;
    this.#view = view;
    this.count = count;
    this.#blobSize = blobSize;
    this.#offsetsStart = offsetsStart;
    this.#blobStart = blobStart;
  }

  get(index: number): string {
    const cached = this.#cache.get(index);
    if (cached !== undefined) return cached;
    if (!Number.isInteger(index) || index < 0 || index >= this.count) {
      throw new Error('Invalid buffer: string index is out of range');
    }
    const start = this.#view.getUint32(this.#offsetsStart + index * 4, true);
    const end =
      index + 1 < this.count
        ? this.#view.getUint32(this.#offsetsStart + (index + 1) * 4, true)
        : this.#blobSize;
    const value = decodeUtf8(
      new Uint8Array(this.#buf.buffer, this.#buf.byteOffset + this.#blobStart + start, end - start),
    );
    this.#cache.set(index, value);
    return value;
  }
}

const EMPTY_STRING_TABLE: ValidatedStringTable = {
  count: 0,
  get: () => {
    throw new Error('Invalid buffer: string index is out of range');
  },
};

export interface ValidatedBufferHeader {
  version: number;
  rowCount: number;
  colCount: number;
  flags: number;
  minCol: number;
  isSparse: boolean;
}

export interface ValidatedRowIndexEntry {
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

function detectMagic(view: DataView): number {
  return view.getUint32(0, true);
}

function isStringCellType(type: number): boolean {
  return (
    type === TYPE_STRING ||
    type === TYPE_ERROR ||
    type === TYPE_FORMULA ||
    type === TYPE_RICH_STRING
  );
}

function requireRange(length: number, offset: number, size: number, section: string): void {
  if (!Number.isSafeInteger(offset) || !Number.isSafeInteger(size) || offset < 0 || size < 0) {
    throw new Error(`Invalid buffer: ${section} offset overflow`);
  }
  const end = offset + size;
  if (!Number.isSafeInteger(end) || end > length) {
    throw new Error(`Invalid buffer: truncated ${section}`);
  }
}

function validateHeader(
  buf: Buffer,
  view: DataView,
  magic: number,
  materializeRowIndex = true,
): { header: ValidatedBufferHeader; rowIndex: ValidatedRowIndexEntry[]; cellDataStart: number } {
  const length = buf.length;
  const version = view.getUint16(4, true);
  const rowCount = view.getUint32(6, true);
  const colCount = view.getUint16(10, true);
  const flags = view.getUint32(12, true);
  const isV2 = magic === MAGIC_V2;
  const expectedVersion = isV2 ? VERSION_V2 : VERSION_V1;
  if (version !== expectedVersion)
    throw new Error(`Invalid buffer: unsupported version ${version}`);
  if (isV2 ? (flags & 0xffff) !== 0 : (flags & 0xfffe) !== 0) {
    throw new Error('Invalid buffer: invalid flags');
  }
  if ((rowCount === 0) !== (colCount === 0)) {
    throw new Error('Invalid buffer: empty dimensions are inconsistent');
  }
  if (rowCount === 0) {
    if (flags !== 0 || length !== HEADER_SIZE)
      throw new Error('Invalid buffer: malformed empty buffer');
    return {
      header: { version, rowCount, colCount, flags, minCol: 1, isSparse: isV2 },
      rowIndex: [],
      cellDataStart: HEADER_SIZE,
    };
  }

  const minCol = flags >>> 16 || 1;
  const maxCol = minCol + colCount - 1;
  if (minCol > MAX_COLUMNS || maxCol > MAX_COLUMNS) {
    throw new Error('Invalid buffer: column range is out of bounds');
  }
  const isSparse = isV2 || (flags & FLAG_SPARSE) !== 0;
  const rowIndexSize = rowCount * ROW_INDEX_ENTRY_SIZE;
  requireRange(length, HEADER_SIZE, rowIndexSize, 'row index');
  const cellDataStart = HEADER_SIZE + rowIndexSize;
  // Callers that read the validated index straight from the buffer skip
  // building one object per row; validation is identical either way.
  const entries: ValidatedRowIndexEntry[] = materializeRowIndex
    ? new Array<ValidatedRowIndexEntry>(rowCount)
    : [];
  // Build the result before the row loop: a large index is usually entered
  // through on-stack replacement, which records no feedback for code after
  // the loop and would deoptimize the next optimized call there.
  const result = {
    header: { version, rowCount, colCount, flags, minCol, isSparse },
    rowIndex: entries,
    cellDataStart,
  };
  let previousRow = 0;
  for (let i = 0; i < rowCount; i++) {
    const offset = HEADER_SIZE + i * ROW_INDEX_ENTRY_SIZE;
    const rowNum = view.getUint32(offset, true);
    if (rowNum === 0 || rowNum > MAX_ROWS) {
      throw new Error('Invalid buffer: row number is out of range');
    }
    if (rowNum <= previousRow) {
      throw new Error('Invalid buffer: row numbers must be strictly increasing');
    }
    previousRow = rowNum;
    if (materializeRowIndex) {
      entries[i] = { rowNum, cellOffset: view.getUint32(offset + 4, true) };
    }
  }
  return result;
}

function validateBuffer(
  buf: Buffer,
  view: DataView,
  magic: number,
  validateUtf8 = true,
): ValidatedBuffer {
  const length = buf.length;
  const isV2 = magic === MAGIC_V2;
  const { header, rowIndex, cellDataStart } = validateHeader(buf, view, magic);
  if (header.rowCount === 0) {
    return isV2
      ? {
          format: 'v2',
          buf,
          view,
          header,
          rowIndex,
          cellDataStart,
          minCol: 1,
          totalRows: 0,
          totalCells: 0,
        }
      : {
          format: 'v1',
          view,
          header,
          rowIndex,
          strings: EMPTY_STRING_TABLE,
          cellDataStart,
          minCol: 1,
          totalRows: 0,
          totalCells: 0,
        };
  }

  if (!isV2) {
    requireRange(length, cellDataStart, 8, 'string table header');
    const stringCount = view.getUint32(cellDataStart, true);
    const blobSize = view.getUint32(cellDataStart + 4, true);
    const offsetsStart = cellDataStart + 8;
    const offsetsSize = stringCount * 4;
    requireRange(length, offsetsStart, offsetsSize, 'string table offsets');
    const blobStart = offsetsStart + offsetsSize;
    requireRange(length, blobStart, blobSize, 'string table blob');
    let previous = 0;
    for (let i = 0; i < stringCount; i++) {
      const offset = view.getUint32(offsetsStart + i * 4, true);
      if (i === 0 && offset !== 0) {
        throw new Error('Invalid buffer: first string offset must be zero');
      }
      if (offset < previous || offset > blobSize) {
        throw new Error('Invalid buffer: string offsets must be monotonic and in bounds');
      }
      previous = offset;
    }
    if (stringCount === 0 && blobSize !== 0) {
      throw new Error('Invalid buffer: string blob has no index entries');
    }
    const dataStart = blobStart + blobSize;
    const strings = readValidatedStringTable(
      buf,
      view,
      stringCount,
      blobSize,
      offsetsStart,
      blobStart,
    );
    if (!header.isSparse) {
      const rowSize = header.colCount * CELL_STRIDE;
      requireRange(length, dataStart, header.rowCount * rowSize, 'dense cell data');
      let totalRows = 0;
      let totalCells = 0;
      for (let i = 0; i < header.rowCount; i++) {
        if (rowIndex[i].cellOffset !== EMPTY_ROW_OFFSET && rowIndex[i].cellOffset !== i * rowSize) {
          throw new Error('Invalid buffer: dense row offsets are inconsistent');
        }
        const includeRow = rowIndex[i].cellOffset !== EMPTY_ROW_OFFSET;
        let rowCells = 0;
        const rowStart = dataStart + i * rowSize;
        for (let column = 0; column < header.colCount; column++) {
          const offset = rowStart + column * CELL_STRIDE;
          validateV1Cell(view, offset, strings, validateUtf8);
          if (includeRow && view.getUint8(offset) !== TYPE_EMPTY) rowCells++;
        }
        if (rowCells > 0) {
          totalRows++;
          totalCells += rowCells;
        }
      }
      if (dataStart + header.rowCount * rowSize !== length)
        throw new Error('Invalid buffer: trailing data');
      return {
        format: 'v1',
        view,
        header,
        rowIndex,
        strings,
        cellDataStart: dataStart,
        minCol: header.minCol,
        totalRows,
        totalCells,
      };
    }
    const counts = validateSparseRows(
      view,
      length,
      rowIndex,
      dataStart,
      false,
      header.colCount,
      strings,
      validateUtf8,
    );
    return {
      format: 'v1',
      view,
      header,
      rowIndex,
      strings,
      cellDataStart: dataStart,
      minCol: header.minCol,
      ...counts,
    };
  }
  const counts = validateUtf8
    ? validateSparseRows(view, length, rowIndex, cellDataStart, true, header.colCount)
    : validateSparseRowsV2ForRaw(view, length, rowIndex, cellDataStart, header.colCount);
  return {
    format: 'v2',
    buf,
    view,
    header,
    rowIndex,
    cellDataStart,
    minCol: header.minCol,
    ...counts,
  };
}

function validateSparseRowsV2ForRaw(
  view: DataView,
  length: number,
  entries: ValidatedRowIndexEntry[],
  dataStart: number,
  colCount: number,
): { totalRows: number; totalCells: number } {
  let cursor = dataStart;
  let totalRows = 0;
  let totalCells = 0;
  for (let rowIndex = 0; rowIndex < entries.length; rowIndex++) {
    if (cursor > length - 2) throw new Error('Invalid buffer: truncated sparse row');
    const entry = entries[rowIndex];
    if (entry.cellOffset === EMPTY_ROW_OFFSET) {
      throw new Error('Invalid buffer: v2 row cannot use an empty-row sentinel');
    }
    if (entry.cellOffset !== cursor - dataStart) {
      throw new Error('Invalid buffer: sparse row offsets are inconsistent');
    }
    const cellCount = view.getUint16(cursor, true);
    cursor += 2;
    let previousColumn = -1;
    let rowCells = 0;

    for (let cellIndex = 0; cellIndex < cellCount; cellIndex++) {
      if (cursor > length - 3) throw new Error('Invalid buffer: truncated sparse cell');
      const column = view.getUint16(cursor, true);
      if (column >= colCount) throw new Error('Invalid buffer: cell column is out of range');
      if (column <= previousColumn) {
        throw new Error('Invalid buffer: cell columns must be strictly increasing');
      }
      previousColumn = column;
      const type = view.getUint8(cursor + 2);
      if (type > TYPE_RICH_STRING) {
        throw new Error(`Invalid buffer: unknown cell type ${type}`);
      }
      if (type !== TYPE_EMPTY) rowCells++;
      cursor += 3;

      let payloadSize = 0;
      if (type === TYPE_NUMBER || type === TYPE_DATE) {
        payloadSize = 8;
      } else if (type === TYPE_BOOL) {
        payloadSize = 1;
      } else if (type >= TYPE_STRING) {
        if (cursor > length - 4) {
          throw new Error('Invalid buffer: truncated v2 string length');
        }
        payloadSize = 4 + view.getUint32(cursor, true);
      }
      if (cursor > length - payloadSize) {
        throw new Error('Invalid buffer: truncated v2 cell payload');
      }
      if (type === TYPE_BOOL && view.getUint8(cursor) > 1) {
        throw new Error('Invalid buffer: boolean payload');
      }
      cursor += payloadSize;
    }

    if (rowCells > 0) {
      totalRows++;
      totalCells += rowCells;
    }
  }
  if (cursor !== length) throw new Error('Invalid buffer: trailing data');
  return { totalRows, totalCells };
}

function validateSparseRows(
  view: DataView,
  length: number,
  entries: ValidatedRowIndexEntry[],
  dataStart: number,
  isV2: boolean,
  colCount: number,
  strings?: ValidatedStringTable,
  validateUtf8 = true,
): { totalRows: number; totalCells: number } {
  let cursor = dataStart;
  let totalRows = 0;
  let totalCells = 0;
  const bytes = new Uint8Array(view.buffer, view.byteOffset, view.byteLength);
  for (const entry of entries) {
    if (cursor > length - 2) throw new Error('Invalid buffer: truncated sparse row');
    const cellCount = view.getUint16(cursor, true);
    if (entry.cellOffset === EMPTY_ROW_OFFSET) {
      if (isV2) throw new Error('Invalid buffer: v2 row cannot use an empty-row sentinel');
      if (cellCount !== 0) throw new Error('Invalid buffer: empty row has cells');
      cursor += 2;
      continue;
    }
    if (entry.cellOffset !== cursor - dataStart) {
      throw new Error('Invalid buffer: sparse row offsets are inconsistent');
    }
    cursor += 2;
    let previousColumn = -1;
    let rowCells = 0;
    for (let i = 0; i < cellCount; i++) {
      const entrySize = isV2 ? 3 : SPARSE_ENTRY_SIZE;
      if (cursor > length - entrySize) {
        throw new Error('Invalid buffer: truncated sparse cell');
      }
      const col = view.getUint16(cursor, true);
      if (col >= colCount) throw new Error('Invalid buffer: cell column is out of range');
      if (col <= previousColumn) {
        throw new Error('Invalid buffer: cell columns must be strictly increasing');
      }
      previousColumn = col;
      const type = view.getUint8(cursor + 2);
      if (type > TYPE_RICH_STRING) throw new Error(`Invalid buffer: unknown cell type ${type}`);
      if (type !== TYPE_EMPTY) rowCells++;
      if (!isV2) {
        if (!strings) throw new Error('Invalid buffer: missing string table');
        validateV1Cell(view, cursor + 2, strings, validateUtf8);
        cursor += SPARSE_ENTRY_SIZE;
        continue;
      }
      cursor += 3;
      let payloadSize = 0;
      if (type === TYPE_NUMBER || type === TYPE_DATE) payloadSize = 8;
      else if (type === TYPE_BOOL) {
        payloadSize = 1;
      } else if (type >= TYPE_STRING) {
        if (cursor > length - 4) {
          throw new Error('Invalid buffer: truncated v2 string length');
        }
        payloadSize = 4 + view.getUint32(cursor, true);
      }
      if (cursor > length - payloadSize) {
        throw new Error('Invalid buffer: truncated v2 cell payload');
      }
      if (type === TYPE_BOOL && view.getUint8(cursor) > 1) {
        throw new Error('Invalid buffer: boolean payload');
      }
      if (validateUtf8 && isStringCellType(type)) {
        const stringStart = cursor + 4;
        validateUtf8Range(bytes, stringStart, stringStart + view.getUint32(cursor, true));
      }
      cursor += payloadSize;
    }
    if (rowCells > 0) {
      totalRows++;
      totalCells += rowCells;
    }
  }
  if (cursor !== length) throw new Error('Invalid buffer: trailing data');
  return { totalRows, totalCells };
}

function readValidatedStringTable(
  buf: Buffer,
  view: DataView,
  count: number,
  blobSize: number,
  offsetsStart: number,
  blobStart: number,
): ValidatedStringTable {
  return new LazyValidatedStringTable(buf, view, count, blobSize, offsetsStart, blobStart);
}

function validateV1Cell(
  view: DataView,
  offset: number,
  strings: ValidatedStringTable,
  validateUtf8 = true,
): void {
  const type = view.getUint8(offset);
  if (type > TYPE_RICH_STRING) throw new Error(`Invalid buffer: unknown cell type ${type}`);
  const payloadOffset = offset + 1;
  if (type === TYPE_BOOL && view.getUint8(payloadOffset) > 1) {
    throw new Error('Invalid buffer: boolean payload');
  }
  if (isStringCellType(type) && view.getUint32(payloadOffset, true) >= strings.count) {
    throw new Error('Invalid buffer: string index is out of range');
  }
  if (validateUtf8 && isStringCellType(type)) {
    strings.get(view.getUint32(payloadOffset, true));
  }
}

function decodeCellToRowCell(
  view: DataView,
  offset: number,
  type: number,
  strings: ValidatedStringTable,
  colName: string,
): JsRowCell | null {
  switch (type) {
    case TYPE_NUMBER: {
      const n = view.getFloat64(offset + 1, true);
      return { column: colName, valueType: 'number', numberValue: n };
    }
    case TYPE_STRING: {
      const idx = view.getUint32(offset + 1, true);
      return { column: colName, valueType: 'string', value: strings.get(idx) };
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
      return { column: colName, valueType: 'error', value: strings.get(idx) };
    }
    case TYPE_FORMULA: {
      const idx = view.getUint32(offset + 1, true);
      return { column: colName, valueType: 'formula', value: strings.get(idx) };
    }
    case TYPE_RICH_STRING: {
      const idx = view.getUint32(offset + 1, true);
      return { column: colName, valueType: 'string', value: strings.get(idx) };
    }
    default:
      return null;
  }
}

/** Read inline string at the given offset. Returns the string and bytes consumed. */
function readInlineString(
  buf: Buffer,
  view: DataView,
  offset: number,
): { value: string; bytesConsumed: number } {
  const len = view.getUint32(offset, true);
  const slice = new Uint8Array(buf.buffer, buf.byteOffset + offset + 4, len);
  return { value: decodeUtf8(slice), bytesConsumed: 4 + len };
}

/** Compute the byte size of a v2 cell's payload (after the type tag). */
function v2PayloadSize(view: DataView, payloadOffset: number, type: number): number {
  switch (type) {
    case TYPE_EMPTY:
      return 0;
    case TYPE_NUMBER:
    case TYPE_DATE:
      return 8;
    case TYPE_BOOL:
      return 1;
    case TYPE_STRING:
    case TYPE_ERROR:
    case TYPE_FORMULA:
    case TYPE_RICH_STRING: {
      const len = view.getUint32(payloadOffset, true);
      return 4 + len;
    }
    default:
      return 0;
  }
}

/** Decode a v2 cell into a JsRowCell with inline string data. */
function decodeCellToRowCellV2(
  buf: Buffer,
  view: DataView,
  payloadOffset: number,
  type: number,
  colName: string,
): JsRowCell | null {
  switch (type) {
    case TYPE_NUMBER: {
      const n = view.getFloat64(payloadOffset, true);
      return { column: colName, valueType: 'number', numberValue: n };
    }
    case TYPE_STRING:
    case TYPE_RICH_STRING: {
      const { value } = readInlineString(buf, view, payloadOffset);
      return { column: colName, valueType: 'string', value };
    }
    case TYPE_BOOL: {
      const b = view.getUint8(payloadOffset) !== 0;
      return { column: colName, valueType: 'boolean', boolValue: b };
    }
    case TYPE_DATE: {
      const serial = view.getFloat64(payloadOffset, true);
      return { column: colName, valueType: 'date', numberValue: serial };
    }
    case TYPE_ERROR: {
      const { value } = readInlineString(buf, view, payloadOffset);
      return { column: colName, valueType: 'error', value };
    }
    case TYPE_FORMULA: {
      const { value } = readInlineString(buf, view, payloadOffset);
      return { column: colName, valueType: 'formula', value };
    }
    default:
      return null;
  }
}

/** Fill raw typed arrays for a single v2 cell. */
function fillCellValueV2(
  buf: Buffer,
  view: DataView,
  payloadOffset: number,
  type: number,
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
      const { value } = readInlineString(buf, view, payloadOffset);
      numericValues[idx] = Number.NaN;
      stringValues[idx] = value;
      break;
    }
    case TYPE_BOOL:
      boolValues[idx] = view.getUint8(payloadOffset);
      numericValues[idx] = Number.NaN;
      stringValues[idx] = '';
      break;
    case TYPE_ERROR:
    case TYPE_FORMULA: {
      const { value } = readInlineString(buf, view, payloadOffset);
      numericValues[idx] = Number.NaN;
      stringValues[idx] = value;
      break;
    }
    default:
      numericValues[idx] = Number.NaN;
      stringValues[idx] = '';
      break;
  }
}

export interface ValidatedBufferV1 {
  format: 'v1';
  view: DataView;
  header: ValidatedBufferHeader;
  rowIndex: ValidatedRowIndexEntry[];
  strings: ValidatedStringTable;
  cellDataStart: number;
  minCol: number;
  totalRows: number;
  totalCells: number;
}

export interface ValidatedBufferV2 {
  format: 'v2';
  buf: Buffer;
  view: DataView;
  header: ValidatedBufferHeader;
  rowIndex: ValidatedRowIndexEntry[];
  cellDataStart: number;
  minCol: number;
  totalRows: number;
  totalCells: number;
}

export type ValidatedBuffer = ValidatedBufferV1 | ValidatedBufferV2;

/**
 * Re-wrap a native buffer in a zero-copy JS view. Buffers created by the
 * binding share a prototype-transitioned hidden class that V8 may discard
 * once the previous buffer is collected, which deoptimizes decode loops
 * specialized on it; a Buffer constructed from JS keeps a stable map.
 */
function stableBufferView(buf: Buffer): Buffer {
  return Buffer.from(buf.buffer, buf.byteOffset, buf.byteLength);
}

function parseRowsBufferInternal(
  input: Buffer | null,
  validateUtf8: boolean,
): ValidatedBuffer | null {
  if (!input || input.length === 0) {
    return null;
  }
  const buf = stableBufferView(input);
  if (buf.length < HEADER_SIZE) throw new Error('Invalid buffer: truncated header');
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  const magic = detectMagic(view);

  if (magic !== MAGIC_V1 && magic !== MAGIC_V2) {
    throw new Error(`Invalid buffer: bad magic 0x${magic.toString(16)}`);
  }
  const parsed = validateBuffer(buf, view, magic, validateUtf8);
  return parsed.header.rowCount === 0 ? null : parsed;
}

/**
 * Parse buffer header and shared data structures used by all decode functions.
 * Returns null if the buffer is empty or invalid. Auto-detects v1 vs v2 format.
 */
export function parseRowsBuffer(buf: Buffer | null): ValidatedBuffer | null {
  return parseRowsBufferInternal(buf, true);
}

/** Decode a binary buffer (v1 or v2) into JsRowData objects. */
export function decodeRowsBuffer(input: Buffer | null): JsRowData[] {
  if (!input || input.length === 0) return [];
  const buf = stableBufferView(input);
  if (buf.length < HEADER_SIZE) throw new Error('Invalid buffer: truncated header');
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  const magic = detectMagic(view);
  if (magic === MAGIC_V2) {
    const { header, cellDataStart } = validateHeader(buf, view, magic, false);
    if (header.rowCount === 0) return [];
    return decodeAndValidateRowsBufferV2(buf, view, header, cellDataStart);
  }
  if (magic !== MAGIC_V1) {
    throw new Error(`Invalid buffer: bad magic 0x${magic.toString(16)}`);
  }
  const parsed = validateBuffer(buf, view, magic);
  if (parsed.header.rowCount === 0) return [];
  if (parsed.format !== 'v1') throw new Error('Invalid buffer: unexpected format');
  return decodeRowsBufferV1(parsed);
}

function decodeRowsBufferV1(parsed: ValidatedBufferV1): JsRowData[] {
  const { view, header, rowIndex, strings, cellDataStart, minCol } = parsed;

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
 * Validate and decode the `cellCount` v2 cells starting at `cursor` into
 * `cells`, returning the cursor after the row. Kept apart from the row loop so
 * this hot body is optimized as an ordinary function: V8 discards optimized
 * code whose allocation-site pretenuring decision flips (as it does between
 * calls that retain and then release millions of cells), and a small function
 * re-optimizes far sooner than the whole decoder can be re-entered via OSR.
 */
function decodeAndValidateRowCellsV2(
  buf: Buffer,
  view: DataView,
  length: number,
  colCount: number,
  columnNameOffset: number,
  start: number,
  cellCount: number,
  cells: JsRowCell[],
): number {
  let cursor = start;
  let previousColumn = -1;
  for (let i = 0; i < cellCount; i++) {
    if (cursor + 3 > length) throw new Error('Invalid buffer: truncated sparse cell');
    const column = view.getUint16(cursor, true);
    if (column >= colCount) {
      throw new Error('Invalid buffer: cell column is out of range');
    }
    if (column <= previousColumn) {
      throw new Error('Invalid buffer: cell columns must be strictly increasing');
    }
    previousColumn = column;
    const type = view.getUint8(cursor + 2);
    if (type > TYPE_RICH_STRING) {
      throw new Error(`Invalid buffer: unknown cell type ${type}`);
    }
    const payloadOffset = cursor + 3;
    let payloadSize = 0;
    if (type === TYPE_NUMBER || type === TYPE_DATE) payloadSize = 8;
    else if (type === TYPE_BOOL) payloadSize = 1;
    else if (isStringCellType(type)) {
      if (payloadOffset + 4 > length) throw new Error('Invalid buffer: truncated v2 string length');
      payloadSize = 4 + view.getUint32(payloadOffset, true);
    }
    if (payloadOffset + payloadSize > length) {
      throw new Error('Invalid buffer: truncated v2 cell payload');
    }
    const colName = colNameCache[columnNameOffset + column];

    if (type === TYPE_NUMBER) {
      cells.push({
        column: colName,
        valueType: 'number',
        numberValue: view.getFloat64(payloadOffset, true),
      });
    } else if (type === TYPE_DATE) {
      cells.push({
        column: colName,
        valueType: 'date',
        numberValue: view.getFloat64(payloadOffset, true),
      });
    } else if (type === TYPE_BOOL) {
      const value = view.getUint8(payloadOffset);
      if (value > 1) throw new Error('Invalid buffer: boolean payload');
      cells.push({ column: colName, valueType: 'boolean', boolValue: value !== 0 });
    } else if (isStringCellType(type)) {
      const stringLength = view.getUint32(payloadOffset, true);
      const stringStart = payloadOffset + 4;
      const value = decodeUtf8Range(buf, stringStart, stringStart + stringLength);
      const valueType =
        type === TYPE_ERROR ? 'error' : type === TYPE_FORMULA ? 'formula' : 'string';
      cells.push({ column: colName, valueType, value });
    }
    cursor = payloadOffset + payloadSize;
  }
  return cursor;
}

function decodeAndValidateRowsBufferV2(
  buf: Buffer,
  view: DataView,
  header: ValidatedBufferHeader,
  cellDataStart: number,
): JsRowData[] {
  const rows: JsRowData[] = [];
  const length = buf.length;
  const colCount = header.colCount;
  // Column names are bounded by the validated header range; warm the cache
  // once so each cell indexes it directly.
  const columnNameOffset = header.minCol - 1;
  cachedColumnName(header.minCol + colCount - 1);
  let cursor = cellDataStart;

  // Offsets below stay safe integers (cursor <= length, sizes < 2^33), so
  // each inline bound check is equivalent to requireRange().
  // The row index was validated by validateHeader(); read it in place.
  for (let ri = 0; ri < header.rowCount; ri++) {
    const indexOffset = HEADER_SIZE + ri * ROW_INDEX_ENTRY_SIZE;
    const rowNum = view.getUint32(indexOffset, true);
    const cellOffset = view.getUint32(indexOffset + 4, true);
    if (cursor + 2 > length) throw new Error('Invalid buffer: truncated sparse row');
    if (cellOffset === EMPTY_ROW_OFFSET) {
      throw new Error('Invalid buffer: v2 row cannot use an empty-row sentinel');
    }
    if (cellOffset !== cursor - cellDataStart) {
      throw new Error('Invalid buffer: sparse row offsets are inconsistent');
    }
    const cellCount = view.getUint16(cursor, true);
    const cells: JsRowCell[] = [];
    cursor = decodeAndValidateRowCellsV2(
      buf,
      view,
      length,
      colCount,
      columnNameOffset,
      cursor + 2,
      cellCount,
      cells,
    );
    if (cells.length > 0) {
      rows.push({ row: rowNum, cells });
    }
  }
  // Use the hoisted length: code after an OSR-entered loop has no feedback.
  if (cursor !== length) throw new Error('Invalid buffer: trailing data');
  return rows;
}

/**
 * Decode a binary buffer (v1 or v2) into typed arrays with minimal object allocation.
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

  const parsed = parseRowsBufferInternal(buf, false);
  if (!parsed) return empty;

  if (parsed.format === 'v2') {
    return decodeRowsRawBufferV2(parsed);
  }
  return decodeRowsRawBufferV1(parsed);
}

function decodeRowsRawBufferV1(parsed: ValidatedBufferV1): RawRowsResult {
  const { view, header, rowIndex, strings, cellDataStart, minCol, totalRows, totalCells } = parsed;

  if (totalRows === 0) {
    return {
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
  }

  const rowNumbers = new Uint32Array(totalRows);
  const rowCellOffsets = new Uint32Array(totalRows);
  const rowCellCounts = new Uint32Array(totalRows);
  const cellColumns = new Uint32Array(totalCells);
  const cellTypes = new Uint8Array(totalCells);
  const cellNumericValues = new Float64Array(totalCells);
  const cellStringValues = new Array<string>(totalCells);
  const cellBoolValues = new Uint8Array(totalCells);

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
          fillCellValueV1(
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
        fillCellValueV1(
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

function decodeRowsRawBufferV2(parsed: ValidatedBufferV2): RawRowsResult {
  const { buf, view, rowIndex, cellDataStart, minCol, totalRows, totalCells } = parsed;

  if (totalRows === 0) {
    return {
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
  }

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
    const absOffset = cellDataStart + entry.cellOffset;
    const cellCount = view.getUint16(absOffset, true);
    let pos = absOffset + 2;

    for (let i = 0; i < cellCount; i++) {
      const col = view.getUint16(pos, true);
      const type = view.getUint8(pos + 2);
      const payloadOffset = pos + 3;
      const payloadSize = v2PayloadSize(view, payloadOffset, type);

      if (type !== TYPE_EMPTY) {
        cellColumns[cellIdx] = minCol + col;
        cellTypes[cellIdx] = type;
        fillCellValueV2(
          buf,
          view,
          payloadOffset,
          type,
          cellNumericValues,
          cellStringValues,
          cellBoolValues,
          cellIdx,
        );
        cellIdx++;
      }

      pos = payloadOffset + payloadSize;
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

function fillCellValueV1(
  view: DataView,
  payloadOffset: number,
  type: number,
  strings: ValidatedStringTable,
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
      stringValues[idx] = strings.get(si);
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
      stringValues[idx] = strings.get(si);
      break;
    }
    default:
      numericValues[idx] = Number.NaN;
      stringValues[idx] = '';
      break;
  }
}

/**
 * Generator that yields one JsRowData at a time from the buffer (v1 or v2),
 * avoiding materializing the entire result array at once.
 */
export function* decodeRowsIterator(buf: Buffer | null): Generator<JsRowData> {
  const parsed = parseRowsBuffer(buf);
  if (!parsed) return;

  if (parsed.format === 'v2') {
    yield* decodeRowsIteratorV2(parsed);
    return;
  }
  yield* decodeRowsIteratorV1(parsed);
}

function* decodeRowsIteratorV1(parsed: ValidatedBufferV1): Generator<JsRowData> {
  const { view, header, rowIndex, strings, cellDataStart, minCol } = parsed;

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

function* decodeRowsIteratorV2(parsed: ValidatedBufferV2): Generator<JsRowData> {
  const { buf, view, rowIndex, cellDataStart, minCol } = parsed;

  for (let ri = 0; ri < rowIndex.length; ri++) {
    const entry = rowIndex[ri];
    if (entry.cellOffset === EMPTY_ROW_OFFSET) continue;

    const cells: JsRowCell[] = [];
    const absOffset = cellDataStart + entry.cellOffset;
    const cellCount = view.getUint16(absOffset, true);
    let pos = absOffset + 2;

    for (let i = 0; i < cellCount; i++) {
      const col = view.getUint16(pos, true);
      const type = view.getUint8(pos + 2);
      const payloadOffset = pos + 3;

      if (type !== TYPE_EMPTY) {
        const colName = cachedColumnName(minCol + col);
        const cell = decodeCellToRowCellV2(buf, view, payloadOffset, type, colName);
        if (cell) cells.push(cell);
      }

      pos = payloadOffset + v2PayloadSize(view, payloadOffset, type);
    }

    if (cells.length > 0) {
      yield { row: entry.rowNum, cells };
    }
  }
}

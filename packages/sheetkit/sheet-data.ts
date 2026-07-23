import assert from 'node:assert';
import {
  parseRowsBuffer,
  type ValidatedRowIndexEntry,
  type ValidatedStringTable,
} from './buffer-codec.js';

/** An Excel date value represented by its serial number. */
export interface DateValue {
  type: 'date';
  serial: number;
  iso: string | null;
}

export type CellValue = string | number | boolean | DateValue | null;
export type CellTypeName = 'empty' | 'number' | 'string' | 'boolean' | 'date' | 'error' | 'formula';

/** An occupied cell with its absolute 1-based worksheet column. */
export interface OccupiedCell {
  column: number;
  value: CellValue;
}

/** An occupied worksheet row with coordinate-bearing cells. */
export interface OccupiedRow {
  row: number;
  cells: OccupiedCell[];
}

export interface TypedOccupiedCell extends OccupiedCell {
  cellType: CellTypeName;
}

export interface TypedOccupiedRow {
  row: number;
  cells: TypedOccupiedCell[];
}

const CELL_STRIDE = 9;
const TYPE_EMPTY = 0x00;
const TYPE_NUMBER = 0x01;
const TYPE_STRING = 0x02;
const TYPE_BOOL = 0x03;
const TYPE_DATE = 0x04;
const TYPE_ERROR = 0x05;
const TYPE_FORMULA = 0x06;
const TYPE_RICH_STRING = 0x07;
const SPARSE_ENTRY_SIZE = 11;
const EMPTY_ROW_OFFSET = 0xffffffff;

const decoder = new TextDecoder('utf-8', { fatal: true });

function decodeUtf8(bytes: Uint8Array): string {
  try {
    return decoder.decode(bytes);
  } catch {
    throw new Error('Invalid buffer: malformed UTF-8 string');
  }
}
const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 31);
const MILLISECONDS_PER_DAY = 86_400_000;

function excelSerialToIso(serial: number): string | null {
  if (!Number.isFinite(serial)) return null;
  const serialDay = Math.floor(serial);
  if (serialDay < 1) return null;
  const adjustedDay = serialDay >= 60 ? serialDay - 1 : serialDay;
  const seconds = Math.round(Math.abs(serial - serialDay) * 86_400);
  if (seconds >= 86_400) return null;
  const timestamp = EXCEL_EPOCH_UTC + adjustedDay * MILLISECONDS_PER_DAY + seconds * 1_000;
  const date = new Date(timestamp);
  if (!Number.isFinite(date.getTime())) return null;
  const iso = date.toISOString();
  return serial % 1 === 0 ? iso.slice(0, 10) : iso.slice(0, 19);
}

const TYPE_NAMES: CellTypeName[] = [
  'empty',
  'number',
  'string',
  'boolean',
  'date',
  'error',
  'formula',
  'string',
];

function columnNumberToName(n: number): string {
  let name = '';
  while (n > 0) {
    n--;
    name = String.fromCharCode(65 + (n % 26)) + name;
    n = Math.floor(n / 26);
  }
  return name;
}

const EMPTY_VIEW = new DataView(new ArrayBuffer(0));

export class SheetData {
  #buf: Buffer | null;
  #view: DataView;
  #strings: ValidatedStringTable | null;
  #rowCount: number;
  #colCount: number;
  #minRow: number;
  #maxRow: number;
  #minCol: number;
  #isSparse: boolean;
  #isV2: boolean;
  #rowIndex: ValidatedRowIndexEntry[];
  #cellDataStart: number;

  constructor(buffer: Buffer | null) {
    this.#buf = buffer;
    const parsed = parseRowsBuffer(buffer);
    if (!parsed) {
      this.#rowCount = 0;
      this.#colCount = 0;
      this.#minRow = 0;
      this.#maxRow = 0;
      this.#minCol = 1;
      this.#isSparse = false;
      this.#isV2 = false;
      this.#rowIndex = [];
      this.#strings = null;
      this.#cellDataStart = 0;
      this.#view = EMPTY_VIEW;
      return;
    }

    this.#view = parsed.view;
    this.#rowCount = parsed.header.rowCount;
    this.#colCount = parsed.header.colCount;
    this.#minRow = parsed.rowIndex[0].rowNum;
    this.#maxRow = parsed.rowIndex[parsed.rowIndex.length - 1].rowNum;
    this.#minCol = parsed.minCol;
    this.#isSparse = parsed.header.isSparse;
    this.#isV2 = parsed.format === 'v2';
    this.#rowIndex = parsed.rowIndex;
    this.#strings = parsed.format === 'v1' ? parsed.strings : null;
    this.#cellDataStart = parsed.cellDataStart;

    if (!this.#isSparse) this.#rowCount = this.#denseOccupiedRowCount();
  }

  /** Number of occupied row-index entries represented by this buffer. */
  get rowCount(): number {
    return this.#rowCount;
  }

  get colCount(): number {
    return this.#colCount;
  }

  /** First absolute worksheet row represented by this buffer, or 0 when empty. */
  get minRow(): number {
    return this.#minRow;
  }

  /** Last absolute worksheet row represented by this buffer, or 0 when empty. */
  get maxRow(): number {
    return this.#maxRow;
  }

  /** Inclusive distance between the first and last represented rows. */
  get rowSpan(): number {
    return this.#rowCount === 0 ? 0 : this.#maxRow - this.#minRow + 1;
  }

  #denseOccupiedRowCount(): number {
    let occupied = 0;
    for (const { cellOffset } of this.#rowIndex) {
      if (cellOffset === EMPTY_ROW_OFFSET) continue;
      const rowStart = this.#cellDataStart + cellOffset;
      for (let column = 0; column < this.#colCount; column++) {
        if (this.#view.getUint8(rowStart + column * CELL_STRIDE) !== TYPE_EMPTY) {
          occupied++;
          break;
        }
      }
    }
    return occupied;
  }

  /** 1-based worksheet column represented by the first buffered column. */
  get minCol(): number {
    return this.#minCol;
  }

  #decodeCellValueV1(type: number, payloadOffset: number): CellValue {
    switch (type) {
      case TYPE_NUMBER:
        return this.#view.getFloat64(payloadOffset, true);
      case TYPE_DATE: {
        const serial = this.#view.getFloat64(payloadOffset, true);
        return { type: 'date', serial, iso: excelSerialToIso(serial) };
      }
      case TYPE_STRING:
      case TYPE_RICH_STRING: {
        const idx = this.#view.getUint32(payloadOffset, true);
        assert(this.#strings != null, 'String table must be available for v1 data');
        return this.#strings.get(idx);
      }
      case TYPE_BOOL:
        return this.#view.getUint8(payloadOffset) !== 0;
      case TYPE_ERROR: {
        const idx = this.#view.getUint32(payloadOffset, true);
        assert(this.#strings != null, 'String table must be available for v1 data');
        return this.#strings.get(idx);
      }
      case TYPE_FORMULA: {
        const idx = this.#view.getUint32(payloadOffset, true);
        assert(this.#strings != null, 'String table must be available for v1 data');
        return this.#strings.get(idx);
      }
      default:
        return null;
    }
  }

  #readInlineString(offset: number): { value: string; bytesConsumed: number } {
    const len = this.#view.getUint32(offset, true);
    assert(this.#buf != null, 'Buffer must not be null when reading inline strings');
    const slice = new Uint8Array(this.#buf.buffer, this.#buf.byteOffset + offset + 4, len);
    return { value: decodeUtf8(slice), bytesConsumed: 4 + len };
  }

  #decodeCellValueV2(type: number, payloadOffset: number): CellValue {
    switch (type) {
      case TYPE_NUMBER:
        return this.#view.getFloat64(payloadOffset, true);
      case TYPE_DATE: {
        const serial = this.#view.getFloat64(payloadOffset, true);
        return { type: 'date', serial, iso: excelSerialToIso(serial) };
      }
      case TYPE_STRING:
      case TYPE_RICH_STRING:
        return this.#readInlineString(payloadOffset).value;
      case TYPE_BOOL:
        return this.#view.getUint8(payloadOffset) !== 0;
      case TYPE_ERROR:
        return this.#readInlineString(payloadOffset).value;
      case TYPE_FORMULA:
        return this.#readInlineString(payloadOffset).value;
      default:
        return null;
    }
  }

  #v2PayloadSize(payloadOffset: number, type: number): number {
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
        const len = this.#view.getUint32(payloadOffset, true);
        return 4 + len;
      }
      default:
        return 0;
    }
  }

  #findRowIndex(rowNum: number): number {
    let low = 0;
    let high = this.#rowIndex.length - 1;
    while (low <= high) {
      const middle = low + Math.floor((high - low) / 2);
      const candidate = this.#rowIndex[middle].rowNum;
      if (candidate === rowNum) return middle;
      if (candidate < rowNum) low = middle + 1;
      else high = middle - 1;
    }
    return -1;
  }

  getCell(row: number, col: number): CellValue {
    if (this.#rowCount === 0) return null;
    const ri = this.#findRowIndex(row);
    if (ri === -1) return null;
    const { cellOffset } = this.#rowIndex[ri];
    if (cellOffset === EMPTY_ROW_OFFSET) return null;
    const colIdx = col - this.#minCol;
    if (colIdx < 0 || colIdx >= this.#colCount) return null;

    if (this.#isV2) {
      return this.#getCellV2(cellOffset, colIdx);
    }

    if (this.#isSparse) {
      const absOff = this.#cellDataStart + cellOffset;
      const cellCount = this.#view.getUint16(absOff, true);
      let pos = absOff + 2;
      for (let i = 0; i < cellCount; i++) {
        const c = this.#view.getUint16(pos, true);
        if (c === colIdx) {
          const type = this.#view.getUint8(pos + 2);
          if (type === TYPE_EMPTY) return null;
          return this.#decodeCellValueV1(type, pos + 3);
        }
        pos += SPARSE_ENTRY_SIZE;
      }
      return null;
    }

    const cellPos = this.#cellDataStart + cellOffset + colIdx * CELL_STRIDE;
    const type = this.#view.getUint8(cellPos);
    if (type === TYPE_EMPTY) return null;
    return this.#decodeCellValueV1(type, cellPos + 1);
  }

  #getCellV2(cellOffset: number, colIdx: number): CellValue {
    const absOff = this.#cellDataStart + cellOffset;
    const cellCount = this.#view.getUint16(absOff, true);
    let pos = absOff + 2;
    for (let i = 0; i < cellCount; i++) {
      const c = this.#view.getUint16(pos, true);
      const type = this.#view.getUint8(pos + 2);
      const payloadOffset = pos + 3;
      if (c === colIdx) {
        if (type === TYPE_EMPTY) return null;
        return this.#decodeCellValueV2(type, payloadOffset);
      }
      pos = payloadOffset + this.#v2PayloadSize(payloadOffset, type);
    }
    return null;
  }

  getCellType(row: number, col: number): CellTypeName {
    if (this.#rowCount === 0) return 'empty';
    const ri = this.#findRowIndex(row);
    if (ri === -1) return 'empty';
    const { cellOffset } = this.#rowIndex[ri];
    if (cellOffset === EMPTY_ROW_OFFSET) return 'empty';
    const colIdx = col - this.#minCol;
    if (colIdx < 0 || colIdx >= this.#colCount) return 'empty';

    if (this.#isV2) {
      return this.#getCellTypeV2(cellOffset, colIdx);
    }

    if (this.#isSparse) {
      const absOff = this.#cellDataStart + cellOffset;
      const cellCount = this.#view.getUint16(absOff, true);
      let pos = absOff + 2;
      for (let i = 0; i < cellCount; i++) {
        const c = this.#view.getUint16(pos, true);
        if (c === colIdx) {
          const type = this.#view.getUint8(pos + 2);
          return TYPE_NAMES[type] ?? 'empty';
        }
        pos += SPARSE_ENTRY_SIZE;
      }
      return 'empty';
    }

    const cellPos = this.#cellDataStart + cellOffset + colIdx * CELL_STRIDE;
    const type = this.#view.getUint8(cellPos);
    return TYPE_NAMES[type] ?? 'empty';
  }

  #getCellTypeV2(cellOffset: number, colIdx: number): CellTypeName {
    const absOff = this.#cellDataStart + cellOffset;
    const cellCount = this.#view.getUint16(absOff, true);
    let pos = absOff + 2;
    for (let i = 0; i < cellCount; i++) {
      const c = this.#view.getUint16(pos, true);
      const type = this.#view.getUint8(pos + 2);
      const payloadOffset = pos + 3;
      if (c === colIdx) {
        return TYPE_NAMES[type] ?? 'empty';
      }
      pos = payloadOffset + this.#v2PayloadSize(payloadOffset, type);
    }
    return 'empty';
  }

  getRow(rowNum: number): CellValue[] {
    if (this.#rowCount === 0) return [];
    const ri = this.#findRowIndex(rowNum);
    if (ri === -1) return [];
    const { cellOffset } = this.#rowIndex[ri];
    if (cellOffset === EMPTY_ROW_OFFSET) return [];

    if (this.#isSparse) return this.#getOccupiedCells(cellOffset).map(({ value }) => value);

    const result: CellValue[] = [];

    for (let c = 0; c < this.#colCount; c++) {
      const cellPos = this.#cellDataStart + cellOffset + c * CELL_STRIDE;
      const type = this.#view.getUint8(cellPos);
      if (type === TYPE_EMPTY) {
        result.push(null);
      } else {
        result.push(this.#decodeCellValueV1(type, cellPos + 1));
      }
    }
    return result;
  }

  #getOccupiedCells(cellOffset: number): OccupiedCell[] {
    if (!this.#isSparse) {
      const cells: OccupiedCell[] = [];
      const rowStart = this.#cellDataStart + cellOffset;
      for (let column = 0; column < this.#colCount; column++) {
        const cellPos = rowStart + column * CELL_STRIDE;
        const type = this.#view.getUint8(cellPos);
        if (type === TYPE_EMPTY) continue;
        cells.push({
          column: this.#minCol + column,
          value: this.#decodeCellValueV1(type, cellPos + 1),
        });
      }
      return cells;
    }

    return this.#isV2 ? this.#getOccupiedCellsV2(cellOffset) : this.#getOccupiedCellsV1(cellOffset);
  }

  #getOccupiedCellsV1(cellOffset: number): OccupiedCell[] {
    const cells: OccupiedCell[] = [];
    const absOff = this.#cellDataStart + cellOffset;
    const cellCount = this.#view.getUint16(absOff, true);
    let pos = absOff + 2;
    for (let i = 0; i < cellCount; i++) {
      const column = this.#view.getUint16(pos, true);
      const type = this.#view.getUint8(pos + 2);
      if (type !== TYPE_EMPTY) {
        cells.push({
          column: this.#minCol + column,
          value: this.#decodeCellValueV1(type, pos + 3),
        });
      }
      pos += SPARSE_ENTRY_SIZE;
    }
    return cells;
  }

  #getOccupiedCellsV2(cellOffset: number): OccupiedCell[] {
    const cells: OccupiedCell[] = [];
    const absOff = this.#cellDataStart + cellOffset;
    const cellCount = this.#view.getUint16(absOff, true);
    let pos = absOff + 2;
    for (let i = 0; i < cellCount; i++) {
      const column = this.#view.getUint16(pos, true);
      const type = this.#view.getUint8(pos + 2);
      const payloadOffset = pos + 3;
      if (type !== TYPE_EMPTY) {
        cells.push({
          column: this.#minCol + column,
          value: this.#decodeCellValueV2(type, payloadOffset),
        });
      }
      pos = payloadOffset + this.#v2PayloadSize(payloadOffset, type);
    }
    return cells;
  }

  #getTypedOccupiedCells(cellOffset: number): TypedOccupiedCell[] {
    if (!this.#isSparse) {
      const cells: TypedOccupiedCell[] = [];
      const rowStart = this.#cellDataStart + cellOffset;
      for (let column = 0; column < this.#colCount; column++) {
        const cellPos = rowStart + column * CELL_STRIDE;
        const type = this.#view.getUint8(cellPos);
        if (type === TYPE_EMPTY) continue;
        cells.push({
          column: this.#minCol + column,
          value: this.#decodeCellValueV1(type, cellPos + 1),
          cellType: TYPE_NAMES[type] ?? 'empty',
        });
      }
      return cells;
    }

    const cells: TypedOccupiedCell[] = [];
    const absOff = this.#cellDataStart + cellOffset;
    const cellCount = this.#view.getUint16(absOff, true);
    let pos = absOff + 2;
    for (let i = 0; i < cellCount; i++) {
      const column = this.#view.getUint16(pos, true);
      const type = this.#view.getUint8(pos + 2);
      const payloadOffset = pos + 3;
      if (type !== TYPE_EMPTY) {
        cells.push({
          column: this.#minCol + column,
          value: this.#isV2
            ? this.#decodeCellValueV2(type, payloadOffset)
            : this.#decodeCellValueV1(type, payloadOffset),
          cellType: TYPE_NAMES[type] ?? 'empty',
        });
      }
      pos = this.#isV2
        ? payloadOffset + this.#v2PayloadSize(payloadOffset, type)
        : pos + SPARSE_ENTRY_SIZE;
    }
    return cells;
  }

  toArray(): CellValue[][] {
    const result: CellValue[][] = [];
    for (const { rowNum, cellOffset } of this.#rowIndex) {
      if (cellOffset === EMPTY_ROW_OFFSET) {
        result.push([]);
        continue;
      }
      result.push(this.getRow(rowNum));
    }
    return result;
  }

  *rows(): Generator<{ row: number; values: CellValue[] }> {
    for (const { rowNum, cellOffset } of this.#rowIndex) {
      if (cellOffset === EMPTY_ROW_OFFSET) {
        yield { row: rowNum, values: [] };
        continue;
      }
      yield { row: rowNum, values: this.getRow(rowNum) };
    }
  }

  /** Iterate occupied rows without allocating arrays for coordinate gaps. */
  *occupiedRows(): Generator<OccupiedRow> {
    for (const { rowNum, cellOffset } of this.#rowIndex) {
      if (cellOffset === EMPTY_ROW_OFFSET) continue;
      const cells = this.#getOccupiedCells(cellOffset);
      if (cells.length > 0) yield { row: rowNum, cells };
    }
  }

  /** Iterate occupied rows with cell types without repeating coordinate lookups. */
  *typedOccupiedRows(): Generator<TypedOccupiedRow> {
    for (const { rowNum, cellOffset } of this.#rowIndex) {
      if (cellOffset === EMPTY_ROW_OFFSET) continue;
      const cells = this.#getTypedOccupiedCells(cellOffset);
      if (cells.length > 0) yield { row: rowNum, cells };
    }
  }

  columnName(colIndex: number): string {
    return columnNumberToName(this.#minCol + colIndex);
  }
}

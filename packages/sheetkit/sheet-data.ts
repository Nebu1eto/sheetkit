export type CellValue = string | number | boolean | null;
export type CellTypeName = 'empty' | 'number' | 'string' | 'boolean' | 'date' | 'error' | 'formula';

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
const MAGIC_V1 = 0x534b5244;
const MAGIC_V2 = 0x534b5232;
const SPARSE_ENTRY_SIZE = 11;
const EMPTY_ROW_OFFSET = 0xffffffff;

const decoder = new TextDecoder();

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

const EMPTY_VIEW = new DataView(new ArrayBuffer(0));

export class SheetData {
  #buf: Buffer | null;
  #view: DataView;
  #strings: string[];
  #rowCount: number;
  #colCount: number;
  #minCol: number;
  #isSparse: boolean;
  #isV2: boolean;
  #rowIndex: RowIndexEntry[];
  #cellDataStart: number;

  constructor(buffer: Buffer | null) {
    this.#buf = buffer;
    if (!buffer || buffer.length < HEADER_SIZE) {
      this.#rowCount = 0;
      this.#colCount = 0;
      this.#minCol = 1;
      this.#isSparse = false;
      this.#isV2 = false;
      this.#rowIndex = [];
      this.#strings = [];
      this.#cellDataStart = 0;
      this.#view = EMPTY_VIEW;
      return;
    }

    this.#view = new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength);
    const magic = this.#view.getUint32(0, true);
    if (magic !== MAGIC_V1 && magic !== MAGIC_V2) {
      throw new Error(`Invalid buffer: bad magic 0x${magic.toString(16)}`);
    }
    this.#isV2 = magic === MAGIC_V2;

    this.#rowCount = this.#view.getUint32(6, true);
    this.#colCount = this.#view.getUint16(10, true);
    const flags = this.#view.getUint32(12, true);
    this.#minCol = flags >>> 16 || 1;
    this.#isSparse = this.#isV2 ? true : (flags & FLAG_SPARSE) !== 0;

    if (this.#rowCount === 0 || this.#colCount === 0) {
      this.#rowIndex = [];
      this.#strings = [];
      this.#cellDataStart = HEADER_SIZE;
      return;
    }

    let offset = HEADER_SIZE;
    this.#rowIndex = new Array<RowIndexEntry>(this.#rowCount);
    for (let i = 0; i < this.#rowCount; i++) {
      this.#rowIndex[i] = {
        rowNum: this.#view.getUint32(offset, true),
        cellOffset: this.#view.getUint32(offset + 4, true),
      };
      offset += 8;
    }

    if (this.#isV2) {
      // v2 has no string table -- cell data starts immediately after row index
      this.#strings = [];
      this.#cellDataStart = offset;
    } else {
      const strCount = this.#view.getUint32(offset, true);
      const blobSize = this.#view.getUint32(offset + 4, true);
      if (strCount === 0) {
        this.#strings = [];
        this.#cellDataStart = offset + 8;
      } else {
        const offsetsStart = offset + 8;
        const blobStart = offsetsStart + strCount * 4;
        this.#strings = new Array<string>(strCount);
        for (let i = 0; i < strCount; i++) {
          const sOff = this.#view.getUint32(offsetsStart + i * 4, true);
          const sEnd =
            i + 1 < strCount ? this.#view.getUint32(offsetsStart + (i + 1) * 4, true) : blobSize;
          const slice = new Uint8Array(
            buffer.buffer,
            buffer.byteOffset + blobStart + sOff,
            sEnd - sOff,
          );
          this.#strings[i] = decoder.decode(slice);
        }
        this.#cellDataStart = blobStart + blobSize;
      }
    }
  }

  get rowCount(): number {
    return this.#rowCount;
  }

  get colCount(): number {
    return this.#colCount;
  }

  #decodeCellValueV1(type: number, payloadOffset: number): CellValue {
    switch (type) {
      case TYPE_NUMBER:
      case TYPE_DATE:
        return this.#view.getFloat64(payloadOffset, true);
      case TYPE_STRING:
      case TYPE_RICH_STRING: {
        const idx = this.#view.getUint32(payloadOffset, true);
        return this.#strings[idx] ?? '';
      }
      case TYPE_BOOL:
        return this.#view.getUint8(payloadOffset) !== 0;
      case TYPE_ERROR: {
        const idx = this.#view.getUint32(payloadOffset, true);
        return this.#strings[idx] ?? '';
      }
      case TYPE_FORMULA: {
        const idx = this.#view.getUint32(payloadOffset, true);
        return this.#strings[idx] ?? '';
      }
      default:
        return null;
    }
  }

  #readInlineString(offset: number): { value: string; bytesConsumed: number } {
    const len = this.#view.getUint32(offset, true);
    const slice = new Uint8Array(this.#buf!.buffer, this.#buf!.byteOffset + offset + 4, len);
    return { value: decoder.decode(slice), bytesConsumed: 4 + len };
  }

  #decodeCellValueV2(type: number, payloadOffset: number): CellValue {
    switch (type) {
      case TYPE_NUMBER:
      case TYPE_DATE:
        return this.#view.getFloat64(payloadOffset, true);
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
    for (let i = 0; i < this.#rowIndex.length; i++) {
      if (this.#rowIndex[i].rowNum === rowNum) {
        return i;
      }
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

    if (this.#isV2) {
      return this.#getRowV2(cellOffset);
    }

    const result: CellValue[] = [];

    if (this.#isSparse) {
      const absOff = this.#cellDataStart + cellOffset;
      const cellCount = this.#view.getUint16(absOff, true);
      let prevCol = -1;
      let pos = absOff + 2;
      for (let i = 0; i < cellCount; i++) {
        const col = this.#view.getUint16(pos, true);
        for (let c = prevCol + 1; c < col; c++) {
          result.push(null);
        }
        const type = this.#view.getUint8(pos + 2);
        if (type === TYPE_EMPTY) {
          result.push(null);
        } else {
          result.push(this.#decodeCellValueV1(type, pos + 3));
        }
        prevCol = col;
        pos += SPARSE_ENTRY_SIZE;
      }
      return result;
    }

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

  #getRowV2(cellOffset: number): CellValue[] {
    const result: CellValue[] = [];
    const absOff = this.#cellDataStart + cellOffset;
    const cellCount = this.#view.getUint16(absOff, true);
    let prevCol = -1;
    let pos = absOff + 2;
    for (let i = 0; i < cellCount; i++) {
      const col = this.#view.getUint16(pos, true);
      const type = this.#view.getUint8(pos + 2);
      const payloadOffset = pos + 3;
      for (let c = prevCol + 1; c < col; c++) {
        result.push(null);
      }
      if (type === TYPE_EMPTY) {
        result.push(null);
      } else {
        result.push(this.#decodeCellValueV2(type, payloadOffset));
      }
      prevCol = col;
      pos = payloadOffset + this.#v2PayloadSize(payloadOffset, type);
    }
    return result;
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

  columnName(colIndex: number): string {
    return columnNumberToName(this.#minCol + colIndex);
  }
}

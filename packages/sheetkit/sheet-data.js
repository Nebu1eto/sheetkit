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

const TYPE_NAMES = ['empty', 'number', 'string', 'boolean', 'date', 'error', 'formula', 'string'];

function columnNumberToName(n) {
  let name = '';
  while (n > 0) {
    n--;
    name = String.fromCharCode(65 + (n % 26)) + name;
    n = Math.floor(n / 26);
  }
  return name;
}

export class SheetData {
  #view;
  #strings;
  #rowCount;
  #colCount;
  #minCol;
  #isSparse;
  #rowIndex;
  #cellDataStart;

  constructor(buffer) {
    if (!buffer || buffer.length < HEADER_SIZE) {
      this.#rowCount = 0;
      this.#colCount = 0;
      this.#minCol = 1;
      this.#isSparse = false;
      this.#rowIndex = [];
      this.#strings = [];
      this.#cellDataStart = 0;
      this.#view = null;
      return;
    }

    this.#view = new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength);
    const magic = this.#view.getUint32(0, true);
    if (magic !== MAGIC) {
      throw new Error(`Invalid buffer: bad magic 0x${magic.toString(16)}`);
    }

    this.#rowCount = this.#view.getUint32(6, true);
    this.#colCount = this.#view.getUint16(10, true);
    const flags = this.#view.getUint32(12, true);
    this.#minCol = flags >>> 16 || 1;
    this.#isSparse = (flags & FLAG_SPARSE) !== 0;

    if (this.#rowCount === 0 || this.#colCount === 0) {
      this.#rowIndex = [];
      this.#strings = [];
      this.#cellDataStart = HEADER_SIZE;
      return;
    }

    let offset = HEADER_SIZE;
    this.#rowIndex = new Array(this.#rowCount);
    for (let i = 0; i < this.#rowCount; i++) {
      this.#rowIndex[i] = {
        rowNum: this.#view.getUint32(offset, true),
        cellOffset: this.#view.getUint32(offset + 4, true),
      };
      offset += 8;
    }

    const strCount = this.#view.getUint32(offset, true);
    const blobSize = this.#view.getUint32(offset + 4, true);
    if (strCount === 0) {
      this.#strings = [];
      this.#cellDataStart = offset + 8;
    } else {
      const offsetsStart = offset + 8;
      const blobStart = offsetsStart + strCount * 4;
      this.#strings = new Array(strCount);
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

  get rowCount() {
    return this.#rowCount;
  }

  get colCount() {
    return this.#colCount;
  }

  #decodeCellValue(type, payloadOffset) {
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

  #findRowIndex(rowNum) {
    for (let i = 0; i < this.#rowIndex.length; i++) {
      if (this.#rowIndex[i].rowNum === rowNum) {
        return i;
      }
    }
    return -1;
  }

  getCell(row, col) {
    if (this.#rowCount === 0) return null;
    const ri = this.#findRowIndex(row);
    if (ri === -1) return null;
    const { cellOffset } = this.#rowIndex[ri];
    if (cellOffset === EMPTY_ROW_OFFSET) return null;
    const colIdx = col - this.#minCol;
    if (colIdx < 0 || colIdx >= this.#colCount) return null;

    if (this.#isSparse) {
      const absOff = this.#cellDataStart + cellOffset;
      const cellCount = this.#view.getUint16(absOff, true);
      let pos = absOff + 2;
      for (let i = 0; i < cellCount; i++) {
        const c = this.#view.getUint16(pos, true);
        if (c === colIdx) {
          const type = this.#view.getUint8(pos + 2);
          if (type === TYPE_EMPTY) return null;
          return this.#decodeCellValue(type, pos + 3);
        }
        pos += SPARSE_ENTRY_SIZE;
      }
      return null;
    }

    const cellPos = this.#cellDataStart + cellOffset + colIdx * CELL_STRIDE;
    const type = this.#view.getUint8(cellPos);
    if (type === TYPE_EMPTY) return null;
    return this.#decodeCellValue(type, cellPos + 1);
  }

  getCellType(row, col) {
    if (this.#rowCount === 0) return 'empty';
    const ri = this.#findRowIndex(row);
    if (ri === -1) return 'empty';
    const { cellOffset } = this.#rowIndex[ri];
    if (cellOffset === EMPTY_ROW_OFFSET) return 'empty';
    const colIdx = col - this.#minCol;
    if (colIdx < 0 || colIdx >= this.#colCount) return 'empty';

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

  getRow(rowNum) {
    if (this.#rowCount === 0) return [];
    const ri = this.#findRowIndex(rowNum);
    if (ri === -1) return [];
    const { cellOffset } = this.#rowIndex[ri];
    if (cellOffset === EMPTY_ROW_OFFSET) return [];

    const result = [];

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
          result.push(this.#decodeCellValue(type, pos + 3));
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
        result.push(this.#decodeCellValue(type, cellPos + 1));
      }
    }
    return result;
  }

  toArray() {
    const result = [];
    for (const { rowNum, cellOffset } of this.#rowIndex) {
      if (cellOffset === EMPTY_ROW_OFFSET) {
        result.push([]);
        continue;
      }
      result.push(this.getRow(rowNum));
    }
    return result;
  }

  *rows() {
    for (const { rowNum, cellOffset } of this.#rowIndex) {
      if (cellOffset === EMPTY_ROW_OFFSET) {
        yield { row: rowNum, values: [] };
        continue;
      }
      yield { row: rowNum, values: this.getRow(rowNum) };
    }
  }

  columnName(colIndex) {
    return columnNumberToName(this.#minCol + colIndex);
  }
}

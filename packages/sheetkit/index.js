import { JsStreamWriter, Workbook as NativeWorkbook } from './binding.js';
import { decodeRowsBuffer } from './buffer-codec.js';

const NATIVE_KEY = Symbol('native');

class Workbook {
  #native;

  constructor(nativeInstance) {
    if (nativeInstance?.[NATIVE_KEY]) {
      this.#native = nativeInstance[NATIVE_KEY];
    } else {
      this.#native = new NativeWorkbook();
    }
  }

  static #wrap(native) {
    return new Workbook({ [NATIVE_KEY]: native });
  }

  static openSync(path) {
    return Workbook.#wrap(NativeWorkbook.openSync(path));
  }

  static async open(path) {
    return Workbook.#wrap(await NativeWorkbook.open(path));
  }

  static openBufferSync(data) {
    return Workbook.#wrap(NativeWorkbook.openBufferSync(data));
  }

  static async openBuffer(data) {
    return Workbook.#wrap(await NativeWorkbook.openBuffer(data));
  }

  static openWithPasswordSync(path, password) {
    return Workbook.#wrap(NativeWorkbook.openWithPasswordSync(path, password));
  }

  static async openWithPassword(path, password) {
    return Workbook.#wrap(await NativeWorkbook.openWithPassword(path, password));
  }

  get sheetNames() {
    return this.#native.sheetNames;
  }

  saveSync(path) {
    return this.#native.saveSync(path);
  }

  async save(path) {
    return this.#native.save(path);
  }

  writeBufferSync() {
    return this.#native.writeBufferSync();
  }

  async writeBuffer() {
    return this.#native.writeBuffer();
  }

  saveWithPasswordSync(path, password) {
    return this.#native.saveWithPasswordSync(path, password);
  }

  async saveWithPassword(path, password) {
    return this.#native.saveWithPassword(path, password);
  }

  getCellValue(sheet, cell) {
    return this.#native.getCellValue(sheet, cell);
  }

  setCellValue(sheet, cell, value) {
    return this.#native.setCellValue(sheet, cell, value);
  }

  setCellValues(sheet, cells) {
    return this.#native.setCellValues(sheet, cells);
  }

  setRowValues(sheet, row, startCol, values) {
    return this.#native.setRowValues(sheet, row, startCol, values);
  }

  setSheetData(sheet, data, startCell) {
    return this.#native.setSheetData(sheet, data, startCell);
  }

  newSheet(name) {
    return this.#native.newSheet(name);
  }

  deleteSheet(name) {
    return this.#native.deleteSheet(name);
  }

  setSheetName(oldName, newName) {
    return this.#native.setSheetName(oldName, newName);
  }

  copySheet(source, target) {
    return this.#native.copySheet(source, target);
  }

  getSheetIndex(name) {
    return this.#native.getSheetIndex(name);
  }

  getActiveSheet() {
    return this.#native.getActiveSheet();
  }

  setActiveSheet(name) {
    return this.#native.setActiveSheet(name);
  }

  insertRows(sheet, startRow, count) {
    return this.#native.insertRows(sheet, startRow, count);
  }

  removeRow(sheet, row) {
    return this.#native.removeRow(sheet, row);
  }

  duplicateRow(sheet, row) {
    return this.#native.duplicateRow(sheet, row);
  }

  setRowHeight(sheet, row, height) {
    return this.#native.setRowHeight(sheet, row, height);
  }

  getRowHeight(sheet, row) {
    return this.#native.getRowHeight(sheet, row);
  }

  setRowVisible(sheet, row, visible) {
    return this.#native.setRowVisible(sheet, row, visible);
  }

  getRowVisible(sheet, row) {
    return this.#native.getRowVisible(sheet, row);
  }

  setRowOutlineLevel(sheet, row, level) {
    return this.#native.setRowOutlineLevel(sheet, row, level);
  }

  getRowOutlineLevel(sheet, row) {
    return this.#native.getRowOutlineLevel(sheet, row);
  }

  setColWidth(sheet, col, width) {
    return this.#native.setColWidth(sheet, col, width);
  }

  getColWidth(sheet, col) {
    return this.#native.getColWidth(sheet, col);
  }

  setColVisible(sheet, col, visible) {
    return this.#native.setColVisible(sheet, col, visible);
  }

  getColVisible(sheet, col) {
    return this.#native.getColVisible(sheet, col);
  }

  setColOutlineLevel(sheet, col, level) {
    return this.#native.setColOutlineLevel(sheet, col, level);
  }

  getColOutlineLevel(sheet, col) {
    return this.#native.getColOutlineLevel(sheet, col);
  }

  insertCols(sheet, col, count) {
    return this.#native.insertCols(sheet, col, count);
  }

  removeCol(sheet, col) {
    return this.#native.removeCol(sheet, col);
  }

  addStyle(style) {
    return this.#native.addStyle(style);
  }

  getCellStyle(sheet, cell) {
    return this.#native.getCellStyle(sheet, cell);
  }

  setCellStyle(sheet, cell, styleId) {
    return this.#native.setCellStyle(sheet, cell, styleId);
  }

  setRowStyle(sheet, row, styleId) {
    return this.#native.setRowStyle(sheet, row, styleId);
  }

  getRowStyle(sheet, row) {
    return this.#native.getRowStyle(sheet, row);
  }

  setColStyle(sheet, col, styleId) {
    return this.#native.setColStyle(sheet, col, styleId);
  }

  getColStyle(sheet, col) {
    return this.#native.getColStyle(sheet, col);
  }

  addChart(sheet, fromCell, toCell, config) {
    return this.#native.addChart(sheet, fromCell, toCell, config);
  }

  addImage(sheet, config) {
    return this.#native.addImage(sheet, config);
  }

  mergeCells(sheet, topLeft, bottomRight) {
    return this.#native.mergeCells(sheet, topLeft, bottomRight);
  }

  unmergeCell(sheet, reference) {
    return this.#native.unmergeCell(sheet, reference);
  }

  getMergeCells(sheet) {
    return this.#native.getMergeCells(sheet);
  }

  addDataValidation(sheet, config) {
    return this.#native.addDataValidation(sheet, config);
  }

  getDataValidations(sheet) {
    return this.#native.getDataValidations(sheet);
  }

  removeDataValidation(sheet, sqref) {
    return this.#native.removeDataValidation(sheet, sqref);
  }

  setConditionalFormat(sheet, sqref, rules) {
    return this.#native.setConditionalFormat(sheet, sqref, rules);
  }

  getConditionalFormats(sheet) {
    return this.#native.getConditionalFormats(sheet);
  }

  deleteConditionalFormat(sheet, sqref) {
    return this.#native.deleteConditionalFormat(sheet, sqref);
  }

  addComment(sheet, config) {
    return this.#native.addComment(sheet, config);
  }

  getComments(sheet) {
    return this.#native.getComments(sheet);
  }

  removeComment(sheet, cell) {
    return this.#native.removeComment(sheet, cell);
  }

  setAutoFilter(sheet, range) {
    return this.#native.setAutoFilter(sheet, range);
  }

  removeAutoFilter(sheet) {
    return this.#native.removeAutoFilter(sheet);
  }

  newStreamWriter(sheetName) {
    return this.#native.newStreamWriter(sheetName);
  }

  applyStreamWriter(writer) {
    return this.#native.applyStreamWriter(writer);
  }

  setDocProps(props) {
    return this.#native.setDocProps(props);
  }

  getDocProps() {
    return this.#native.getDocProps();
  }

  setAppProps(props) {
    return this.#native.setAppProps(props);
  }

  getAppProps() {
    return this.#native.getAppProps();
  }

  setCustomProperty(name, value) {
    return this.#native.setCustomProperty(name, value);
  }

  getCustomProperty(name) {
    return this.#native.getCustomProperty(name);
  }

  deleteCustomProperty(name) {
    return this.#native.deleteCustomProperty(name);
  }

  protectWorkbook(config) {
    return this.#native.protectWorkbook(config);
  }

  unprotectWorkbook() {
    return this.#native.unprotectWorkbook();
  }

  isWorkbookProtected() {
    return this.#native.isWorkbookProtected();
  }

  setPanes(sheet, cell) {
    return this.#native.setPanes(sheet, cell);
  }

  unsetPanes(sheet) {
    return this.#native.unsetPanes(sheet);
  }

  getPanes(sheet) {
    return this.#native.getPanes(sheet);
  }

  setPageMargins(sheet, margins) {
    return this.#native.setPageMargins(sheet, margins);
  }

  getPageMargins(sheet) {
    return this.#native.getPageMargins(sheet);
  }

  setPageSetup(sheet, setup) {
    return this.#native.setPageSetup(sheet, setup);
  }

  getPageSetup(sheet) {
    return this.#native.getPageSetup(sheet);
  }

  setHeaderFooter(sheet, header, footer) {
    return this.#native.setHeaderFooter(sheet, header, footer);
  }

  getHeaderFooter(sheet) {
    return this.#native.getHeaderFooter(sheet);
  }

  setPrintOptions(sheet, opts) {
    return this.#native.setPrintOptions(sheet, opts);
  }

  getPrintOptions(sheet) {
    return this.#native.getPrintOptions(sheet);
  }

  insertPageBreak(sheet, row) {
    return this.#native.insertPageBreak(sheet, row);
  }

  removePageBreak(sheet, row) {
    return this.#native.removePageBreak(sheet, row);
  }

  getPageBreaks(sheet) {
    return this.#native.getPageBreaks(sheet);
  }

  setCellHyperlink(sheet, cell, opts) {
    return this.#native.setCellHyperlink(sheet, cell, opts);
  }

  getCellHyperlink(sheet, cell) {
    return this.#native.getCellHyperlink(sheet, cell);
  }

  deleteCellHyperlink(sheet, cell) {
    return this.#native.deleteCellHyperlink(sheet, cell);
  }

  getRows(sheet) {
    const buf = this.#native.getRowsBuffer(sheet);
    return decodeRowsBuffer(buf);
  }

  getRowsBuffer(sheet) {
    return this.#native.getRowsBuffer(sheet);
  }

  setSheetDataBuffer(sheet, buf, startCell) {
    return this.#native.setSheetDataBuffer(sheet, buf, startCell);
  }

  getCols(sheet) {
    return this.#native.getCols(sheet);
  }

  setCellFormula(sheet, cell, formula) {
    return this.#native.setCellFormula(sheet, cell, formula);
  }

  fillFormula(sheet, range, formula) {
    return this.#native.fillFormula(sheet, range, formula);
  }

  evaluateFormula(sheet, formula) {
    return this.#native.evaluateFormula(sheet, formula);
  }

  calculateAll() {
    return this.#native.calculateAll();
  }

  addPivotTable(config) {
    return this.#native.addPivotTable(config);
  }

  getPivotTables() {
    return this.#native.getPivotTables();
  }

  deletePivotTable(name) {
    return this.#native.deletePivotTable(name);
  }

  addSparkline(sheet, config) {
    return this.#native.addSparkline(sheet, config);
  }

  getSparklines(sheet) {
    return this.#native.getSparklines(sheet);
  }

  removeSparkline(sheet, location) {
    return this.#native.removeSparkline(sheet, location);
  }

  setCellRichText(sheet, cell, runs) {
    return this.#native.setCellRichText(sheet, cell, runs);
  }

  getCellRichText(sheet, cell) {
    return this.#native.getCellRichText(sheet, cell);
  }

  getThemeColor(index, tint) {
    return this.#native.getThemeColor(index, tint);
  }

  setDefinedName(config) {
    return this.#native.setDefinedName(config);
  }

  getDefinedName(name, scope) {
    return this.#native.getDefinedName(name, scope);
  }

  getDefinedNames() {
    return this.#native.getDefinedNames();
  }

  deleteDefinedName(name, scope) {
    return this.#native.deleteDefinedName(name, scope);
  }

  protectSheet(sheet, config) {
    return this.#native.protectSheet(sheet, config);
  }

  unprotectSheet(sheet) {
    return this.#native.unprotectSheet(sheet);
  }

  isSheetProtected(sheet) {
    return this.#native.isSheetProtected(sheet);
  }
}

export { JsStreamWriter, Workbook };

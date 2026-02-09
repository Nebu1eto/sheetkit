import type {
  DateValue,
  JsAppProperties,
  JsCellEntry,
  JsChartConfig,
  JsColData,
  JsCommentConfig,
  JsConditionalFormatEntry,
  JsConditionalFormatRule,
  JsDataValidationConfig,
  JsDefinedNameConfig,
  JsDefinedNameInfo,
  JsDocProperties,
  JsHeaderFooter,
  JsHyperlinkInfo,
  JsHyperlinkOptions,
  JsImageConfig,
  JsPageMargins,
  JsPageSetup,
  JsPivotTableConfig,
  JsPivotTableInfo,
  JsPrintOptions,
  JsRichTextRun,
  JsRowData,
  JsSheetProtectionConfig,
  JsSparklineConfig,
  JsStyle,
  JsWorkbookProtectionConfig,
} from './binding.js';
import { JsStreamWriter, Workbook as NativeWorkbook } from './binding.js';
import { decodeRowsBuffer } from './buffer-codec.js';
import type { CellTypeName, CellValue } from './sheet-data.js';
import { SheetData } from './sheet-data.js';

export type {
  DateValue,
  JsAlignmentStyle,
  JsAppProperties,
  JsBorderSideStyle,
  JsBorderStyle,
  JsCellEntry,
  JsChartConfig,
  JsChartSeries,
  JsColCell,
  JsColData,
  JsCommentConfig,
  JsConditionalFormatEntry,
  JsConditionalFormatRule,
  JsConditionalStyle,
  JsDataValidationConfig,
  JsDefinedNameConfig,
  JsDefinedNameInfo,
  JsDocProperties,
  JsFillStyle,
  JsFontStyle,
  JsHeaderFooter,
  JsHyperlinkInfo,
  JsHyperlinkOptions,
  JsImageConfig,
  JsPageMargins,
  JsPageSetup,
  JsPivotDataField,
  JsPivotField,
  JsPivotTableConfig,
  JsPivotTableInfo,
  JsPrintOptions,
  JsProtectionStyle,
  JsRichTextRun,
  JsRowCell,
  JsRowData,
  JsSheetProtectionConfig,
  JsSparklineConfig,
  JsStyle,
  JsView3DConfig,
  JsWorkbookProtectionConfig,
} from './binding.js';

type CellValueInput = string | number | boolean | DateValue | null;

/** Excel workbook for reading and writing .xlsx files. */
class Workbook {
  #native: NativeWorkbook;

  constructor() {
    this.#native = new NativeWorkbook();
  }

  static #wrap(native: NativeWorkbook): Workbook {
    const wb = new Workbook();
    wb.#native = native;
    return wb;
  }

  /** Open an existing .xlsx file from disk. */
  static openSync(path: string): Workbook {
    return Workbook.#wrap(NativeWorkbook.openSync(path));
  }

  /** Open an existing .xlsx file from disk asynchronously. */
  static async open(path: string): Promise<Workbook> {
    return Workbook.#wrap(await NativeWorkbook.open(path));
  }

  /** Open a workbook from an in-memory Buffer. */
  static openBufferSync(data: Buffer): Workbook {
    return Workbook.#wrap(NativeWorkbook.openBufferSync(data));
  }

  /** Open a workbook from an in-memory Buffer asynchronously. */
  static async openBuffer(data: Buffer): Promise<Workbook> {
    return Workbook.#wrap(await NativeWorkbook.openBuffer(data));
  }

  /** Open an encrypted .xlsx file using a password. */
  static openWithPasswordSync(path: string, password: string): Workbook {
    return Workbook.#wrap(NativeWorkbook.openWithPasswordSync(path, password));
  }

  /** Open an encrypted .xlsx file using a password asynchronously. */
  static async openWithPassword(path: string, password: string): Promise<Workbook> {
    return Workbook.#wrap(await NativeWorkbook.openWithPassword(path, password));
  }

  /** Get the names of all sheets in workbook order. */
  get sheetNames(): string[] {
    return this.#native.sheetNames;
  }

  /** Save the workbook to a .xlsx file. */
  saveSync(path: string): void {
    this.#native.saveSync(path);
  }

  /** Save the workbook to a .xlsx file asynchronously. */
  async save(path: string): Promise<void> {
    await this.#native.save(path);
  }

  /** Serialize the workbook to an in-memory Buffer. */
  writeBufferSync(): Buffer {
    return this.#native.writeBufferSync();
  }

  /** Serialize the workbook to an in-memory Buffer asynchronously. */
  async writeBuffer(): Promise<Buffer> {
    return this.#native.writeBuffer();
  }

  /** Save the workbook as an encrypted .xlsx file. */
  saveWithPasswordSync(path: string, password: string): void {
    this.#native.saveWithPasswordSync(path, password);
  }

  /** Save the workbook as an encrypted .xlsx file asynchronously. */
  async saveWithPassword(path: string, password: string): Promise<void> {
    await this.#native.saveWithPassword(path, password);
  }

  /** Get the value of a cell. Returns string, number, boolean, DateValue, or null. */
  getCellValue(sheet: string, cell: string): null | boolean | number | string | DateValue {
    return this.#native.getCellValue(sheet, cell);
  }

  /** Set the value of a cell. Pass string, number, boolean, DateValue, or null to clear. */
  setCellValue(sheet: string, cell: string, value: CellValueInput): void {
    this.#native.setCellValue(sheet, cell, value);
  }

  /** Set multiple cell values at once. */
  setCellValues(sheet: string, cells: JsCellEntry[]): void {
    this.#native.setCellValues(sheet, cells);
  }

  /** Set values in a single row starting from the given column. */
  setRowValues(sheet: string, row: number, startCol: string, values: CellValueInput[]): void {
    this.#native.setRowValues(sheet, row, startCol, values);
  }

  /** Set a block of cell values from a 2D array. */
  setSheetData(
    sheet: string,
    data: CellValueInput[][],
    startCell?: string | undefined | null,
  ): void {
    this.#native.setSheetData(sheet, data, startCell);
  }

  /** Create a new empty sheet. Returns the 0-based sheet index. */
  newSheet(name: string): number {
    return this.#native.newSheet(name);
  }

  /** Delete a sheet by name. */
  deleteSheet(name: string): void {
    this.#native.deleteSheet(name);
  }

  /** Rename a sheet. */
  setSheetName(oldName: string, newName: string): void {
    this.#native.setSheetName(oldName, newName);
  }

  /** Copy a sheet. Returns the new sheet's 0-based index. */
  copySheet(source: string, target: string): number {
    return this.#native.copySheet(source, target);
  }

  /** Get the 0-based index of a sheet, or null if not found. */
  getSheetIndex(name: string): number | null {
    return this.#native.getSheetIndex(name);
  }

  /** Get the name of the active sheet. */
  getActiveSheet(): string {
    return this.#native.getActiveSheet();
  }

  /** Set the active sheet by name. */
  setActiveSheet(name: string): void {
    this.#native.setActiveSheet(name);
  }

  /** Insert empty rows starting at the given 1-based row number. */
  insertRows(sheet: string, startRow: number, count: number): void {
    this.#native.insertRows(sheet, startRow, count);
  }

  /** Remove a row (1-based). */
  removeRow(sheet: string, row: number): void {
    this.#native.removeRow(sheet, row);
  }

  /** Duplicate a row (1-based). */
  duplicateRow(sheet: string, row: number): void {
    this.#native.duplicateRow(sheet, row);
  }

  /** Set the height of a row (1-based). */
  setRowHeight(sheet: string, row: number, height: number): void {
    this.#native.setRowHeight(sheet, row, height);
  }

  /** Get the height of a row, or null if not explicitly set. */
  getRowHeight(sheet: string, row: number): number | null {
    return this.#native.getRowHeight(sheet, row);
  }

  /** Set whether a row is visible. */
  setRowVisible(sheet: string, row: number, visible: boolean): void {
    this.#native.setRowVisible(sheet, row, visible);
  }

  /** Get whether a row is visible. */
  getRowVisible(sheet: string, row: number): boolean {
    return this.#native.getRowVisible(sheet, row);
  }

  /** Set the outline level of a row (0-7). */
  setRowOutlineLevel(sheet: string, row: number, level: number): void {
    this.#native.setRowOutlineLevel(sheet, row, level);
  }

  /** Get the outline level of a row. Returns 0 if not set. */
  getRowOutlineLevel(sheet: string, row: number): number {
    return this.#native.getRowOutlineLevel(sheet, row);
  }

  /** Set the width of a column (e.g., "A", "B", "AA"). */
  setColWidth(sheet: string, col: string, width: number): void {
    this.#native.setColWidth(sheet, col, width);
  }

  /** Get the width of a column, or null if not explicitly set. */
  getColWidth(sheet: string, col: string): number | null {
    return this.#native.getColWidth(sheet, col);
  }

  /** Set whether a column is visible. */
  setColVisible(sheet: string, col: string, visible: boolean): void {
    this.#native.setColVisible(sheet, col, visible);
  }

  /** Get whether a column is visible. */
  getColVisible(sheet: string, col: string): boolean {
    return this.#native.getColVisible(sheet, col);
  }

  /** Set the outline level of a column (0-7). */
  setColOutlineLevel(sheet: string, col: string, level: number): void {
    this.#native.setColOutlineLevel(sheet, col, level);
  }

  /** Get the outline level of a column. Returns 0 if not set. */
  getColOutlineLevel(sheet: string, col: string): number {
    return this.#native.getColOutlineLevel(sheet, col);
  }

  /** Insert empty columns starting at the given column letter. */
  insertCols(sheet: string, col: string, count: number): void {
    this.#native.insertCols(sheet, col, count);
  }

  /** Remove a column by letter. */
  removeCol(sheet: string, col: string): void {
    this.#native.removeCol(sheet, col);
  }

  /** Add a style definition. Returns the style ID for use with setCellStyle. */
  addStyle(style: JsStyle): number {
    return this.#native.addStyle(style);
  }

  /** Get the style ID applied to a cell, or null if default. */
  getCellStyle(sheet: string, cell: string): number | null {
    return this.#native.getCellStyle(sheet, cell);
  }

  /** Apply a style ID to a cell. */
  setCellStyle(sheet: string, cell: string, styleId: number): void {
    this.#native.setCellStyle(sheet, cell, styleId);
  }

  /** Apply a style ID to an entire row. */
  setRowStyle(sheet: string, row: number, styleId: number): void {
    this.#native.setRowStyle(sheet, row, styleId);
  }

  /** Get the style ID for a row. Returns 0 if not set. */
  getRowStyle(sheet: string, row: number): number {
    return this.#native.getRowStyle(sheet, row);
  }

  /** Apply a style ID to an entire column. */
  setColStyle(sheet: string, col: string, styleId: number): void {
    this.#native.setColStyle(sheet, col, styleId);
  }

  /** Get the style ID for a column. Returns 0 if not set. */
  getColStyle(sheet: string, col: string): number {
    return this.#native.getColStyle(sheet, col);
  }

  /** Add a chart to a sheet. */
  addChart(sheet: string, fromCell: string, toCell: string, config: JsChartConfig): void {
    this.#native.addChart(sheet, fromCell, toCell, config);
  }

  /** Add an image to a sheet. */
  addImage(sheet: string, config: JsImageConfig): void {
    this.#native.addImage(sheet, config);
  }

  /** Merge a range of cells on a sheet. */
  mergeCells(sheet: string, topLeft: string, bottomRight: string): void {
    this.#native.mergeCells(sheet, topLeft, bottomRight);
  }

  /** Remove a merged cell range from a sheet. */
  unmergeCell(sheet: string, reference: string): void {
    this.#native.unmergeCell(sheet, reference);
  }

  /** Get all merged cell ranges on a sheet. */
  getMergeCells(sheet: string): string[] {
    return this.#native.getMergeCells(sheet);
  }

  /** Add a data validation rule to a sheet. */
  addDataValidation(sheet: string, config: JsDataValidationConfig): void {
    this.#native.addDataValidation(sheet, config);
  }

  /** Get all data validations on a sheet. */
  getDataValidations(sheet: string): JsDataValidationConfig[] {
    return this.#native.getDataValidations(sheet);
  }

  /** Remove a data validation by sqref. */
  removeDataValidation(sheet: string, sqref: string): void {
    this.#native.removeDataValidation(sheet, sqref);
  }

  /** Set conditional formatting rules on a cell range. */
  setConditionalFormat(sheet: string, sqref: string, rules: JsConditionalFormatRule[]): void {
    this.#native.setConditionalFormat(sheet, sqref, rules);
  }

  /** Get all conditional formatting rules for a sheet. */
  getConditionalFormats(sheet: string): JsConditionalFormatEntry[] {
    return this.#native.getConditionalFormats(sheet);
  }

  /** Delete conditional formatting for a specific cell range. */
  deleteConditionalFormat(sheet: string, sqref: string): void {
    this.#native.deleteConditionalFormat(sheet, sqref);
  }

  /** Add a comment to a cell. */
  addComment(sheet: string, config: JsCommentConfig): void {
    this.#native.addComment(sheet, config);
  }

  /** Get all comments on a sheet. */
  getComments(sheet: string): JsCommentConfig[] {
    return this.#native.getComments(sheet);
  }

  /** Remove a comment from a cell. */
  removeComment(sheet: string, cell: string): void {
    this.#native.removeComment(sheet, cell);
  }

  /** Set an auto-filter on a sheet. */
  setAutoFilter(sheet: string, range: string): void {
    this.#native.setAutoFilter(sheet, range);
  }

  /** Remove the auto-filter from a sheet. */
  removeAutoFilter(sheet: string): void {
    this.#native.removeAutoFilter(sheet);
  }

  /** Create a new stream writer for a new sheet. */
  newStreamWriter(sheetName: string): JsStreamWriter {
    return this.#native.newStreamWriter(sheetName);
  }

  /** Apply a stream writer's output to the workbook. Returns the sheet index. */
  applyStreamWriter(writer: JsStreamWriter): number {
    return this.#native.applyStreamWriter(writer);
  }

  /** Set core document properties (title, creator, etc.). */
  setDocProps(props: JsDocProperties): void {
    this.#native.setDocProps(props);
  }

  /** Get core document properties. */
  getDocProps(): JsDocProperties {
    return this.#native.getDocProps();
  }

  /** Set application properties (company, app version, etc.). */
  setAppProps(props: JsAppProperties): void {
    this.#native.setAppProps(props);
  }

  /** Get application properties. */
  getAppProps(): JsAppProperties {
    return this.#native.getAppProps();
  }

  /** Set a custom property. Value can be string, number, or boolean. */
  setCustomProperty(name: string, value: string | number | boolean): void {
    this.#native.setCustomProperty(name, value);
  }

  /** Get a custom property value, or null if not found. */
  getCustomProperty(name: string): string | number | boolean | null {
    return this.#native.getCustomProperty(name);
  }

  /** Delete a custom property. Returns true if it existed. */
  deleteCustomProperty(name: string): boolean {
    return this.#native.deleteCustomProperty(name);
  }

  /** Protect the workbook structure/windows with optional password. */
  protectWorkbook(config: JsWorkbookProtectionConfig): void {
    this.#native.protectWorkbook(config);
  }

  /** Remove workbook protection. */
  unprotectWorkbook(): void {
    this.#native.unprotectWorkbook();
  }

  /** Check if the workbook is protected. */
  isWorkbookProtected(): boolean {
    return this.#native.isWorkbookProtected();
  }

  /** Set freeze panes on a sheet. */
  setPanes(sheet: string, cell: string): void {
    this.#native.setPanes(sheet, cell);
  }

  /** Remove any freeze or split panes from a sheet. */
  unsetPanes(sheet: string): void {
    this.#native.unsetPanes(sheet);
  }

  /** Get the current freeze pane cell reference for a sheet, or null if none. */
  getPanes(sheet: string): string | null {
    return this.#native.getPanes(sheet);
  }

  /** Set page margins on a sheet (values in inches). */
  setPageMargins(sheet: string, margins: JsPageMargins): void {
    this.#native.setPageMargins(sheet, margins);
  }

  /** Get page margins for a sheet. Returns defaults if not explicitly set. */
  getPageMargins(sheet: string): JsPageMargins {
    return this.#native.getPageMargins(sheet);
  }

  /** Set page setup options (paper size, orientation, scale, fit-to-page). */
  setPageSetup(sheet: string, setup: JsPageSetup): void {
    this.#native.setPageSetup(sheet, setup);
  }

  /** Get the page setup for a sheet. */
  getPageSetup(sheet: string): JsPageSetup {
    return this.#native.getPageSetup(sheet);
  }

  /** Set header and footer text for printing. */
  setHeaderFooter(
    sheet: string,
    header?: string | undefined | null,
    footer?: string | undefined | null,
  ): void {
    this.#native.setHeaderFooter(sheet, header, footer);
  }

  /** Get the header and footer text for a sheet. */
  getHeaderFooter(sheet: string): JsHeaderFooter {
    return this.#native.getHeaderFooter(sheet);
  }

  /** Set print options on a sheet. */
  setPrintOptions(sheet: string, opts: JsPrintOptions): void {
    this.#native.setPrintOptions(sheet, opts);
  }

  /** Get print options for a sheet. */
  getPrintOptions(sheet: string): JsPrintOptions {
    return this.#native.getPrintOptions(sheet);
  }

  /** Insert a horizontal page break before the given 1-based row. */
  insertPageBreak(sheet: string, row: number): void {
    this.#native.insertPageBreak(sheet, row);
  }

  /** Remove a horizontal page break at the given 1-based row. */
  removePageBreak(sheet: string, row: number): void {
    this.#native.removePageBreak(sheet, row);
  }

  /** Get all row page break positions (1-based row numbers). */
  getPageBreaks(sheet: string): number[] {
    return this.#native.getPageBreaks(sheet);
  }

  /** Set a hyperlink on a cell. */
  setCellHyperlink(sheet: string, cell: string, opts: JsHyperlinkOptions): void {
    this.#native.setCellHyperlink(sheet, cell, opts);
  }

  /** Get hyperlink information for a cell, or null if no hyperlink exists. */
  getCellHyperlink(sheet: string, cell: string): JsHyperlinkInfo | null {
    return this.#native.getCellHyperlink(sheet, cell);
  }

  /** Delete a hyperlink from a cell. */
  deleteCellHyperlink(sheet: string, cell: string): void {
    this.#native.deleteCellHyperlink(sheet, cell);
  }

  /** Get all rows with their data from a sheet using buffer-based transfer. */
  getRows(sheet: string): JsRowData[] {
    const buf = this.#native.getRowsBuffer(sheet);
    return decodeRowsBuffer(buf);
  }

  /** Serialize a sheet's cell data into a compact binary buffer. */
  getRowsBuffer(sheet: string): Buffer {
    return this.#native.getRowsBuffer(sheet);
  }

  /** Apply cell data from a binary buffer to a sheet. */
  setSheetDataBuffer(sheet: string, buf: Buffer, startCell?: string | undefined | null): void {
    this.#native.setSheetDataBuffer(sheet, buf, startCell);
  }

  /** Get all columns with their data from a sheet. */
  getCols(sheet: string): JsColData[] {
    return this.#native.getCols(sheet);
  }

  /** Set a formula on a cell. */
  setCellFormula(sheet: string, cell: string, formula: string): void {
    this.#native.setCellFormula(sheet, cell, formula);
  }

  /** Fill a range with a formula, adjusting row references. */
  fillFormula(sheet: string, range: string, formula: string): void {
    this.#native.fillFormula(sheet, range, formula);
  }

  /** Evaluate a formula string against the current workbook data. */
  evaluateFormula(sheet: string, formula: string): null | boolean | number | string | DateValue {
    return this.#native.evaluateFormula(sheet, formula);
  }

  /** Recalculate all formula cells in the workbook. */
  calculateAll(): void {
    this.#native.calculateAll();
  }

  /** Add a pivot table to the workbook. */
  addPivotTable(config: JsPivotTableConfig): void {
    this.#native.addPivotTable(config);
  }

  /** Get all pivot tables in the workbook. */
  getPivotTables(): JsPivotTableInfo[] {
    return this.#native.getPivotTables();
  }

  /** Delete a pivot table by name. */
  deletePivotTable(name: string): void {
    this.#native.deletePivotTable(name);
  }

  /** Add a sparkline to a worksheet. */
  addSparkline(sheet: string, config: JsSparklineConfig): void {
    this.#native.addSparkline(sheet, config);
  }

  /** Get all sparklines for a worksheet. */
  getSparklines(sheet: string): JsSparklineConfig[] {
    return this.#native.getSparklines(sheet);
  }

  /** Remove a sparkline by its location cell reference. */
  removeSparkline(sheet: string, location: string): void {
    this.#native.removeSparkline(sheet, location);
  }

  /** Set a cell to a rich text value with multiple formatted runs. */
  setCellRichText(sheet: string, cell: string, runs: JsRichTextRun[]): void {
    this.#native.setCellRichText(sheet, cell, runs);
  }

  /** Get rich text runs for a cell, or null if not rich text. */
  getCellRichText(sheet: string, cell: string): JsRichTextRun[] | null {
    return this.#native.getCellRichText(sheet, cell);
  }

  /** Resolve a theme color by index (0-11) with optional tint. */
  getThemeColor(index: number, tint?: number | undefined | null): string | null {
    return this.#native.getThemeColor(index, tint);
  }

  /** Add or update a defined name. */
  setDefinedName(config: JsDefinedNameConfig): void {
    this.#native.setDefinedName(config);
  }

  /** Get a defined name by name and optional scope. */
  getDefinedName(name: string, scope?: string | undefined | null): JsDefinedNameInfo | null {
    return this.#native.getDefinedName(name, scope);
  }

  /** Get all defined names in the workbook. */
  getDefinedNames(): JsDefinedNameInfo[] {
    return this.#native.getDefinedNames();
  }

  /** Delete a defined name by name and optional scope. */
  deleteDefinedName(name: string, scope?: string | undefined | null): void {
    this.#native.deleteDefinedName(name, scope);
  }

  /** Protect a sheet with optional password and permission settings. */
  protectSheet(sheet: string, config?: JsSheetProtectionConfig | undefined | null): void {
    this.#native.protectSheet(sheet, config);
  }

  /** Remove sheet protection. */
  unprotectSheet(sheet: string): void {
    this.#native.unprotectSheet(sheet);
  }

  /** Check if a sheet is protected. */
  isSheetProtected(sheet: string): boolean {
    return this.#native.isSheetProtected(sheet);
  }
}

export { JsStreamWriter, SheetData, Workbook };
export type { CellTypeName, CellValue };

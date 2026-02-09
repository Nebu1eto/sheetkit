export {
  JsStreamWriter,
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
} from './binding.d.ts';

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
  JsStreamWriter,
  JsStyle,
  JsWorkbookProtectionConfig,
} from './binding.d.ts';

/** Excel workbook for reading and writing .xlsx files. */
export declare class Workbook {
  /** Create a new empty workbook with a single sheet named "Sheet1". */
  constructor();
  /** Open an existing .xlsx file from disk. */
  static openSync(path: string): Workbook;
  /** Open an existing .xlsx file from disk asynchronously. */
  static open(path: string): Promise<Workbook>;
  /** Save the workbook to a .xlsx file. */
  saveSync(path: string): void;
  /** Save the workbook to a .xlsx file asynchronously. */
  save(path: string): Promise<void>;
  /** Open a workbook from an in-memory Buffer. */
  static openBufferSync(data: Buffer): Workbook;
  /** Open a workbook from an in-memory Buffer asynchronously. */
  static openBuffer(data: Buffer): Promise<Workbook>;
  /** Open an encrypted .xlsx file using a password. */
  static openWithPasswordSync(path: string, password: string): Workbook;
  /** Open an encrypted .xlsx file using a password asynchronously. */
  static openWithPassword(path: string, password: string): Promise<Workbook>;
  /** Serialize the workbook to an in-memory Buffer. */
  writeBufferSync(): Buffer;
  /** Serialize the workbook to an in-memory Buffer asynchronously. */
  writeBuffer(): Promise<Buffer>;
  /** Save the workbook as an encrypted .xlsx file. */
  saveWithPasswordSync(path: string, password: string): void;
  /** Save the workbook as an encrypted .xlsx file asynchronously. */
  saveWithPassword(path: string, password: string): Promise<void>;
  /** Get the names of all sheets in workbook order. */
  get sheetNames(): Array<string>;
  /** Get the value of a cell. Returns string, number, boolean, DateValue, or null. */
  getCellValue(sheet: string, cell: string): null | boolean | number | string | DateValue;
  /** Set the value of a cell. Pass string, number, boolean, DateValue, or null to clear. */
  setCellValue(
    sheet: string,
    cell: string,
    value: string | number | boolean | DateValue | null,
  ): void;
  /**
   * Set multiple cell values at once. More efficient than calling
   * setCellValue repeatedly because it crosses the FFI boundary only once.
   */
  setCellValues(sheet: string, cells: Array<JsCellEntry>): void;
  /**
   * Set values in a single row starting from the given column.
   * Values are placed left-to-right starting at startCol (e.g., "A").
   */
  setRowValues(
    sheet: string,
    row: number,
    startCol: string,
    values: Array<string | number | boolean | DateValue | null>,
  ): void;
  /**
   * Set a block of cell values from a 2D array.
   * Each inner array is a row, each element is a cell value.
   * Optionally specify a start cell (default "A1").
   */
  setSheetData(
    sheet: string,
    data: Array<Array<string | number | boolean | DateValue | null>>,
    startCell?: string | undefined | null,
  ): void;
  /** Create a new empty sheet. Returns the 0-based sheet index. */
  newSheet(name: string): number;
  /** Delete a sheet by name. */
  deleteSheet(name: string): void;
  /** Rename a sheet. */
  setSheetName(oldName: string, newName: string): void;
  /** Copy a sheet. Returns the new sheet's 0-based index. */
  copySheet(source: string, target: string): number;
  /** Get the 0-based index of a sheet, or null if not found. */
  getSheetIndex(name: string): number | null;
  /** Get the name of the active sheet. */
  getActiveSheet(): string;
  /** Set the active sheet by name. */
  setActiveSheet(name: string): void;
  /** Insert empty rows starting at the given 1-based row number. */
  insertRows(sheet: string, startRow: number, count: number): void;
  /** Remove a row (1-based). */
  removeRow(sheet: string, row: number): void;
  /** Duplicate a row (1-based). */
  duplicateRow(sheet: string, row: number): void;
  /** Set the height of a row (1-based). */
  setRowHeight(sheet: string, row: number, height: number): void;
  /** Get the height of a row, or null if not explicitly set. */
  getRowHeight(sheet: string, row: number): number | null;
  /** Set whether a row is visible. */
  setRowVisible(sheet: string, row: number, visible: boolean): void;
  /** Get whether a row is visible. Returns true if visible (not hidden). */
  getRowVisible(sheet: string, row: number): boolean;
  /** Set the outline level of a row (0-7). */
  setRowOutlineLevel(sheet: string, row: number, level: number): void;
  /** Get the outline level of a row. Returns 0 if not set. */
  getRowOutlineLevel(sheet: string, row: number): number;
  /** Set the width of a column (e.g., "A", "B", "AA"). */
  setColWidth(sheet: string, col: string, width: number): void;
  /** Get the width of a column, or null if not explicitly set. */
  getColWidth(sheet: string, col: string): number | null;
  /** Set whether a column is visible. */
  setColVisible(sheet: string, col: string, visible: boolean): void;
  /** Get whether a column is visible. Returns true if visible (not hidden). */
  getColVisible(sheet: string, col: string): boolean;
  /** Set the outline level of a column (0-7). */
  setColOutlineLevel(sheet: string, col: string, level: number): void;
  /** Get the outline level of a column. Returns 0 if not set. */
  getColOutlineLevel(sheet: string, col: string): number;
  /** Insert empty columns starting at the given column letter. */
  insertCols(sheet: string, col: string, count: number): void;
  /** Remove a column by letter. */
  removeCol(sheet: string, col: string): void;
  /** Add a style definition. Returns the style ID for use with setCellStyle. */
  addStyle(style: JsStyle): number;
  /** Get the style ID applied to a cell, or null if default. */
  getCellStyle(sheet: string, cell: string): number | null;
  /** Apply a style ID to a cell. */
  setCellStyle(sheet: string, cell: string, styleId: number): void;
  /** Apply a style ID to an entire row. */
  setRowStyle(sheet: string, row: number, styleId: number): void;
  /** Get the style ID for a row. Returns 0 if not set. */
  getRowStyle(sheet: string, row: number): number;
  /** Apply a style ID to an entire column. */
  setColStyle(sheet: string, col: string, styleId: number): void;
  /** Get the style ID for a column. Returns 0 if not set. */
  getColStyle(sheet: string, col: string): number;
  /** Add a chart to a sheet. */
  addChart(sheet: string, fromCell: string, toCell: string, config: JsChartConfig): void;
  /** Add an image to a sheet. */
  addImage(sheet: string, config: JsImageConfig): void;
  /** Merge a range of cells on a sheet. */
  mergeCells(sheet: string, topLeft: string, bottomRight: string): void;
  /** Remove a merged cell range from a sheet. */
  unmergeCell(sheet: string, reference: string): void;
  /** Get all merged cell ranges on a sheet. */
  getMergeCells(sheet: string): Array<string>;
  /** Add a data validation rule to a sheet. */
  addDataValidation(sheet: string, config: JsDataValidationConfig): void;
  /** Get all data validations on a sheet. */
  getDataValidations(sheet: string): Array<JsDataValidationConfig>;
  /** Remove a data validation by sqref. */
  removeDataValidation(sheet: string, sqref: string): void;
  /** Set conditional formatting rules on a cell range. */
  setConditionalFormat(sheet: string, sqref: string, rules: Array<JsConditionalFormatRule>): void;
  /** Get all conditional formatting rules for a sheet. */
  getConditionalFormats(sheet: string): Array<JsConditionalFormatEntry>;
  /** Delete conditional formatting for a specific cell range. */
  deleteConditionalFormat(sheet: string, sqref: string): void;
  /** Add a comment to a cell. */
  addComment(sheet: string, config: JsCommentConfig): void;
  /** Get all comments on a sheet. */
  getComments(sheet: string): Array<JsCommentConfig>;
  /** Remove a comment from a cell. */
  removeComment(sheet: string, cell: string): void;
  /** Set an auto-filter on a sheet. */
  setAutoFilter(sheet: string, range: string): void;
  /** Remove the auto-filter from a sheet. */
  removeAutoFilter(sheet: string): void;
  /** Create a new stream writer for a new sheet. */
  newStreamWriter(sheetName: string): JsStreamWriter;
  /** Apply a stream writer's output to the workbook. Returns the sheet index. */
  applyStreamWriter(writer: JsStreamWriter): number;
  /** Set core document properties (title, creator, etc.). */
  setDocProps(props: JsDocProperties): void;
  /** Get core document properties. */
  getDocProps(): JsDocProperties;
  /** Set application properties (company, app version, etc.). */
  setAppProps(props: JsAppProperties): void;
  /** Get application properties. */
  getAppProps(): JsAppProperties;
  /** Set a custom property. Value can be string, number, or boolean. */
  setCustomProperty(name: string, value: string | number | boolean): void;
  /** Get a custom property value, or null if not found. */
  getCustomProperty(name: string): string | number | boolean | null;
  /** Delete a custom property. Returns true if it existed. */
  deleteCustomProperty(name: string): boolean;
  /** Protect the workbook structure/windows with optional password. */
  protectWorkbook(config: JsWorkbookProtectionConfig): void;
  /** Remove workbook protection. */
  unprotectWorkbook(): void;
  /** Check if the workbook is protected. */
  isWorkbookProtected(): boolean;
  /**
   * Set freeze panes on a sheet.
   * The cell reference indicates the top-left cell of the scrollable area.
   */
  setPanes(sheet: string, cell: string): void;
  /** Remove any freeze or split panes from a sheet. */
  unsetPanes(sheet: string): void;
  /** Get the current freeze pane cell reference for a sheet, or null if none. */
  getPanes(sheet: string): string | null;
  /** Set page margins on a sheet (values in inches). */
  setPageMargins(sheet: string, margins: JsPageMargins): void;
  /** Get page margins for a sheet. Returns defaults if not explicitly set. */
  getPageMargins(sheet: string): JsPageMargins;
  /** Set page setup options (paper size, orientation, scale, fit-to-page). */
  setPageSetup(sheet: string, setup: JsPageSetup): void;
  /** Get the page setup for a sheet. */
  getPageSetup(sheet: string): JsPageSetup;
  /** Set header and footer text for printing. */
  setHeaderFooter(
    sheet: string,
    header?: string | undefined | null,
    footer?: string | undefined | null,
  ): void;
  /** Get the header and footer text for a sheet. */
  getHeaderFooter(sheet: string): JsHeaderFooter;
  /** Set print options on a sheet. */
  setPrintOptions(sheet: string, opts: JsPrintOptions): void;
  /** Get print options for a sheet. */
  getPrintOptions(sheet: string): JsPrintOptions;
  /** Insert a horizontal page break before the given 1-based row. */
  insertPageBreak(sheet: string, row: number): void;
  /** Remove a horizontal page break at the given 1-based row. */
  removePageBreak(sheet: string, row: number): void;
  /** Get all row page break positions (1-based row numbers). */
  getPageBreaks(sheet: string): Array<number>;
  /** Set a hyperlink on a cell. */
  setCellHyperlink(sheet: string, cell: string, opts: JsHyperlinkOptions): void;
  /** Get hyperlink information for a cell, or null if no hyperlink exists. */
  getCellHyperlink(sheet: string, cell: string): JsHyperlinkInfo | null;
  /** Delete a hyperlink from a cell. */
  deleteCellHyperlink(sheet: string, cell: string): void;
  /**
   * Get all rows with their data from a sheet.
   * Only rows that contain at least one cell are included.
   * Uses buffer-based transfer for efficient memory usage.
   */
  getRows(sheet: string): Array<JsRowData>;
  /**
   * Serialize a sheet's cell data into a compact binary buffer.
   * Returns the raw bytes suitable for efficient JS-side decoding.
   */
  getRowsBuffer(sheet: string): Buffer;
  /**
   * Apply cell data from a binary buffer to a sheet.
   * The buffer must follow the raw transfer binary format.
   * Optionally specify a start cell (default "A1").
   */
  setSheetDataBuffer(sheet: string, buf: Buffer, startCell?: string | undefined | null): void;
  /**
   * Get all columns with their data from a sheet.
   * Only columns that have data are included.
   */
  getCols(sheet: string): Array<JsColData>;
  /** Set a formula on a cell. */
  setCellFormula(sheet: string, cell: string, formula: string): void;
  /**
   * Fill a single-column range with a formula, adjusting row references
   * for each row relative to the first cell.
   */
  fillFormula(sheet: string, range: string, formula: string): void;
  /** Evaluate a formula string against the current workbook data. */
  evaluateFormula(sheet: string, formula: string): null | boolean | number | string | DateValue;
  /** Recalculate all formula cells in the workbook. */
  calculateAll(): void;
  /** Add a pivot table to the workbook. */
  addPivotTable(config: JsPivotTableConfig): void;
  /** Get all pivot tables in the workbook. */
  getPivotTables(): Array<JsPivotTableInfo>;
  /** Delete a pivot table by name. */
  deletePivotTable(name: string): void;
  /** Add a sparkline to a worksheet. */
  addSparkline(sheet: string, config: JsSparklineConfig): void;
  /** Get all sparklines for a worksheet. */
  getSparklines(sheet: string): Array<JsSparklineConfig>;
  /** Remove a sparkline by its location cell reference. */
  removeSparkline(sheet: string, location: string): void;
  /** Set a cell to a rich text value with multiple formatted runs. */
  setCellRichText(sheet: string, cell: string, runs: Array<JsRichTextRun>): void;
  /** Get rich text runs for a cell, or null if not rich text. */
  getCellRichText(sheet: string, cell: string): Array<JsRichTextRun> | null;
  /**
   * Resolve a theme color by index (0-11) with optional tint.
   * Returns the ARGB hex string (e.g. "FF4472C4") or null if out of range.
   */
  getThemeColor(index: number, tint?: number | undefined | null): string | null;
  /**
   * Add or update a defined name. If a name with the same name and scope
   * already exists, its value and comment are updated.
   */
  setDefinedName(config: JsDefinedNameConfig): void;
  /**
   * Get a defined name by name and optional scope (sheet name).
   * Returns null if no matching defined name is found.
   */
  getDefinedName(name: string, scope?: string | undefined | null): JsDefinedNameInfo | null;
  /** Get all defined names in the workbook. */
  getDefinedNames(): Array<JsDefinedNameInfo>;
  /** Delete a defined name by name and optional scope (sheet name). */
  deleteDefinedName(name: string, scope?: string | undefined | null): void;
  /** Protect a sheet with optional password and permission settings. */
  protectSheet(sheet: string, config?: JsSheetProtectionConfig | undefined | null): void;
  /** Remove sheet protection. */
  unprotectSheet(sheet: string): void;
  /** Check if a sheet is protected. */
  isSheetProtected(sheet: string): boolean;
}

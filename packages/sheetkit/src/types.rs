use napi_derive::napi;

#[napi(object)]
pub struct JsFontStyle {
    pub name: Option<String>,
    pub size: Option<f64>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub strikethrough: Option<bool>,
    pub color: Option<String>,
}

#[napi(object)]
pub struct JsFillStyle {
    pub pattern: Option<String>,
    pub fg_color: Option<String>,
    pub bg_color: Option<String>,
}

#[napi(object)]
pub struct JsBorderSideStyle {
    pub style: Option<String>,
    pub color: Option<String>,
}

#[napi(object)]
pub struct JsBorderStyle {
    pub left: Option<JsBorderSideStyle>,
    pub right: Option<JsBorderSideStyle>,
    pub top: Option<JsBorderSideStyle>,
    pub bottom: Option<JsBorderSideStyle>,
    pub diagonal: Option<JsBorderSideStyle>,
}

#[napi(object)]
pub struct JsAlignmentStyle {
    pub horizontal: Option<String>,
    pub vertical: Option<String>,
    pub wrap_text: Option<bool>,
    pub text_rotation: Option<u32>,
    pub indent: Option<u32>,
    pub shrink_to_fit: Option<bool>,
}

#[napi(object)]
pub struct JsProtectionStyle {
    pub locked: Option<bool>,
    pub hidden: Option<bool>,
}

#[napi(object)]
pub struct DateValue {
    #[napi(js_name = "type")]
    pub kind: String,
    pub serial: f64,
    pub iso: Option<String>,
}

#[napi(object)]
pub struct JsStyle {
    pub font: Option<JsFontStyle>,
    pub fill: Option<JsFillStyle>,
    pub border: Option<JsBorderStyle>,
    pub alignment: Option<JsAlignmentStyle>,
    pub num_fmt_id: Option<u32>,
    pub custom_num_fmt: Option<String>,
    pub protection: Option<JsProtectionStyle>,
}

#[napi(object)]
pub struct JsChartSeries {
    pub name: String,
    pub categories: String,
    pub values: String,
    pub x_values: Option<String>,
    pub bubble_sizes: Option<String>,
}

#[napi(object)]
pub struct JsChartConfig {
    pub chart_type: String,
    pub title: Option<String>,
    pub series: Vec<JsChartSeries>,
    pub show_legend: Option<bool>,
    pub view_3d: Option<JsView3DConfig>,
}

#[napi(object)]
pub struct JsView3DConfig {
    pub rot_x: Option<i32>,
    pub rot_y: Option<i32>,
    pub depth_percent: Option<u32>,
    pub right_angle_axes: Option<bool>,
    pub perspective: Option<u32>,
}

#[napi(object)]
pub struct JsImageConfig {
    pub data: napi::bindgen_prelude::Buffer,
    pub format: String,
    pub from_cell: String,
    pub width_px: u32,
    pub height_px: u32,
}

#[napi(object)]
pub struct JsCommentConfig {
    pub cell: String,
    pub author: String,
    pub text: String,
}

#[napi(object)]
pub struct JsDataValidationConfig {
    pub sqref: String,
    pub validation_type: String,
    pub operator: Option<String>,
    pub formula1: Option<String>,
    pub formula2: Option<String>,
    pub allow_blank: Option<bool>,
    pub error_style: Option<String>,
    pub error_title: Option<String>,
    pub error_message: Option<String>,
    pub prompt_title: Option<String>,
    pub prompt_message: Option<String>,
    pub show_input_message: Option<bool>,
    pub show_error_message: Option<bool>,
}

#[napi(object)]
pub struct JsDocProperties {
    pub title: Option<String>,
    pub subject: Option<String>,
    pub creator: Option<String>,
    pub keywords: Option<String>,
    pub description: Option<String>,
    pub last_modified_by: Option<String>,
    pub revision: Option<String>,
    pub created: Option<String>,
    pub modified: Option<String>,
    pub category: Option<String>,
    pub content_status: Option<String>,
}

#[napi(object)]
pub struct JsAppProperties {
    pub application: Option<String>,
    pub doc_security: Option<u32>,
    pub company: Option<String>,
    pub app_version: Option<String>,
    pub manager: Option<String>,
    pub template: Option<String>,
}

#[napi(object)]
pub struct JsWorkbookProtectionConfig {
    pub password: Option<String>,
    pub lock_structure: Option<bool>,
    pub lock_windows: Option<bool>,
    pub lock_revision: Option<bool>,
}

/// Page margins configuration in inches.
#[napi(object)]
pub struct JsPageMargins {
    /// Left margin in inches (default 0.7).
    pub left: f64,
    /// Right margin in inches (default 0.7).
    pub right: f64,
    /// Top margin in inches (default 0.75).
    pub top: f64,
    /// Bottom margin in inches (default 0.75).
    pub bottom: f64,
    /// Header margin in inches (default 0.3).
    pub header: f64,
    /// Footer margin in inches (default 0.3).
    pub footer: f64,
}

/// Page setup configuration.
#[napi(object)]
pub struct JsPageSetup {
    /// Paper size: "letter", "tabloid", "legal", "a3", "a4", "a5", "b4", "b5".
    pub paper_size: Option<String>,
    /// Orientation: "portrait" or "landscape".
    pub orientation: Option<String>,
    /// Print scale percentage (10-400).
    pub scale: Option<u32>,
    /// Fit to this many pages wide.
    pub fit_to_width: Option<u32>,
    /// Fit to this many pages tall.
    pub fit_to_height: Option<u32>,
}

/// Print options configuration.
#[napi(object)]
pub struct JsPrintOptions {
    /// Print gridlines.
    pub grid_lines: Option<bool>,
    /// Print row/column headings.
    pub headings: Option<bool>,
    /// Center horizontally on page.
    pub horizontal_centered: Option<bool>,
    /// Center vertically on page.
    pub vertical_centered: Option<bool>,
}

/// Header and footer text.
#[napi(object)]
pub struct JsHeaderFooter {
    /// Header text (may use Excel formatting codes like &L, &C, &R).
    pub header: Option<String>,
    /// Footer text (may use Excel formatting codes like &L, &C, &R).
    pub footer: Option<String>,
}

#[napi(object)]
pub struct JsHyperlinkOptions {
    /// Type of hyperlink: "external", "internal", or "email".
    pub link_type: String,
    /// The target URL, sheet reference, or email address.
    pub target: String,
    /// Optional display text.
    pub display: Option<String>,
    /// Optional tooltip text.
    pub tooltip: Option<String>,
}

#[napi(object)]
pub struct JsHyperlinkInfo {
    /// Type of hyperlink: "external", "internal", or "email".
    pub link_type: String,
    /// The target URL, sheet reference, or email address.
    pub target: String,
    /// Optional display text.
    pub display: Option<String>,
    /// Optional tooltip text.
    pub tooltip: Option<String>,
}

/// Conditional formatting style (differential format).
#[napi(object)]
pub struct JsConditionalStyle {
    pub font: Option<JsFontStyle>,
    pub fill: Option<JsFillStyle>,
    pub border: Option<JsBorderStyle>,
    pub custom_num_fmt: Option<String>,
}

/// Conditional formatting rule configuration.
#[napi(object)]
pub struct JsConditionalFormatRule {
    /// Rule type: "cellIs", "expression", "colorScale", "dataBar",
    /// "duplicateValues", "uniqueValues", "top10", "bottom10",
    /// "aboveAverage", "containsBlanks", "notContainsBlanks",
    /// "containsErrors", "notContainsErrors", "containsText",
    /// "notContainsText", "beginsWith", "endsWith".
    pub rule_type: String,
    /// Comparison operator for cellIs rules.
    pub operator: Option<String>,
    /// First formula/value.
    pub formula: Option<String>,
    /// Second formula/value (for between/notBetween).
    pub formula2: Option<String>,
    /// Text for text-based rules.
    pub text: Option<String>,
    /// Rank for top10/bottom10 rules.
    pub rank: Option<u32>,
    /// Whether rank is a percentage.
    pub percent: Option<bool>,
    /// Whether rule is above average (for aboveAverage rules).
    pub above: Option<bool>,
    /// Whether equal values count as matching (for aboveAverage rules).
    pub equal_average: Option<bool>,
    /// Color scale minimum value type.
    pub min_type: Option<String>,
    /// Color scale minimum value.
    pub min_value: Option<String>,
    /// Color scale minimum color (ARGB hex).
    pub min_color: Option<String>,
    /// Color scale middle value type.
    pub mid_type: Option<String>,
    /// Color scale middle value.
    pub mid_value: Option<String>,
    /// Color scale middle color (ARGB hex).
    pub mid_color: Option<String>,
    /// Color scale maximum value type.
    pub max_type: Option<String>,
    /// Color scale maximum value.
    pub max_value: Option<String>,
    /// Color scale maximum color (ARGB hex).
    pub max_color: Option<String>,
    /// Data bar color (ARGB hex).
    pub bar_color: Option<String>,
    /// Whether to show the cell value alongside the data bar.
    pub show_value: Option<bool>,
    /// Differential style to apply.
    pub format: Option<JsConditionalStyle>,
    /// Rule priority (lower = higher precedence).
    pub priority: Option<u32>,
    /// If true, no lower-priority rules apply when this matches.
    pub stop_if_true: Option<bool>,
}

/// Result of getting conditional formats from a sheet.
#[napi(object)]
pub struct JsConditionalFormatEntry {
    /// Cell range (e.g., "A1:A100").
    pub sqref: String,
    /// Rules applied to this range.
    pub rules: Vec<JsConditionalFormatRule>,
}

/// A single cell entry with its column name and value.
#[napi(object)]
pub struct JsRowCell {
    /// Column name (e.g., "A", "B", "AA").
    pub column: String,
    /// Cell value type: "string", "number", "boolean", "date", "empty", "error", "formula".
    pub value_type: String,
    /// String representation of the cell value.
    pub value: Option<String>,
    /// Numeric value (only set when value_type is "number").
    pub number_value: Option<f64>,
    /// Boolean value (only set when value_type is "boolean").
    pub bool_value: Option<bool>,
}

/// A row with its 1-based row number and cell data.
#[napi(object)]
pub struct JsRowData {
    /// 1-based row number.
    pub row: u32,
    /// Cells with data in this row.
    pub cells: Vec<JsRowCell>,
}

/// A single cell entry with its row number and value.
#[napi(object)]
pub struct JsColCell {
    /// 1-based row number.
    pub row: u32,
    /// Cell value type: "string", "number", "boolean", "date", "empty", "error", "formula".
    pub value_type: String,
    /// String representation of the cell value.
    pub value: Option<String>,
    /// Numeric value (only set when value_type is "number").
    pub number_value: Option<f64>,
    /// Boolean value (only set when value_type is "boolean").
    pub bool_value: Option<bool>,
}

/// A column with its name and cell data.
#[napi(object)]
pub struct JsColData {
    /// Column name (e.g., "A", "B", "AA").
    pub column: String,
    /// Cells with data in this column.
    pub cells: Vec<JsColCell>,
}

#[napi(object)]
pub struct JsPivotField {
    pub name: String,
}

#[napi(object)]
pub struct JsPivotDataField {
    pub name: String,
    pub function: String,
    pub display_name: Option<String>,
}

#[napi(object)]
pub struct JsPivotTableConfig {
    pub name: String,
    pub source_sheet: String,
    pub source_range: String,
    pub target_sheet: String,
    pub target_cell: String,
    pub rows: Vec<JsPivotField>,
    pub columns: Vec<JsPivotField>,
    pub data: Vec<JsPivotDataField>,
}

#[napi(object)]
pub struct JsPivotTableInfo {
    pub name: String,
    pub source_sheet: String,
    pub source_range: String,
    pub target_sheet: String,
    pub location: String,
}

#[napi(object)]
pub struct JsSparklineConfig {
    pub data_range: String,
    pub location: String,
    pub sparkline_type: Option<String>,
    pub markers: Option<bool>,
    pub high_point: Option<bool>,
    pub low_point: Option<bool>,
    pub first_point: Option<bool>,
    pub last_point: Option<bool>,
    pub negative_points: Option<bool>,
    pub show_axis: Option<bool>,
    pub line_weight: Option<f64>,
    pub style: Option<u32>,
}

/// Configuration for setting a defined name.
#[napi(object)]
pub struct JsDefinedNameConfig {
    /// The name to define (e.g., "SalesData").
    pub name: String,
    /// The reference or formula (e.g., "Sheet1!$A$1:$D$10").
    pub value: String,
    /// Optional sheet name for sheet-scoped names. Omit for workbook scope.
    pub scope: Option<String>,
    /// Optional comment for the defined name.
    pub comment: Option<String>,
}

/// Information about a defined name returned by getDefinedName/getDefinedNames.
#[napi(object)]
pub struct JsDefinedNameInfo {
    /// The defined name.
    pub name: String,
    /// The reference or formula.
    pub value: String,
    /// Sheet name if sheet-scoped, or null if workbook-scoped.
    pub scope: Option<String>,
    /// Optional comment.
    pub comment: Option<String>,
}

/// Configuration for sheet protection.
#[napi(object)]
pub struct JsSheetProtectionConfig {
    /// Optional password (hashed with legacy Excel algorithm).
    pub password: Option<String>,
    /// Allow selecting locked cells.
    pub select_locked_cells: Option<bool>,
    /// Allow selecting unlocked cells.
    pub select_unlocked_cells: Option<bool>,
    /// Allow formatting cells.
    pub format_cells: Option<bool>,
    /// Allow formatting columns.
    pub format_columns: Option<bool>,
    /// Allow formatting rows.
    pub format_rows: Option<bool>,
    /// Allow inserting columns.
    pub insert_columns: Option<bool>,
    /// Allow inserting rows.
    pub insert_rows: Option<bool>,
    /// Allow inserting hyperlinks.
    pub insert_hyperlinks: Option<bool>,
    /// Allow deleting columns.
    pub delete_columns: Option<bool>,
    /// Allow deleting rows.
    pub delete_rows: Option<bool>,
    /// Allow sorting.
    pub sort: Option<bool>,
    /// Allow using auto-filter.
    pub auto_filter: Option<bool>,
    /// Allow using pivot tables.
    pub pivot_tables: Option<bool>,
}

/// A cell reference and value pair for batch operations.
#[napi(object)]
pub struct JsCellEntry {
    /// Cell reference (e.g., "A1", "B2").
    pub cell: String,
    /// Cell value: string, number, boolean, DateValue, or null.
    pub value:
        napi::bindgen_prelude::Either5<String, f64, bool, DateValue, napi::bindgen_prelude::Null>,
}

/// Sheet view options for controlling how a sheet is displayed.
#[napi(object)]
pub struct JsSheetViewOptions {
    /// Whether gridlines are shown. Defaults to true.
    pub show_gridlines: Option<bool>,
    /// Whether formulas are shown instead of their results. Defaults to false.
    pub show_formulas: Option<bool>,
    /// Whether row and column headers are shown. Defaults to true.
    pub show_row_col_headers: Option<bool>,
    /// Zoom scale as a percentage (10-400). Defaults to 100.
    pub zoom_scale: Option<u32>,
    /// View mode: "normal", "pageBreak", or "pageLayout".
    pub view_mode: Option<String>,
    /// Top-left cell visible in the view (e.g. "A1").
    pub top_left_cell: Option<String>,
}

/// A single formatted text segment within a rich text cell.
#[napi(object)]
pub struct JsRichTextRun {
    pub text: String,
    pub font: Option<String>,
    pub size: Option<f64>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub color: Option<String>,
}

/// A column definition within a table.
#[napi(object)]
pub struct JsTableColumn {
    /// The column header name.
    pub name: String,
    /// Optional totals row function (e.g., "sum", "count", "average").
    pub totals_row_function: Option<String>,
    /// Optional totals row label (used for the first column in totals row).
    pub totals_row_label: Option<String>,
}

/// Configuration for creating a table.
#[napi(object)]
pub struct JsTableConfig {
    /// The table name (must be unique within the workbook).
    pub name: String,
    /// The display name shown in the UI.
    pub display_name: String,
    /// The cell range (e.g. "A1:D10").
    pub range: String,
    /// Column definitions.
    pub columns: Vec<JsTableColumn>,
    /// Whether to show the header row. Defaults to true.
    pub show_header_row: Option<bool>,
    /// The table style name (e.g. "TableStyleMedium2").
    pub style_name: Option<String>,
    /// Whether to enable auto-filter on the table.
    pub auto_filter: Option<bool>,
    /// Whether to show first column formatting.
    pub show_first_column: Option<bool>,
    /// Whether to show last column formatting.
    pub show_last_column: Option<bool>,
    /// Whether to show row stripes.
    pub show_row_stripes: Option<bool>,
    /// Whether to show column stripes.
    pub show_column_stripes: Option<bool>,
}

/// Metadata about an existing table.
#[napi(object)]
pub struct JsTableInfo {
    /// The table name.
    pub name: String,
    /// The display name.
    pub display_name: String,
    /// The cell range (e.g. "A1:D10").
    pub range: String,
    /// Whether the table has a header row.
    pub show_header_row: bool,
    /// Whether auto-filter is enabled.
    pub auto_filter: bool,
    /// Column names.
    pub columns: Vec<String>,
    /// The style name, if any.
    pub style_name: Option<String>,
}

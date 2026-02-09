//! Worksheet XML schema structures.
//!
//! Represents `xl/worksheets/sheet*.xml` in the OOXML package.

use std::fmt;

use serde::de::Deserializer;
use serde::ser::Serializer;
use serde::{Deserialize, Serialize};

use crate::namespaces;

/// Worksheet root element.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename = "worksheet")]
pub struct WorksheetXml {
    #[serde(rename = "@xmlns")]
    pub xmlns: String,

    #[serde(rename = "@xmlns:r")]
    pub xmlns_r: String,

    #[serde(rename = "sheetPr", skip_serializing_if = "Option::is_none")]
    pub sheet_pr: Option<SheetPr>,

    #[serde(rename = "dimension", skip_serializing_if = "Option::is_none")]
    pub dimension: Option<Dimension>,

    #[serde(rename = "sheetViews", skip_serializing_if = "Option::is_none")]
    pub sheet_views: Option<SheetViews>,

    #[serde(rename = "sheetFormatPr", skip_serializing_if = "Option::is_none")]
    pub sheet_format_pr: Option<SheetFormatPr>,

    #[serde(rename = "cols", skip_serializing_if = "Option::is_none")]
    pub cols: Option<Cols>,

    #[serde(rename = "sheetData")]
    pub sheet_data: SheetData,

    #[serde(rename = "sheetProtection", skip_serializing_if = "Option::is_none")]
    pub sheet_protection: Option<SheetProtection>,

    #[serde(rename = "autoFilter", skip_serializing_if = "Option::is_none")]
    pub auto_filter: Option<AutoFilter>,

    #[serde(rename = "mergeCells", skip_serializing_if = "Option::is_none")]
    pub merge_cells: Option<MergeCells>,

    #[serde(
        rename = "conditionalFormatting",
        default,
        skip_serializing_if = "Vec::is_empty"
    )]
    pub conditional_formatting: Vec<ConditionalFormatting>,

    #[serde(rename = "dataValidations", skip_serializing_if = "Option::is_none")]
    pub data_validations: Option<DataValidations>,

    #[serde(rename = "hyperlinks", skip_serializing_if = "Option::is_none")]
    pub hyperlinks: Option<Hyperlinks>,

    #[serde(rename = "printOptions", skip_serializing_if = "Option::is_none")]
    pub print_options: Option<PrintOptions>,

    #[serde(rename = "pageMargins", skip_serializing_if = "Option::is_none")]
    pub page_margins: Option<PageMargins>,

    #[serde(rename = "pageSetup", skip_serializing_if = "Option::is_none")]
    pub page_setup: Option<PageSetup>,

    #[serde(rename = "headerFooter", skip_serializing_if = "Option::is_none")]
    pub header_footer: Option<HeaderFooter>,

    #[serde(rename = "rowBreaks", skip_serializing_if = "Option::is_none")]
    pub row_breaks: Option<RowBreaks>,

    #[serde(rename = "drawing", skip_serializing_if = "Option::is_none")]
    pub drawing: Option<DrawingRef>,

    #[serde(rename = "legacyDrawing", skip_serializing_if = "Option::is_none")]
    pub legacy_drawing: Option<LegacyDrawingRef>,

    #[serde(rename = "tableParts", skip_serializing_if = "Option::is_none")]
    pub table_parts: Option<TableParts>,
}

/// Sheet dimension reference.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Dimension {
    #[serde(rename = "@ref")]
    pub reference: String,
}

/// Sheet views container.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SheetViews {
    #[serde(rename = "sheetView")]
    pub sheet_views: Vec<SheetView>,
}

/// Individual sheet view.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SheetView {
    #[serde(rename = "@tabSelected", skip_serializing_if = "Option::is_none")]
    pub tab_selected: Option<bool>,

    #[serde(rename = "@zoomScale", skip_serializing_if = "Option::is_none")]
    pub zoom_scale: Option<u32>,

    #[serde(rename = "@workbookViewId")]
    pub workbook_view_id: u32,

    #[serde(rename = "pane", skip_serializing_if = "Option::is_none")]
    pub pane: Option<Pane>,

    #[serde(rename = "selection", default)]
    pub selection: Vec<Selection>,
}

/// Pane definition for split or frozen panes.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Pane {
    #[serde(rename = "@xSplit", skip_serializing_if = "Option::is_none")]
    pub x_split: Option<u32>,

    #[serde(rename = "@ySplit", skip_serializing_if = "Option::is_none")]
    pub y_split: Option<u32>,

    #[serde(rename = "@topLeftCell", skip_serializing_if = "Option::is_none")]
    pub top_left_cell: Option<String>,

    #[serde(rename = "@activePane", skip_serializing_if = "Option::is_none")]
    pub active_pane: Option<String>,

    #[serde(rename = "@state", skip_serializing_if = "Option::is_none")]
    pub state: Option<String>,
}

/// Cell selection.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Selection {
    #[serde(rename = "@pane", skip_serializing_if = "Option::is_none")]
    pub pane: Option<String>,

    #[serde(rename = "@activeCell", skip_serializing_if = "Option::is_none")]
    pub active_cell: Option<String>,

    #[serde(rename = "@sqref", skip_serializing_if = "Option::is_none")]
    pub sqref: Option<String>,
}

/// Sheet properties.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct SheetPr {
    #[serde(rename = "@codeName", skip_serializing_if = "Option::is_none")]
    pub code_name: Option<String>,

    #[serde(rename = "@filterMode", skip_serializing_if = "Option::is_none")]
    pub filter_mode: Option<bool>,

    #[serde(rename = "tabColor", skip_serializing_if = "Option::is_none")]
    pub tab_color: Option<TabColor>,

    #[serde(rename = "outlinePr", skip_serializing_if = "Option::is_none")]
    pub outline_pr: Option<OutlinePr>,
}

/// Tab color specification.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct TabColor {
    #[serde(rename = "@rgb", skip_serializing_if = "Option::is_none")]
    pub rgb: Option<String>,

    #[serde(rename = "@theme", skip_serializing_if = "Option::is_none")]
    pub theme: Option<u32>,

    #[serde(rename = "@indexed", skip_serializing_if = "Option::is_none")]
    pub indexed: Option<u32>,
}

/// Outline properties for grouping.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct OutlinePr {
    #[serde(rename = "@summaryBelow", skip_serializing_if = "Option::is_none")]
    pub summary_below: Option<bool>,

    #[serde(rename = "@summaryRight", skip_serializing_if = "Option::is_none")]
    pub summary_right: Option<bool>,
}

/// Sheet format properties.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SheetFormatPr {
    #[serde(rename = "@defaultRowHeight")]
    pub default_row_height: f64,

    #[serde(rename = "@defaultColWidth", skip_serializing_if = "Option::is_none")]
    pub default_col_width: Option<f64>,

    #[serde(rename = "@customHeight", skip_serializing_if = "Option::is_none")]
    pub custom_height: Option<bool>,

    #[serde(rename = "@outlineLevelRow", skip_serializing_if = "Option::is_none")]
    pub outline_level_row: Option<u8>,

    #[serde(rename = "@outlineLevelCol", skip_serializing_if = "Option::is_none")]
    pub outline_level_col: Option<u8>,
}

/// Sheet protection settings.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct SheetProtection {
    #[serde(rename = "@password", skip_serializing_if = "Option::is_none")]
    pub password: Option<String>,

    #[serde(rename = "@sheet", skip_serializing_if = "Option::is_none")]
    pub sheet: Option<bool>,

    #[serde(rename = "@objects", skip_serializing_if = "Option::is_none")]
    pub objects: Option<bool>,

    #[serde(rename = "@scenarios", skip_serializing_if = "Option::is_none")]
    pub scenarios: Option<bool>,

    #[serde(rename = "@selectLockedCells", skip_serializing_if = "Option::is_none")]
    pub select_locked_cells: Option<bool>,

    #[serde(
        rename = "@selectUnlockedCells",
        skip_serializing_if = "Option::is_none"
    )]
    pub select_unlocked_cells: Option<bool>,

    #[serde(rename = "@formatCells", skip_serializing_if = "Option::is_none")]
    pub format_cells: Option<bool>,

    #[serde(rename = "@formatColumns", skip_serializing_if = "Option::is_none")]
    pub format_columns: Option<bool>,

    #[serde(rename = "@formatRows", skip_serializing_if = "Option::is_none")]
    pub format_rows: Option<bool>,

    #[serde(rename = "@insertColumns", skip_serializing_if = "Option::is_none")]
    pub insert_columns: Option<bool>,

    #[serde(rename = "@insertRows", skip_serializing_if = "Option::is_none")]
    pub insert_rows: Option<bool>,

    #[serde(rename = "@insertHyperlinks", skip_serializing_if = "Option::is_none")]
    pub insert_hyperlinks: Option<bool>,

    #[serde(rename = "@deleteColumns", skip_serializing_if = "Option::is_none")]
    pub delete_columns: Option<bool>,

    #[serde(rename = "@deleteRows", skip_serializing_if = "Option::is_none")]
    pub delete_rows: Option<bool>,

    #[serde(rename = "@sort", skip_serializing_if = "Option::is_none")]
    pub sort: Option<bool>,

    #[serde(rename = "@autoFilter", skip_serializing_if = "Option::is_none")]
    pub auto_filter: Option<bool>,

    #[serde(rename = "@pivotTables", skip_serializing_if = "Option::is_none")]
    pub pivot_tables: Option<bool>,
}

/// Columns container.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Cols {
    #[serde(rename = "col")]
    pub cols: Vec<Col>,
}

/// Individual column definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Col {
    #[serde(rename = "@min")]
    pub min: u32,

    #[serde(rename = "@max")]
    pub max: u32,

    #[serde(rename = "@width", skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,

    #[serde(rename = "@style", skip_serializing_if = "Option::is_none")]
    pub style: Option<u32>,

    #[serde(rename = "@hidden", skip_serializing_if = "Option::is_none")]
    pub hidden: Option<bool>,

    #[serde(rename = "@customWidth", skip_serializing_if = "Option::is_none")]
    pub custom_width: Option<bool>,

    #[serde(rename = "@outlineLevel", skip_serializing_if = "Option::is_none")]
    pub outline_level: Option<u8>,
}

/// Sheet data container holding all rows.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SheetData {
    #[serde(rename = "row", default)]
    pub rows: Vec<Row>,
}

/// A single row of cells.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Row {
    /// 1-based row number.
    #[serde(rename = "@r")]
    pub r: u32,

    #[serde(rename = "@spans", skip_serializing_if = "Option::is_none")]
    pub spans: Option<String>,

    #[serde(rename = "@s", skip_serializing_if = "Option::is_none")]
    pub s: Option<u32>,

    #[serde(rename = "@customFormat", skip_serializing_if = "Option::is_none")]
    pub custom_format: Option<bool>,

    #[serde(rename = "@ht", skip_serializing_if = "Option::is_none")]
    pub ht: Option<f64>,

    #[serde(rename = "@hidden", skip_serializing_if = "Option::is_none")]
    pub hidden: Option<bool>,

    #[serde(rename = "@customHeight", skip_serializing_if = "Option::is_none")]
    pub custom_height: Option<bool>,

    #[serde(rename = "@outlineLevel", skip_serializing_if = "Option::is_none")]
    pub outline_level: Option<u8>,

    #[serde(rename = "c", default)]
    pub cells: Vec<Cell>,
}

/// Inline cell reference (e.g., "A1", "XFD1048576") stored without heap allocation.
/// Max Excel cell ref is "XFD1048576" = 10 chars, so [u8; 10] + u8 length suffices.
#[derive(Clone, Copy, Default, PartialEq, Eq)]
pub struct CompactCellRef {
    buf: [u8; 10],
    len: u8,
}

impl CompactCellRef {
    /// Create a new CompactCellRef from a string slice. Panics if the input exceeds 10 bytes.
    pub fn new(s: &str) -> Self {
        assert!(
            s.len() <= 10,
            "cell reference too long ({} bytes): {s}",
            s.len()
        );
        let mut buf = [0u8; 10];
        buf[..s.len()].copy_from_slice(s.as_bytes());
        Self {
            buf,
            len: s.len() as u8,
        }
    }

    /// Return the cell reference as a string slice.
    pub fn as_str(&self) -> &str {
        // Safety: we only ever store valid UTF-8 (ASCII cell refs).
        unsafe { std::str::from_utf8_unchecked(&self.buf[..self.len as usize]) }
    }
}

impl fmt::Display for CompactCellRef {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.as_str())
    }
}

impl fmt::Debug for CompactCellRef {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "CompactCellRef(\"{}\")", self.as_str())
    }
}

impl From<&str> for CompactCellRef {
    fn from(s: &str) -> Self {
        Self::new(s)
    }
}

impl From<String> for CompactCellRef {
    fn from(s: String) -> Self {
        Self::new(&s)
    }
}

impl AsRef<str> for CompactCellRef {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl PartialEq<&str> for CompactCellRef {
    fn eq(&self, other: &&str) -> bool {
        self.as_str() == *other
    }
}

impl PartialEq<str> for CompactCellRef {
    fn eq(&self, other: &str) -> bool {
        self.as_str() == other
    }
}

impl Serialize for CompactCellRef {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        serializer.serialize_str(self.as_str())
    }
}

impl<'de> Deserialize<'de> for CompactCellRef {
    fn deserialize<D: Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
        let s = String::deserialize(deserializer)?;
        Ok(CompactCellRef::new(&s))
    }
}

/// A single cell.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Cell {
    /// Cell reference (e.g., "A1").
    #[serde(rename = "@r")]
    pub r: CompactCellRef,

    /// Cached 1-based column number parsed from `r`. Populated at load time
    /// or at creation time to avoid repeated string parsing.
    #[serde(skip)]
    pub col: u32,

    /// Style index.
    #[serde(rename = "@s", skip_serializing_if = "Option::is_none")]
    pub s: Option<u32>,

    /// Cell data type.
    #[serde(rename = "@t", default, skip_serializing_if = "CellTypeTag::is_none")]
    pub t: CellTypeTag,

    /// Cell value.
    #[serde(rename = "v", skip_serializing_if = "Option::is_none")]
    pub v: Option<String>,

    /// Cell formula.
    #[serde(rename = "f", skip_serializing_if = "Option::is_none")]
    pub f: Option<CellFormula>,

    /// Inline string.
    #[serde(rename = "is", skip_serializing_if = "Option::is_none")]
    pub is: Option<InlineString>,
}

/// Cell data type tag, replacing `Option<String>` for zero-allocation matching.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CellTypeTag {
    /// No type attribute (typically means Number).
    #[default]
    None,
    /// Shared string ("s").
    SharedString,
    /// Explicit number ("n").
    Number,
    /// Boolean ("b").
    Boolean,
    /// Error ("e").
    Error,
    /// Inline string ("inlineStr").
    InlineString,
    /// Formula string result ("str").
    FormulaString,
    /// Date ("d").
    Date,
}

impl CellTypeTag {
    /// Returns `true` when the tag is `None`, used for `skip_serializing_if`.
    pub fn is_none(&self) -> bool {
        matches!(self, CellTypeTag::None)
    }

    /// Returns the XML string representation, or `None` for the default variant.
    pub fn as_str(&self) -> Option<&'static str> {
        match self {
            CellTypeTag::None => Option::None,
            CellTypeTag::SharedString => Some("s"),
            CellTypeTag::Number => Some("n"),
            CellTypeTag::Boolean => Some("b"),
            CellTypeTag::Error => Some("e"),
            CellTypeTag::InlineString => Some("inlineStr"),
            CellTypeTag::FormulaString => Some("str"),
            CellTypeTag::Date => Some("d"),
        }
    }
}

impl Serialize for CellTypeTag {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        match self.as_str() {
            Some(s) => serializer.serialize_str(s),
            Option::None => serializer.serialize_none(),
        }
    }
}

impl<'de> Deserialize<'de> for CellTypeTag {
    fn deserialize<D: Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
        let s = String::deserialize(deserializer)?;
        match s.as_str() {
            "s" => Ok(CellTypeTag::SharedString),
            "n" => Ok(CellTypeTag::Number),
            "b" => Ok(CellTypeTag::Boolean),
            "e" => Ok(CellTypeTag::Error),
            "inlineStr" => Ok(CellTypeTag::InlineString),
            "str" => Ok(CellTypeTag::FormulaString),
            "d" => Ok(CellTypeTag::Date),
            _ => Ok(CellTypeTag::None),
        }
    }
}

/// Cell formula.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CellFormula {
    #[serde(rename = "@t", skip_serializing_if = "Option::is_none")]
    pub t: Option<String>,

    #[serde(rename = "@ref", skip_serializing_if = "Option::is_none")]
    pub reference: Option<String>,

    #[serde(rename = "@si", skip_serializing_if = "Option::is_none")]
    pub si: Option<u32>,

    #[serde(rename = "$value", skip_serializing_if = "Option::is_none")]
    pub value: Option<String>,
}

/// Inline string within a cell.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct InlineString {
    #[serde(rename = "t", skip_serializing_if = "Option::is_none")]
    pub t: Option<String>,
}

/// Auto filter.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct AutoFilter {
    #[serde(rename = "@ref")]
    pub reference: String,
}

/// Data validations container.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct DataValidations {
    #[serde(rename = "@count", skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,

    #[serde(
        rename = "@disablePrompts",
        skip_serializing_if = "Option::is_none",
        default
    )]
    pub disable_prompts: Option<bool>,

    #[serde(rename = "@xWindow", skip_serializing_if = "Option::is_none", default)]
    pub x_window: Option<u32>,

    #[serde(rename = "@yWindow", skip_serializing_if = "Option::is_none", default)]
    pub y_window: Option<u32>,

    #[serde(rename = "dataValidation", default)]
    pub data_validations: Vec<DataValidation>,
}

/// Individual data validation rule.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct DataValidation {
    #[serde(rename = "@type", skip_serializing_if = "Option::is_none")]
    pub validation_type: Option<String>,

    #[serde(rename = "@operator", skip_serializing_if = "Option::is_none")]
    pub operator: Option<String>,

    #[serde(rename = "@allowBlank", skip_serializing_if = "Option::is_none")]
    pub allow_blank: Option<bool>,

    #[serde(
        rename = "@showDropDown",
        skip_serializing_if = "Option::is_none",
        default
    )]
    pub show_drop_down: Option<bool>,

    #[serde(rename = "@showInputMessage", skip_serializing_if = "Option::is_none")]
    pub show_input_message: Option<bool>,

    #[serde(rename = "@showErrorMessage", skip_serializing_if = "Option::is_none")]
    pub show_error_message: Option<bool>,

    #[serde(rename = "@errorStyle", skip_serializing_if = "Option::is_none")]
    pub error_style: Option<String>,

    #[serde(rename = "@imeMode", skip_serializing_if = "Option::is_none", default)]
    pub ime_mode: Option<String>,

    #[serde(rename = "@errorTitle", skip_serializing_if = "Option::is_none")]
    pub error_title: Option<String>,

    #[serde(rename = "@error", skip_serializing_if = "Option::is_none")]
    pub error: Option<String>,

    #[serde(rename = "@promptTitle", skip_serializing_if = "Option::is_none")]
    pub prompt_title: Option<String>,

    #[serde(rename = "@prompt", skip_serializing_if = "Option::is_none")]
    pub prompt: Option<String>,

    #[serde(rename = "@sqref")]
    pub sqref: String,

    #[serde(rename = "formula1", skip_serializing_if = "Option::is_none")]
    pub formula1: Option<String>,

    #[serde(rename = "formula2", skip_serializing_if = "Option::is_none")]
    pub formula2: Option<String>,
}

/// Merge cells container.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct MergeCells {
    #[serde(rename = "@count", skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,

    #[serde(rename = "mergeCell", default)]
    pub merge_cells: Vec<MergeCell>,
}

/// Individual merge cell reference.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct MergeCell {
    #[serde(rename = "@ref")]
    pub reference: String,
}

/// Hyperlinks container.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Hyperlinks {
    #[serde(rename = "hyperlink", default)]
    pub hyperlinks: Vec<Hyperlink>,
}

/// Individual hyperlink.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Hyperlink {
    #[serde(rename = "@ref")]
    pub reference: String,

    #[serde(
        rename = "@r:id",
        alias = "@id",
        skip_serializing_if = "Option::is_none"
    )]
    pub r_id: Option<String>,

    #[serde(rename = "@location", skip_serializing_if = "Option::is_none")]
    pub location: Option<String>,

    #[serde(rename = "@display", skip_serializing_if = "Option::is_none")]
    pub display: Option<String>,

    #[serde(rename = "@tooltip", skip_serializing_if = "Option::is_none")]
    pub tooltip: Option<String>,
}

/// Page margins.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct PageMargins {
    #[serde(rename = "@left")]
    pub left: f64,

    #[serde(rename = "@right")]
    pub right: f64,

    #[serde(rename = "@top")]
    pub top: f64,

    #[serde(rename = "@bottom")]
    pub bottom: f64,

    #[serde(rename = "@header")]
    pub header: f64,

    #[serde(rename = "@footer")]
    pub footer: f64,
}

/// Page setup.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct PageSetup {
    #[serde(rename = "@paperSize", skip_serializing_if = "Option::is_none")]
    pub paper_size: Option<u32>,

    #[serde(rename = "@orientation", skip_serializing_if = "Option::is_none")]
    pub orientation: Option<String>,

    #[serde(rename = "@scale", skip_serializing_if = "Option::is_none")]
    pub scale: Option<u32>,

    #[serde(rename = "@fitToWidth", skip_serializing_if = "Option::is_none")]
    pub fit_to_width: Option<u32>,

    #[serde(rename = "@fitToHeight", skip_serializing_if = "Option::is_none")]
    pub fit_to_height: Option<u32>,

    #[serde(rename = "@firstPageNumber", skip_serializing_if = "Option::is_none")]
    pub first_page_number: Option<u32>,

    #[serde(rename = "@horizontalDpi", skip_serializing_if = "Option::is_none")]
    pub horizontal_dpi: Option<u32>,

    #[serde(rename = "@verticalDpi", skip_serializing_if = "Option::is_none")]
    pub vertical_dpi: Option<u32>,

    #[serde(
        rename = "@r:id",
        alias = "@id",
        skip_serializing_if = "Option::is_none"
    )]
    pub r_id: Option<String>,
}

/// Header and footer for printing.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct HeaderFooter {
    #[serde(rename = "oddHeader", skip_serializing_if = "Option::is_none")]
    pub odd_header: Option<String>,

    #[serde(rename = "oddFooter", skip_serializing_if = "Option::is_none")]
    pub odd_footer: Option<String>,
}

/// Print options.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct PrintOptions {
    #[serde(rename = "@gridLines", skip_serializing_if = "Option::is_none")]
    pub grid_lines: Option<bool>,

    #[serde(rename = "@headings", skip_serializing_if = "Option::is_none")]
    pub headings: Option<bool>,

    #[serde(
        rename = "@horizontalCentered",
        skip_serializing_if = "Option::is_none"
    )]
    pub horizontal_centered: Option<bool>,

    #[serde(rename = "@verticalCentered", skip_serializing_if = "Option::is_none")]
    pub vertical_centered: Option<bool>,
}

/// Row page breaks container.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct RowBreaks {
    #[serde(rename = "@count", skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,

    #[serde(rename = "@manualBreakCount", skip_serializing_if = "Option::is_none")]
    pub manual_break_count: Option<u32>,

    #[serde(rename = "brk", default)]
    pub brk: Vec<Break>,
}

/// Individual page break entry.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Break {
    #[serde(rename = "@id")]
    pub id: u32,

    #[serde(rename = "@max", skip_serializing_if = "Option::is_none")]
    pub max: Option<u32>,

    #[serde(rename = "@man", skip_serializing_if = "Option::is_none")]
    pub man: Option<bool>,
}

/// Drawing reference.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct DrawingRef {
    #[serde(rename = "@r:id", alias = "@id")]
    pub r_id: String,
}

/// Legacy drawing reference (VML).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct LegacyDrawingRef {
    #[serde(rename = "@r:id", alias = "@id")]
    pub r_id: String,
}

/// Table parts container.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct TableParts {
    #[serde(rename = "@count", skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,

    #[serde(rename = "tablePart", default)]
    pub table_parts: Vec<TablePart>,
}

/// Individual table part reference.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct TablePart {
    #[serde(rename = "@r:id", alias = "@id")]
    pub r_id: String,
}

/// Conditional formatting container.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ConditionalFormatting {
    #[serde(rename = "@sqref")]
    pub sqref: String,

    #[serde(rename = "cfRule", default)]
    pub cf_rules: Vec<CfRule>,
}

/// Conditional formatting rule.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CfRule {
    #[serde(rename = "@type")]
    pub rule_type: String,

    #[serde(rename = "@dxfId", skip_serializing_if = "Option::is_none")]
    pub dxf_id: Option<u32>,

    #[serde(rename = "@priority")]
    pub priority: u32,

    #[serde(rename = "@operator", skip_serializing_if = "Option::is_none")]
    pub operator: Option<String>,

    #[serde(rename = "@text", skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,

    #[serde(rename = "@stopIfTrue", skip_serializing_if = "Option::is_none")]
    pub stop_if_true: Option<bool>,

    #[serde(rename = "@aboveAverage", skip_serializing_if = "Option::is_none")]
    pub above_average: Option<bool>,

    #[serde(rename = "@equalAverage", skip_serializing_if = "Option::is_none")]
    pub equal_average: Option<bool>,

    #[serde(rename = "@percent", skip_serializing_if = "Option::is_none")]
    pub percent: Option<bool>,

    #[serde(rename = "@rank", skip_serializing_if = "Option::is_none")]
    pub rank: Option<u32>,

    #[serde(rename = "@bottom", skip_serializing_if = "Option::is_none")]
    pub bottom: Option<bool>,

    #[serde(rename = "formula", default, skip_serializing_if = "Vec::is_empty")]
    pub formulas: Vec<String>,

    #[serde(rename = "colorScale", skip_serializing_if = "Option::is_none")]
    pub color_scale: Option<CfColorScale>,

    #[serde(rename = "dataBar", skip_serializing_if = "Option::is_none")]
    pub data_bar: Option<CfDataBar>,

    #[serde(rename = "iconSet", skip_serializing_if = "Option::is_none")]
    pub icon_set: Option<CfIconSet>,
}

/// Color scale definition for conditional formatting.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CfColorScale {
    #[serde(rename = "cfvo", default)]
    pub cfvos: Vec<CfVo>,

    #[serde(rename = "color", default)]
    pub colors: Vec<CfColor>,
}

/// Data bar definition for conditional formatting.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CfDataBar {
    #[serde(rename = "@showValue", skip_serializing_if = "Option::is_none")]
    pub show_value: Option<bool>,

    #[serde(rename = "cfvo", default)]
    pub cfvos: Vec<CfVo>,

    #[serde(rename = "color", skip_serializing_if = "Option::is_none")]
    pub color: Option<CfColor>,
}

/// Icon set definition for conditional formatting.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CfIconSet {
    #[serde(rename = "@iconSet", skip_serializing_if = "Option::is_none")]
    pub icon_set: Option<String>,

    #[serde(rename = "cfvo", default)]
    pub cfvos: Vec<CfVo>,
}

/// Conditional formatting value object.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CfVo {
    #[serde(rename = "@type")]
    pub value_type: String,

    #[serde(rename = "@val", skip_serializing_if = "Option::is_none")]
    pub val: Option<String>,
}

/// Color reference for conditional formatting.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CfColor {
    #[serde(rename = "@rgb", skip_serializing_if = "Option::is_none")]
    pub rgb: Option<String>,

    #[serde(rename = "@theme", skip_serializing_if = "Option::is_none")]
    pub theme: Option<u32>,

    #[serde(rename = "@tint", skip_serializing_if = "Option::is_none")]
    pub tint: Option<f64>,
}

impl Default for WorksheetXml {
    fn default() -> Self {
        Self {
            xmlns: namespaces::SPREADSHEET_ML.to_string(),
            xmlns_r: namespaces::RELATIONSHIPS.to_string(),
            sheet_pr: None,
            dimension: None,
            sheet_views: None,
            sheet_format_pr: None,
            cols: None,
            sheet_data: SheetData { rows: vec![] },
            sheet_protection: None,
            auto_filter: None,
            merge_cells: None,
            conditional_formatting: vec![],
            data_validations: None,
            hyperlinks: None,
            print_options: None,
            page_margins: None,
            page_setup: None,
            header_footer: None,
            row_breaks: None,
            drawing: None,
            legacy_drawing: None,
            table_parts: None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_worksheet_default() {
        let ws = WorksheetXml::default();
        assert_eq!(ws.xmlns, namespaces::SPREADSHEET_ML);
        assert_eq!(ws.xmlns_r, namespaces::RELATIONSHIPS);
        assert!(ws.sheet_data.rows.is_empty());
        assert!(ws.dimension.is_none());
        assert!(ws.sheet_views.is_none());
        assert!(ws.cols.is_none());
        assert!(ws.merge_cells.is_none());
        assert!(ws.page_margins.is_none());
        assert!(ws.sheet_pr.is_none());
        assert!(ws.sheet_protection.is_none());
    }

    #[test]
    fn test_worksheet_roundtrip() {
        let ws = WorksheetXml::default();
        let xml = quick_xml::se::to_string(&ws).unwrap();
        let parsed: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(ws.xmlns, parsed.xmlns);
        assert_eq!(ws.xmlns_r, parsed.xmlns_r);
        assert_eq!(ws.sheet_data.rows.len(), parsed.sheet_data.rows.len());
    }

    #[test]
    fn test_worksheet_with_data() {
        let ws = WorksheetXml {
            sheet_data: SheetData {
                rows: vec![Row {
                    r: 1,
                    spans: Some("1:3".to_string()),
                    s: None,
                    custom_format: None,
                    ht: None,
                    hidden: None,
                    custom_height: None,
                    outline_level: None,
                    cells: vec![
                        Cell {
                            r: CompactCellRef::new("A1"),
                            col: 1,
                            s: None,
                            t: CellTypeTag::SharedString,
                            v: Some("0".to_string()),
                            f: None,
                            is: None,
                        },
                        Cell {
                            r: CompactCellRef::new("B1"),
                            col: 2,
                            s: None,
                            t: CellTypeTag::None,
                            v: Some("42".to_string()),
                            f: None,
                            is: None,
                        },
                    ],
                }],
            },
            ..WorksheetXml::default()
        };

        let xml = quick_xml::se::to_string(&ws).unwrap();
        let parsed: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.sheet_data.rows.len(), 1);
        assert_eq!(parsed.sheet_data.rows[0].r, 1);
        assert_eq!(parsed.sheet_data.rows[0].cells.len(), 2);
        assert_eq!(parsed.sheet_data.rows[0].cells[0].r, "A1");
        assert_eq!(
            parsed.sheet_data.rows[0].cells[0].t,
            CellTypeTag::SharedString
        );
        assert_eq!(parsed.sheet_data.rows[0].cells[0].v, Some("0".to_string()));
        assert_eq!(parsed.sheet_data.rows[0].cells[1].r, "B1");
        assert_eq!(parsed.sheet_data.rows[0].cells[1].v, Some("42".to_string()));
    }

    #[test]
    fn test_cell_with_formula() {
        let cell = Cell {
            r: CompactCellRef::new("C1"),
            col: 3,
            s: None,
            t: CellTypeTag::None,
            v: Some("84".to_string()),
            f: Some(CellFormula {
                t: None,
                reference: None,
                si: None,
                value: Some("A1+B1".to_string()),
            }),
            is: None,
        };
        let xml = quick_xml::se::to_string(&cell).unwrap();
        assert!(xml.contains("A1+B1"));
        let parsed: Cell = quick_xml::de::from_str(&xml).unwrap();
        assert!(parsed.f.is_some());
        assert_eq!(parsed.f.unwrap().value, Some("A1+B1".to_string()));
    }

    #[test]
    fn test_cell_with_inline_string() {
        let cell = Cell {
            r: CompactCellRef::new("A1"),
            col: 1,
            s: None,
            t: CellTypeTag::InlineString,
            v: None,
            f: None,
            is: Some(InlineString {
                t: Some("Hello World".to_string()),
            }),
        };
        let xml = quick_xml::se::to_string(&cell).unwrap();
        assert!(xml.contains("Hello World"));
        let parsed: Cell = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.t, CellTypeTag::InlineString);
        assert!(parsed.is.is_some());
        assert_eq!(parsed.is.unwrap().t, Some("Hello World".to_string()));
    }

    #[test]
    fn test_parse_real_excel_worksheet() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:B2"/>
  <sheetData>
    <row r="1" spans="1:2">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
    </row>
    <row r="2" spans="1:2">
      <c r="A2"><v>100</v></c>
      <c r="B2"><v>200</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let parsed: WorksheetXml = quick_xml::de::from_str(xml).unwrap();
        assert_eq!(parsed.dimension.as_ref().unwrap().reference, "A1:B2");
        assert_eq!(parsed.sheet_data.rows.len(), 2);
        assert_eq!(parsed.sheet_data.rows[0].cells.len(), 2);
        assert_eq!(parsed.sheet_data.rows[0].cells[0].r, "A1");
        assert_eq!(
            parsed.sheet_data.rows[0].cells[0].t,
            CellTypeTag::SharedString
        );
        assert_eq!(parsed.sheet_data.rows[0].cells[0].v, Some("0".to_string()));
        assert_eq!(parsed.sheet_data.rows[1].cells[0].r, "A2");
        assert_eq!(
            parsed.sheet_data.rows[1].cells[0].v,
            Some("100".to_string())
        );
    }

    #[test]
    fn test_worksheet_with_merge_cells() {
        let ws = WorksheetXml {
            merge_cells: Some(MergeCells {
                count: Some(1),
                merge_cells: vec![MergeCell {
                    reference: "A1:B2".to_string(),
                }],
            }),
            ..WorksheetXml::default()
        };
        let xml = quick_xml::se::to_string(&ws).unwrap();
        assert!(xml.contains("mergeCells"));
        assert!(xml.contains("A1:B2"));
        let parsed: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
        assert!(parsed.merge_cells.is_some());
        assert_eq!(parsed.merge_cells.as_ref().unwrap().merge_cells.len(), 1);
    }

    #[test]
    fn test_empty_sheet_data_serialization() {
        let sd = SheetData { rows: vec![] };
        let xml = quick_xml::se::to_string(&sd).unwrap();
        // Empty SheetData should still be serializable
        let parsed: SheetData = quick_xml::de::from_str(&xml).unwrap();
        assert!(parsed.rows.is_empty());
    }

    #[test]
    fn test_row_optional_fields_not_serialized() {
        let row = Row {
            r: 1,
            spans: None,
            s: None,
            custom_format: None,
            ht: None,
            hidden: None,
            custom_height: None,
            outline_level: None,
            cells: vec![],
        };
        let xml = quick_xml::se::to_string(&row).unwrap();
        assert!(!xml.contains("spans"));
        assert!(!xml.contains("ht"));
        assert!(!xml.contains("hidden"));
    }

    #[test]
    fn test_cell_type_tag_as_str() {
        assert_eq!(CellTypeTag::Boolean.as_str(), Some("b"));
        assert_eq!(CellTypeTag::Date.as_str(), Some("d"));
        assert_eq!(CellTypeTag::Error.as_str(), Some("e"));
        assert_eq!(CellTypeTag::InlineString.as_str(), Some("inlineStr"));
        assert_eq!(CellTypeTag::Number.as_str(), Some("n"));
        assert_eq!(CellTypeTag::SharedString.as_str(), Some("s"));
        assert_eq!(CellTypeTag::FormulaString.as_str(), Some("str"));
        assert_eq!(CellTypeTag::None.as_str(), Option::None);
    }

    #[test]
    fn test_cell_type_tag_default_is_none() {
        assert_eq!(CellTypeTag::default(), CellTypeTag::None);
        assert!(CellTypeTag::None.is_none());
        assert!(!CellTypeTag::SharedString.is_none());
    }

    #[test]
    fn test_cell_type_tag_serde_round_trip() {
        let variants = [
            (CellTypeTag::SharedString, "s"),
            (CellTypeTag::Number, "n"),
            (CellTypeTag::Boolean, "b"),
            (CellTypeTag::Error, "e"),
            (CellTypeTag::InlineString, "inlineStr"),
            (CellTypeTag::FormulaString, "str"),
            (CellTypeTag::Date, "d"),
        ];
        for (tag, expected_str) in &variants {
            let cell = Cell {
                r: CompactCellRef::new("A1"),
                col: 1,
                s: None,
                t: *tag,
                v: Some("0".to_string()),
                f: None,
                is: None,
            };
            let xml = quick_xml::se::to_string(&cell).unwrap();
            assert!(
                xml.contains(&format!("t=\"{expected_str}\"")),
                "expected t=\"{expected_str}\" in: {xml}"
            );
            let parsed: Cell = quick_xml::de::from_str(&xml).unwrap();
            assert_eq!(parsed.t, *tag);
        }

        let cell_none = Cell {
            r: CompactCellRef::new("A1"),
            col: 1,
            s: None,
            t: CellTypeTag::None,
            v: Some("42".to_string()),
            f: None,
            is: None,
        };
        let xml = quick_xml::se::to_string(&cell_none).unwrap();
        assert!(
            !xml.contains("t="),
            "None variant should not emit t attribute: {xml}"
        );
        let parsed: Cell = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.t, CellTypeTag::None);
    }

    #[test]
    fn test_worksheet_with_cols() {
        let ws = WorksheetXml {
            cols: Some(Cols {
                cols: vec![Col {
                    min: 1,
                    max: 1,
                    width: Some(15.0),
                    style: None,
                    hidden: None,
                    custom_width: Some(true),
                    outline_level: None,
                }],
            }),
            ..WorksheetXml::default()
        };
        let xml = quick_xml::se::to_string(&ws).unwrap();
        let parsed: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
        assert!(parsed.cols.is_some());
        let cols = parsed.cols.unwrap();
        assert_eq!(cols.cols.len(), 1);
        assert_eq!(cols.cols[0].min, 1);
        assert_eq!(cols.cols[0].width, Some(15.0));
        assert_eq!(cols.cols[0].custom_width, Some(true));
    }

    #[test]
    fn test_sheet_protection_roundtrip() {
        let prot = SheetProtection {
            password: Some("ABCD".to_string()),
            sheet: Some(true),
            objects: Some(true),
            scenarios: Some(true),
            format_cells: Some(false),
            ..SheetProtection::default()
        };
        let xml = quick_xml::se::to_string(&prot).unwrap();
        let parsed: SheetProtection = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.password, Some("ABCD".to_string()));
        assert_eq!(parsed.sheet, Some(true));
        assert_eq!(parsed.objects, Some(true));
        assert_eq!(parsed.scenarios, Some(true));
        assert_eq!(parsed.format_cells, Some(false));
        assert!(parsed.sort.is_none());
    }

    #[test]
    fn test_sheet_pr_roundtrip() {
        let pr = SheetPr {
            code_name: Some("Sheet1".to_string()),
            tab_color: Some(TabColor {
                rgb: Some("FF0000".to_string()),
                theme: None,
                indexed: None,
            }),
            ..SheetPr::default()
        };
        let xml = quick_xml::se::to_string(&pr).unwrap();
        let parsed: SheetPr = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.code_name, Some("Sheet1".to_string()));
        assert!(parsed.tab_color.is_some());
        assert_eq!(parsed.tab_color.unwrap().rgb, Some("FF0000".to_string()));
    }

    #[test]
    fn test_sheet_format_pr_extended_fields() {
        let fmt = SheetFormatPr {
            default_row_height: 15.0,
            default_col_width: Some(10.0),
            custom_height: Some(true),
            outline_level_row: Some(2),
            outline_level_col: Some(1),
        };
        let xml = quick_xml::se::to_string(&fmt).unwrap();
        let parsed: SheetFormatPr = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.default_row_height, 15.0);
        assert_eq!(parsed.default_col_width, Some(10.0));
        assert_eq!(parsed.custom_height, Some(true));
        assert_eq!(parsed.outline_level_row, Some(2));
        assert_eq!(parsed.outline_level_col, Some(1));
    }

    #[test]
    fn test_compact_cell_ref_basic() {
        let r = CompactCellRef::new("A1");
        assert_eq!(r.as_str(), "A1");
        assert_eq!(r.len, 2);
    }

    #[test]
    fn test_compact_cell_ref_max_length() {
        let r = CompactCellRef::new("XFD1048576");
        assert_eq!(r.as_str(), "XFD1048576");
        assert_eq!(r.len, 10);
    }

    #[test]
    fn test_compact_cell_ref_various_lengths() {
        for s in &["A1", "B5", "Z99", "AA100", "XFD1048576"] {
            let r = CompactCellRef::new(s);
            assert_eq!(r.as_str(), *s);
        }
    }

    #[test]
    fn test_compact_cell_ref_display() {
        let r = CompactCellRef::new("C3");
        assert_eq!(format!("{r}"), "C3");
    }

    #[test]
    fn test_compact_cell_ref_debug() {
        let r = CompactCellRef::new("C3");
        let dbg = format!("{r:?}");
        assert!(dbg.contains("CompactCellRef"));
        assert!(dbg.contains("C3"));
    }

    #[test]
    fn test_compact_cell_ref_default() {
        let r = CompactCellRef::default();
        assert_eq!(r.as_str(), "");
        assert_eq!(r.len, 0);
    }

    #[test]
    fn test_compact_cell_ref_from_str() {
        let r: CompactCellRef = "D4".into();
        assert_eq!(r.as_str(), "D4");
    }

    #[test]
    fn test_compact_cell_ref_from_string() {
        let r: CompactCellRef = String::from("E5").into();
        assert_eq!(r.as_str(), "E5");
    }

    #[test]
    fn test_compact_cell_ref_as_ref_str() {
        let r = CompactCellRef::new("F6");
        let s: &str = r.as_ref();
        assert_eq!(s, "F6");
    }

    #[test]
    fn test_compact_cell_ref_partial_eq_str() {
        let r = CompactCellRef::new("G7");
        assert_eq!(r, "G7");
        assert!(r == "G7");
        assert!(r != "H8");
    }

    #[test]
    fn test_compact_cell_ref_copy() {
        let r1 = CompactCellRef::new("A1");
        let r2 = r1;
        assert_eq!(r1.as_str(), "A1");
        assert_eq!(r2.as_str(), "A1");
    }

    #[test]
    fn test_compact_cell_ref_serde_roundtrip() {
        let cell = Cell {
            r: CompactCellRef::new("XFD1048576"),
            col: 16384,
            s: None,
            t: CellTypeTag::None,
            v: Some("42".to_string()),
            f: None,
            is: None,
        };
        let xml = quick_xml::se::to_string(&cell).unwrap();
        assert!(xml.contains("XFD1048576"));
        let parsed: Cell = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.r, "XFD1048576");
        assert_eq!(parsed.v, Some("42".to_string()));
    }

    #[test]
    #[should_panic(expected = "cell reference too long")]
    fn test_compact_cell_ref_panics_on_overflow() {
        CompactCellRef::new("ABCDEFGHIJK");
    }
}

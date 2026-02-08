//! Worksheet XML schema structures.
//!
//! Represents `xl/worksheets/sheet*.xml` in the OOXML package.

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

    #[serde(rename = "mergeCells", skip_serializing_if = "Option::is_none")]
    pub merge_cells: Option<MergeCells>,

    #[serde(rename = "hyperlinks", skip_serializing_if = "Option::is_none")]
    pub hyperlinks: Option<Hyperlinks>,

    #[serde(rename = "pageMargins", skip_serializing_if = "Option::is_none")]
    pub page_margins: Option<PageMargins>,

    #[serde(rename = "pageSetup", skip_serializing_if = "Option::is_none")]
    pub page_setup: Option<PageSetup>,

    #[serde(rename = "drawing", skip_serializing_if = "Option::is_none")]
    pub drawing: Option<DrawingRef>,

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

    #[serde(rename = "selection", default)]
    pub selection: Vec<Selection>,
}

/// Cell selection.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Selection {
    #[serde(rename = "@activeCell", skip_serializing_if = "Option::is_none")]
    pub active_cell: Option<String>,

    #[serde(rename = "@sqref", skip_serializing_if = "Option::is_none")]
    pub sqref: Option<String>,
}

/// Sheet format properties.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SheetFormatPr {
    #[serde(rename = "@defaultRowHeight")]
    pub default_row_height: f64,

    #[serde(rename = "@defaultColWidth", skip_serializing_if = "Option::is_none")]
    pub default_col_width: Option<f64>,
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

/// A single cell.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Cell {
    /// Cell reference (e.g., "A1").
    #[serde(rename = "@r")]
    pub r: String,

    /// Style index.
    #[serde(rename = "@s", skip_serializing_if = "Option::is_none")]
    pub s: Option<u32>,

    /// Cell type: "b", "d", "e", "inlineStr", "n", "s", "str".
    #[serde(rename = "@t", skip_serializing_if = "Option::is_none")]
    pub t: Option<String>,

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

/// Cell type constants.
pub mod cell_types {
    pub const BOOLEAN: &str = "b";
    pub const DATE: &str = "d";
    pub const ERROR: &str = "e";
    pub const INLINE_STRING: &str = "inlineStr";
    pub const NUMBER: &str = "n";
    pub const SHARED_STRING: &str = "s";
    pub const FORMULA_STRING: &str = "str";
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

    #[serde(
        rename = "@r:id",
        alias = "@id",
        skip_serializing_if = "Option::is_none"
    )]
    pub r_id: Option<String>,
}

/// Drawing reference.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct DrawingRef {
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

impl Default for WorksheetXml {
    fn default() -> Self {
        Self {
            xmlns: namespaces::SPREADSHEET_ML.to_string(),
            xmlns_r: namespaces::RELATIONSHIPS.to_string(),
            dimension: None,
            sheet_views: None,
            sheet_format_pr: None,
            cols: None,
            sheet_data: SheetData { rows: vec![] },
            merge_cells: None,
            hyperlinks: None,
            page_margins: None,
            page_setup: None,
            drawing: None,
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
                            r: "A1".to_string(),
                            s: None,
                            t: Some(cell_types::SHARED_STRING.to_string()),
                            v: Some("0".to_string()),
                            f: None,
                            is: None,
                        },
                        Cell {
                            r: "B1".to_string(),
                            s: None,
                            t: None,
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
        assert_eq!(parsed.sheet_data.rows[0].cells[0].t, Some("s".to_string()));
        assert_eq!(parsed.sheet_data.rows[0].cells[0].v, Some("0".to_string()));
        assert_eq!(parsed.sheet_data.rows[0].cells[1].r, "B1");
        assert_eq!(parsed.sheet_data.rows[0].cells[1].v, Some("42".to_string()));
    }

    #[test]
    fn test_cell_with_formula() {
        let cell = Cell {
            r: "C1".to_string(),
            s: None,
            t: None,
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
            r: "A1".to_string(),
            s: None,
            t: Some(cell_types::INLINE_STRING.to_string()),
            v: None,
            f: None,
            is: Some(InlineString {
                t: Some("Hello World".to_string()),
            }),
        };
        let xml = quick_xml::se::to_string(&cell).unwrap();
        assert!(xml.contains("Hello World"));
        let parsed: Cell = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.t, Some("inlineStr".to_string()));
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
        assert_eq!(parsed.sheet_data.rows[0].cells[0].t, Some("s".to_string()));
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
    fn test_cell_types_constants() {
        assert_eq!(cell_types::BOOLEAN, "b");
        assert_eq!(cell_types::DATE, "d");
        assert_eq!(cell_types::ERROR, "e");
        assert_eq!(cell_types::INLINE_STRING, "inlineStr");
        assert_eq!(cell_types::NUMBER, "n");
        assert_eq!(cell_types::SHARED_STRING, "s");
        assert_eq!(cell_types::FORMULA_STRING, "str");
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
}
